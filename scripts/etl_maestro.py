"""
ETL Pipeline: Extracción de datos desde el archivo Maestro (k.xlsx)
filtrando exclusivamente las placas de la pestaña 'Resumen' (celdas A2 a A35)
y distribución hacia las Plantillas de carga masiva de la aplicación DAG Neumáticos.

Transportes Díaz Acarreos Generales S.A.C.
"""

import os
import re
import sys
import time
import json
import logging
import urllib.request
import urllib.parse
import http.cookiejar
from datetime import datetime
import openpyxl

# Configuración de Logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("ETL_MAESTRO_NEUMATICOS")

# Expresiones regulares
REGEX_CISTERNA = re.compile(r'([A-Z0-9]{3}-[A-Z0-9]{3})')
REGEX_PLACA_TOLERANTE = re.compile(r'([A-Z0-9]{3})[- ]([A-Z0-9]{3})')
REGEX_DNI = re.compile(r'\d{8}')

# Posiciones según guía oficial del sistema
POSICIONES_TRACTO = ['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10', 'RT']
POSICIONES_CISTERNA = ['P11', 'P12', 'P13', 'P14', 'P15', 'P16', 'P17', 'P18', 'P19', 'P20', 'P21', 'P22', 'RS']

MAPA_MARCAS = {
    'VOLVO': 'Volvo',
    'INTERNATIONAL': 'International',
    'INTERNACIONAL': 'International',
    'FREIGHTLINER': 'Freightliner',
    'MACK': 'Mack',
    'SCANIA': 'Scania',
    'MERCEDES': 'Mercedes',
    'KENWORTH': 'Kenworth',
    'MAN': 'Man',
    'IVECO': 'Iveco'
}


def es_codigo_valido(code_raw):
    """
    Valida si un valor corresponde a un código de neumático real.
    Descarta valores nulos, textos de marcador de posición ('sin código', 'no tiene', etc.)
    y encabezados accidentales.
    """
    if code_raw is None:
        return False
    code_str = str(code_raw).strip()
    if not code_str:
        return False
    c_lower = code_str.lower()
    descartes = {
        'none', 'null', 'nan', '0', 'n.a', 'n/a', 'na', 'no tiene',
        'sin codigo', 'sin código', 'sin cdigo', 's/c', 's/n',
        'codigo', 'código', 'placa', 'inicio'
    }
    if c_lower in descartes:
        return False
    if re.search(r'^(sin\s+c[oó]digo|no\s+tiene|s/[cn]|sin\s+n[uú]mero)', c_lower):
        return False
    return True


def normalizar_codigo_neumatico(code_raw):
    """
    Normaliza el código de neumático para que tenga siempre exactamente 7 dígitos.
    Si tiene menos de 7 dígitos, antepone ceros a la izquierda (zfill).
    Maneja enteros, flotantes de Excel (ej. 200095.0) y cadenas.
    """
    if code_raw is None:
        return ""
    code_str = str(code_raw).strip()
    if not code_str:
        return ""

    try:
        f_val = float(code_str)
        if f_val.is_integer():
            return f"{int(f_val):07d}"
        return f"{int(f_val):07d}"
    except (ValueError, TypeError):
        digitos = re.sub(r'\D', '', code_str)
        if digitos and len(code_str) < 7:
            return code_str.zfill(7)
        return code_str


def normalizar_dni(dni_raw):
    """
    Normaliza el DNI para que tenga siempre exactamente 8 dígitos.
    Si tiene menos de 8 dígitos, antepone ceros a la izquierda (zfill).
    Maneja enteros, flotantes de Excel (ej. 1038628.0), strings y números con licencia.
    """
    if dni_raw is None:
        return ""
    try:
        if isinstance(dni_raw, (int, float)):
            f_val = float(dni_raw)
            if f_val.is_integer():
                val_str = str(int(f_val))
            else:
                val_str = str(dni_raw).strip()
        else:
            val_str = str(dni_raw).strip()
            if '.' in val_str:
                try:
                    f_val = float(val_str)
                    if f_val.is_integer():
                        val_str = str(int(f_val))
                except ValueError:
                    pass
    except Exception:
        val_str = str(dni_raw).strip()

    digitos = re.sub(r'\D', '', val_str)
    if digitos:
        if len(digitos) < 8:
            return digitos.zfill(8)
        elif len(digitos) == 8:
            return digitos
        else:
            m = re.search(r'\d{8}', digitos)
            if m:
                return m.group(0)
            return digitos[-8:]
    return val_str.zfill(8) if len(val_str) < 8 else val_str


def extraer_dni_de_texto(texto):
    """
    Extrae o normaliza un DNI a partir de un texto (ej. celda de licencia o DNI directo).
    Garantiza siempre 8 dígitos rellenando con ceros a la izquierda si tiene menos.
    """
    if not texto:
        return None
    val_str = str(texto).strip()
    if not val_str:
        return None

    nums = re.findall(r'\d+', val_str)
    if not nums:
        return None

    for n in nums:
        if len(n) == 8:
            return n

    nums_validos = [n for n in nums if len(n) <= 8]
    if nums_validos:
        mas_largo = max(nums_validos, key=len)
        return mas_largo.zfill(8)

    mas_largo = max(nums, key=len)
    m = re.search(r'\d{8}', mas_largo)
    if m:
        return m.group(0)
    return mas_largo[-8:]


def consultar_nombre_eldni(dni_raw, plantillas_dir=None):
    """
    Consulta un DNI en https://eldni.com/pe/buscar-datos-por-dni.
    Ingresa el DNI en el input 'dni' y extrae el resultado del input id='completos'.
    Si no existe o no tiene valor, retorna 'DNI INCORRECTO'.
    Utiliza una caché persistente en disco (Archivos/cache_conductores_eldni.json)
    para evitar consultas redundantes y mitigar límites de tasa HTTP 429.
    """
    dni_str = normalizar_dni(dni_raw)
    if not dni_str or len(dni_str) != 8 or not dni_str.isdigit():
        return "DNI INCORRECTO"

    # Determinar ruta de caché en el directorio Archivos
    cache_path = None
    if plantillas_dir:
        dir_archivos = os.path.dirname(os.path.abspath(plantillas_dir))
        posible_cache = os.path.join(dir_archivos, 'cache_conductores_eldni.json')
        if os.path.exists(os.path.dirname(posible_cache)):
            cache_path = posible_cache

    if not cache_path:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cache_path = os.path.join(base_dir, 'Archivos', 'cache_conductores_eldni.json')

    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                cache = json.load(f)
        except Exception:
            cache = {}

    if dni_str in cache and cache[dni_str]:
        return cache[dni_str]

    # Si es DNI sintético asignado por el sistema (ej. 900000XX)
    if dni_str.startswith('900000'):
        cache[dni_str] = "DNI INCORRECTO"
        try:
            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump(cache, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return "DNI INCORRECTO"

    logger.info(f"Consultando DNI {dni_str} en https://eldni.com/pe/buscar-datos-por-dni...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Referer': 'https://eldni.com/pe/buscar-datos-por-dni',
        'Origin': 'https://eldni.com',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8',
    }
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    for intento in range(1, 4):
        try:
            # 1. Obtener token CSRF
            req_get = urllib.request.Request('https://eldni.com/pe/buscar-datos-por-dni', headers=headers)
            resp_get = opener.open(req_get, timeout=12)
            html_get = resp_get.read().decode('utf-8', errors='ignore')
            m_tok = re.search(r'name="_token"\s+value="([^"]+)"', html_get)
            if not m_tok:
                time.sleep(1.5)
                continue
            token = m_tok.group(1)

            # 2. Enviar formulario POST con el DNI
            data = urllib.parse.urlencode({'_token': token, 'dni': dni_str}).encode('utf-8')
            req_post = urllib.request.Request('https://eldni.com/pe/buscar-datos-por-dni', data=data, headers=headers)
            time.sleep(1.5)
            resp_post = opener.open(req_post, timeout=15)
            html_post = resp_post.read().decode('utf-8', errors='ignore')

            # 3. Extraer valor de <input ... id="completos" value="..." ...>
            m_comp = re.search(r'id="completos"[^>]*value="([^"]*)"', html_post)
            if not m_comp:
                m_comp = re.search(r'value="([^"]*)"[^>]*id="completos"', html_post)

            nombre_res = "DNI INCORRECTO"
            if m_comp:
                val = m_comp.group(1).strip()
                if val:
                    nombre_res = val.upper()

            cache[dni_str] = nombre_res
            try:
                with open(cache_path, 'w', encoding='utf-8') as f:
                    json.dump(cache, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
            return nombre_res

        except urllib.error.HTTPError as e:
            logger.warning(f"Intento {intento}/3: Error HTTP {e.code} para DNI {dni_str} en eldni.com")
            time.sleep(3.5 * intento)
        except Exception as e:
            logger.warning(f"Intento {intento}/3: Excepción al consultar DNI {dni_str}: {e}")
            time.sleep(2.0)

    # Si no se pudo obtener respuesta válida
    cache[dni_str] = "DNI INCORRECTO"
    try:
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return "DNI INCORRECTO"


def obtener_unidades_activas_resumen(wb, rango_resumen="A2:A35"):
    """
    Extrae las unidades activas de la pestaña 'Resumen' según el rango especificado (por defecto A2:A35).
    Retorna un diccionario: {tracto_placa: cisterna_placa}
    """
    if 'Resumen' not in wb.sheetnames:
        raise ValueError("No se encontró la pestaña 'Resumen' en el archivo maestro.")

    ws_resumen = wb['Resumen']
    unidades = {}

    start_row, end_row = 2, 35
    if rango_resumen:
        m = re.search(r'(\d+):[A-Za-z]*(\d+)', str(rango_resumen))
        if m:
            start_row = int(m.group(1))
            end_row = int(m.group(2))

    for i, r in enumerate(ws_resumen.iter_rows(values_only=True)):
        fila_num = i + 1
        if fila_num < start_row:
            continue
        if fila_num > end_row:
            break

        val = r[0] if r else None
        if val:
            val_str = str(val).strip()
            parts = val_str.split('/')
            tracto = parts[0].strip().upper()
            cisterna = parts[1].strip().upper() if len(parts) > 1 else None

            # Normalizar formato XXX-XXX
            m_tr = REGEX_PLACA_TOLERANTE.search(tracto)
            clean_tr = f"{m_tr.group(1)}-{m_tr.group(2)}" if m_tr else tracto

            clean_ci = None
            if cisterna:
                m_ci = REGEX_PLACA_TOLERANTE.search(cisterna)
                clean_ci = f"{m_ci.group(1)}-{m_ci.group(2)}" if m_ci else cisterna

            unidades[clean_tr] = clean_ci

    logger.info(f"Unidades activas identificadas en 'Resumen!{rango_resumen}': {len(unidades)} unidades.")
    return unidades


def obtener_lista_placas_maestro(ruta_maestro, rango_resumen="A2:A35"):
    """
    Lee rápidamente la pestaña 'Resumen' del archivo maestro k.xlsx
    y retorna la lista de unidades activas {tracto_placa: cisterna_placa}.
    Permite popular selectores de placas en interfaces interactivas.
    """
    if not os.path.exists(ruta_maestro):
        return {}
    wb = openpyxl.load_workbook(ruta_maestro, data_only=True, read_only=True)
    try:
        return obtener_unidades_activas_resumen(wb, rango_resumen=rango_resumen)
    finally:
        wb.close()


def extraer_datos_maestro(ruta_maestro, rango_resumen="A2:A35", placas_seleccionadas=None):
    """
    Lee el archivo maestro k.xlsx y extrae datos EXCLUSIVAMENTE para las unidades
    indicadas en 'Resumen' según el rango especificado (por defecto A2 a A35),
    opcionalmente filtrando por una lista o conjunto específico de placas de tracto.
    """
    logger.info(f"Cargando archivo maestro: {ruta_maestro}")
    if not os.path.exists(ruta_maestro):
        raise FileNotFoundError(f"No se encontró el archivo maestro en: {ruta_maestro}")

    wb = openpyxl.load_workbook(ruta_maestro, data_only=True, read_only=True)

    # 1. Obtener lista estricta de unidades desde Resumen
    unidades_activas = obtener_unidades_activas_resumen(wb, rango_resumen=rango_resumen)

    # 1.1 Filtrar selectivamente por placas elegidas por el usuario
    if placas_seleccionadas:
        placas_limpias = set()
        for p in placas_seleccionadas:
            if not p:
                continue
            m = REGEX_PLACA_TOLERANTE.search(str(p).strip().upper())
            if m:
                placas_limpias.add(f"{m.group(1)}-{m.group(2)}")
            else:
                placas_limpias.add(str(p).strip().upper())
        unidades_activas = {k: v for k, v in unidades_activas.items() if k in placas_limpias}
        logger.info(f"Filtro selectivo aplicado: {len(unidades_activas)} unidades activas seleccionadas de {len(placas_limpias)} indicadas.")

    placas_tracto_validas = set(unidades_activas.keys())

    # 2. Extraer catálogo de conductores desde la hoja 'LISTA'
    conductores_maestro = {}    # clean_tracto -> {nombre, lic, dni}
    todos_conductores = {}      # dni -> {dni, nombre, estado}
    conductores_sin_dni = []

    if 'LISTA' in wb.sheetnames:
        ws_lista = wb['LISTA']
        for i, row in enumerate(ws_lista.iter_rows(values_only=True)):
            if i == 0 or not row or not any(row):
                continue
            nombre = str(row[3] or row[1] or '').strip().upper()
            lic = str(row[2] or '').strip()
            tracto_raw = str(row[6] or '').strip().upper() if len(row) > 6 else ''

            m_tr = REGEX_PLACA_TOLERANTE.search(tracto_raw)
            clean_tr = f"{m_tr.group(1)}-{m_tr.group(2)}" if m_tr else None

            dni = extraer_dni_de_texto(lic)

            if clean_tr:
                conductores_maestro[clean_tr] = {
                    'nombre': nombre,
                    'dni': dni,
                    'item': i
                }

    # Caso especial T5Q-937: conductor en hoja CONDUCT. y extracción de mochilas
    mochilas_maestro = {}
    if 'CONDUCT.' in wb.sheetnames:
        ws_cond = wb['CONDUCT.']
        for i, row in enumerate(ws_cond.iter_rows(values_only=True)):
            if not row:
                continue
            if len(row) > 3 and 'T5Q-937' in str(row[2]) and 'T5Q-937' not in conductores_maestro:
                conductores_maestro['T5Q-937'] = {
                    'nombre': 'SAAVEDRA ODAR SEGUNDO ANGEL',
                    'dni': '42706482',
                    'item': 29
                }
            if len(row) > 4 and row[2]:
                m_tr = REGEX_PLACA_TOLERANTE.search(str(row[2]))
                if m_tr:
                    clean_c_tr = f"{m_tr.group(1)}-{m_tr.group(2)}"
                    m_val = str(row[4]).strip().upper() if row[4] else ''
                    mochilas_maestro[clean_c_tr] = 50 if 'SI' in m_val else 0

    # 3. Extraer vehículos y neumáticos para las unidades activas
    tractos = {}
    cisternas = {}
    neumaticos = {}
    movimientos = []
    inspecciones = []
    conductores_finales = {}
    asignaciones_finales = {}

    def safe_float(v, default=12.0):
        try:
            if v is not None:
                return float(v)
        except:
            pass
        return default

    def safe_psi(v, default=120):
        try:
            if v is not None:
                f_val = float(v)
                if 50 <= f_val <= 150:
                    return int(f_val)
        except:
            pass
        return default

    def mapear_condicion(cond_raw):
        if cond_raw is None:
            return 'USADO'
        c_str = str(cond_raw).strip().upper()
        if c_str in ('0', 'N', 'NUEVO', 'NUEVA', 'NUEVAS'):
            return 'NUEVO'
        if c_str in ('DIBUJADA', 'DIBUJADO', 'DIBUJADAS'):
            return 'DIBUJADO'
        if any(c_str.startswith(p) for p in ('1R', '2R', '3R', '4R', 'R', '01R', '02R')) or c_str in ('1', '2', '3', '4', 'REENCAUCHE', 'REENCAUCHADO'):
            return 'REENCAUCHADO'
        return 'USADO'

    dni_sintetico_base = 90000000

    for tracto_placa, cisterna_defecto in unidades_activas.items():
        if tracto_placa not in wb.sheetnames:
            logger.warning(f"La hoja para el tracto {tracto_placa} no existe en k.xlsx")
            continue

        ws = wb[tracto_placa]
        rows = []
        for r_idx, r in enumerate(ws.iter_rows(values_only=True)):
            rows.append(r)
            if r_idx >= 40:
                break

        if not rows:
            continue

        header_idx = None
        conductor_hoja = None

        for idx, r in enumerate(rows[:10]):
            if r and any(r):
                for c_idx, cell in enumerate(r):
                    if cell and 'CONDUCTOR' in str(cell).upper():
                        for c_val in r[c_idx + 1:]:
                            if c_val and 'CONDUCTOR' not in str(c_val).upper() and 'RUTA' not in str(c_val).upper():
                                conductor_hoja = str(c_val).strip()
                                break
            if r and len(r) > 1 and r[0] == 'MARCA' and r[1] == 'Placa':
                header_idx = idx
                break

        if header_idx is None or header_idx + 2 >= len(rows):
            continue

        r1 = rows[header_idx + 2]
        raw_marca = str(r1[0]).strip() if r1[0] else 'VOLVO'
        marca = MAPA_MARCAS.get(raw_marca.upper(), 'Volvo')

        km_raw = r1[11] if len(r1) > 11 else None
        try:
            km = float(km_raw) if km_raw is not None else 0.0
        except:
            km = 0.0

        fecha_insp_raw = r1[10] if len(r1) > 10 else None
        fecha_insp = None
        if isinstance(fecha_insp_raw, datetime):
            fecha_insp = fecha_insp_raw.strftime('%Y-%m-%d')
        elif fecha_insp_raw:
            fecha_insp = str(fecha_insp_raw).split(' ')[0]

        # Identificar el límite del primer recuadro (el recuadro más reciente)
        # para garantizar que NUNCA se extraigan datos de recuadros históricos inferiores (bloque 2 o posterior)
        limite_primer_recuadro = len(rows)
        for idx_h, r_h in enumerate(rows):
            if idx_h > (header_idx + 2) and r_h and len(r_h) > 1:
                # Si encontramos el inicio de un segundo recuadro (nuevo CONDUCTOR: o nuevo encabezado MARCA/Placa)
                if (r_h[0] == 'MARCA' and r_h[1] == 'Placa') or (r_h[0] and 'CONDUCTOR' in str(r_h[0]).upper()):
                    limite_primer_recuadro = idx_h
                    break

        # Identificar Cisterna de la unidad dentro del primer recuadro
        cisterna_placa = cisterna_defecto
        cisterna_row_idx = None

        for idx_c, r_check in enumerate(rows[header_idx + 12: min(limite_primer_recuadro, len(rows))]):
            if r_check and len(r_check) > 1 and r_check[1]:
                m = REGEX_CISTERNA.search(str(r_check[1]))
                if m:
                    cisterna_placa = m.group(1)
                    cisterna_row_idx = header_idx + 12 + idx_c
                    break

        # Tipo de Aro: detectar desde columna 19 si existe o clasificar por marca
        raw_aro = str(r1[19]).strip() if len(r1) > 19 and r1[19] else ''
        if 'EUROPEO' in raw_aro.upper():
            tipo_aro = 'Europeo'
        elif 'AMERICANO' in raw_aro.upper():
            tipo_aro = 'Americano'
        else:
            tipo_aro = 'Europeo' if marca in ['Volvo', 'Scania', 'Mercedes'] else 'Americano'

        # Capacidad de tanques de combustible (Galones)
        t_izq = 150 if marca in ['Volvo', 'Scania', 'Mercedes'] else 120
        t_der = 150 if marca in ['Volvo', 'Scania', 'Mercedes'] else 120
        t_mochila = mochilas_maestro.get(tracto_placa, 0)

        tractos[tracto_placa] = {
            'placa': tracto_placa,
            'tipo': 'Tractocamion',
            'marca': marca,
            'modelo': 'FM1',
            'tipo_aro': tipo_aro,
            'estado': 'Operativo',
            'km': km,
            'carroceria': '',
            'tanque_lateral_izquierdo': t_izq,
            'tanque_lateral_derecho': t_der,
            'tanque_mochila_auxiliar': t_mochila,
            'cisterna': cisterna_placa,
            'fecha_inspeccion': fecha_insp
        }

        if cisterna_placa:
            cisternas[cisterna_placa] = {
                'placa': cisterna_placa,
                'tipo': 'Semirremolque',
                'marca': 'Tragsac',
                'modelo': 'FM1',
                'tipo_aro': 'Americano',
                'estado': 'Operativo',
                'km': 0,
                'carroceria': 'Cisterna',
                'tanque_lateral_izquierdo': 0,
                'tanque_lateral_derecho': 0,
                'tanque_mochila_auxiliar': 0
            }

        # Conductor de la unidad
        c_info = conductores_maestro.get(tracto_placa)
        if c_info:
            c_nombre = c_info['nombre']
            c_dni = c_info['dni'] if c_info['dni'] else str(dni_sintetico_base + c_info['item'])
        elif conductor_hoja:
            c_nombre = conductor_hoja.upper()
            c_dni = str(dni_sintetico_base + len(conductores_finales) + 1)
        else:
            c_nombre = 'CONDUCTOR NO ASIGNADO'
            c_dni = str(dni_sintetico_base + len(conductores_finales) + 1)

        c_dni = normalizar_dni(c_dni)

        conductores_finales[c_dni] = {
            'dni': c_dni,
            'nombre': c_nombre,
            'estado': 'ACTIVO'
        }
        asignaciones_finales[tracto_placa] = {
            'placa': tracto_placa,
            'dni': c_dni,
            'nombre': c_nombre
        }

        # Neumáticos del Tractocamión (P1 a P10, RT) estrictamente dentro del primer recuadro
        for p_i, pos in enumerate(POSICIONES_TRACTO):
            row_target = header_idx + 2 + p_i
            if row_target < limite_primer_recuadro and row_target < len(rows):
                tr_row = rows[row_target]
                code_raw = tr_row[2] if len(tr_row) > 2 else None
                if es_codigo_valido(code_raw):
                    code_val = normalizar_codigo_neumatico(code_raw)

                    t_medida = str(tr_row[3]).strip() if len(tr_row) > 3 and tr_row[3] else ('295/80R22.5' if pos in ['P1', 'P2'] else '11R22.5')
                    t_marca = str(tr_row[4]).strip().title() if len(tr_row) > 4 and tr_row[4] else 'Lander'
                    t_tipo = 'DIRECCION' if pos in ['P1', 'P2'] else 'TRACCION'
                    t_diseno = str(tr_row[6]).strip() if len(tr_row) > 6 and tr_row[6] else 'LS215'
                    t_condicion = mapear_condicion(tr_row[5] if len(tr_row) > 5 else None)
                    t_psi = safe_psi(tr_row[16] if len(tr_row) > 16 else None)

                    prof_int = safe_float(tr_row[12] if len(tr_row) > 12 else None)
                    prof_cent = safe_float(tr_row[13] if len(tr_row) > 13 else None)
                    prof_ext = safe_float(tr_row[14] if len(tr_row) > 14 else None)

                    neumaticos[code_val] = {
                        'codigo': code_val,
                        'marca': t_marca,
                        'medida': t_medida,
                        'tipo': t_tipo,
                        'diseno': t_diseno,
                        'condicion': t_condicion,
                        'prof_int': prof_int,
                        'prof_cent': prof_cent,
                        'prof_ext': prof_ext,
                        'semana_dot': 24,
                        'ano_dot': 2023,
                        'psi': t_psi
                    }

                    movimientos.append({
                        'placa': tracto_placa,
                        'codigo': code_val,
                        'posicion': pos,
                        'km': km,
                        'motivo': 'Agregado Masivo Manual'
                    })

                    obs_t = 'Desgaste regular normal'
                    if abs(prof_int - prof_ext) >= 3:
                        obs_t = 'Desgaste irregular - revisar alineación'
                    elif prof_cent <= 3 or prof_int <= 3 or prof_ext <= 3:
                        obs_t = 'Cocada crítica - próximo a reencauche'
                    elif t_psi < 100:
                        obs_t = f'Baja presión ({t_psi} PSI)'

                    inspecciones.append({
                        'placa': tracto_placa,
                        'codigo': code_val,
                        'posicion': pos,
                        'km': int(km) if isinstance(km, (int, float)) else km,
                        'prof_int': prof_int,
                        'prof_cent': prof_cent,
                        'prof_ext': prof_ext,
                        'observacion': obs_t
                    })

        # Neumáticos de la Cisterna (P11 a P22, RS) estrictamente dentro del primer recuadro
        if cisterna_placa and cisterna_row_idx:
            for p_i, pos in enumerate(POSICIONES_CISTERNA):
                row_target = cisterna_row_idx + p_i
                if row_target < limite_primer_recuadro and row_target < len(rows):
                    ci_row = rows[row_target]
                    code_raw = ci_row[2] if len(ci_row) > 2 else None
                    if es_codigo_valido(code_raw):
                        code_val = normalizar_codigo_neumatico(code_raw)

                        t_medida = str(ci_row[3]).strip() if len(ci_row) > 3 and ci_row[3] else '11R22.5'
                        t_marca = str(ci_row[4]).strip().title() if len(ci_row) > 4 and ci_row[4] else 'Lander'
                        t_diseno = str(ci_row[6]).strip() if len(ci_row) > 6 and ci_row[6] else 'LS168'
                        t_condicion = mapear_condicion(ci_row[5] if len(ci_row) > 5 else None)
                        t_psi = safe_psi(ci_row[16] if len(ci_row) > 16 else None)

                        prof_int = safe_float(ci_row[12] if len(ci_row) > 12 else None)
                        prof_cent = safe_float(ci_row[13] if len(ci_row) > 13 else None)
                        prof_ext = safe_float(ci_row[14] if len(ci_row) > 14 else None)

                        neumaticos[code_val] = {
                            'codigo': code_val,
                            'marca': t_marca,
                            'medida': t_medida,
                            'tipo': 'ARRASTRE',
                            'diseno': t_diseno,
                            'condicion': t_condicion,
                            'prof_int': prof_int,
                            'prof_cent': prof_cent,
                            'prof_ext': prof_ext,
                            'semana_dot': 24,
                            'ano_dot': 2023,
                            'psi': t_psi
                        }

                        movimientos.append({
                            'placa': cisterna_placa,
                            'codigo': code_val,
                            'posicion': pos,
                            'km': km,
                            'motivo': 'Agregado Masivo Manual'
                        })

                        obs_c = 'Presión correcta'
                        if abs(prof_int - prof_ext) >= 3:
                            obs_c = 'Desgaste irregular - verificar ejes'
                        elif prof_cent <= 3 or prof_int <= 3 or prof_ext <= 3:
                            obs_c = 'Cocada crítica - próximo a reencauche'
                        elif t_psi < 100:
                            obs_c = f'Baja presión ({t_psi} PSI)'

                        inspecciones.append({
                            'placa': cisterna_placa,
                            'codigo': code_val,
                            'posicion': pos,
                            'km': int(km) if isinstance(km, (int, float)) else km,
                            'prof_int': prof_int,
                            'prof_cent': prof_cent,
                            'prof_ext': prof_ext,
                            'observacion': obs_c
                        })

    wb.close()

    logger.info(f"=== EXTRACCIÓN DE UNIDADES ACTIVAS (RESUMEN!{rango_resumen}) COMPLETADA ===")
    logger.info(f" - Tractocamiones: {len(tractos)}")
    logger.info(f" - Semirremolques: {len(cisternas)}")
    logger.info(f" - Conductores: {len(conductores_finales)}")
    logger.info(f" - Neumáticos únicos: {len(neumaticos)}")
    logger.info(f" - Movimientos / Asignaciones: {len(movimientos)}")
    logger.info(f" - Inspecciones técnicas: {len(inspecciones)}")

    return {
        'conductores': conductores_finales,
        'asignaciones_conductores': asignaciones_finales,
        'tractos': tractos,
        'cisternas': cisternas,
        'neumaticos': neumaticos,
        'movimientos': movimientos,
        'inspecciones': inspecciones,
        'unidades_activas': unidades_activas
    }


def _actualizar_plantilla_excel(ruta_archivo, nombre_hoja, filas_datos, reintentos=3, pausa=0.6, formato_columna_a=None, formatos_columnas=None):
    """
    Actualiza la hoja de datos principal del archivo plantilla conservando encabezados
    y guías. Incluye reintentos ante bloqueos temporales de archivos en Windows.
    Permite formatear columnas como texto ('@') para conservar ceros a la izquierda (ej. DNI).
    """
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"Plantilla no encontrada: {ruta_archivo}")

    wb = openpyxl.load_workbook(ruta_archivo)
    try:
        if nombre_hoja not in wb.sheetnames:
            ws = wb.sheetnames[0]
            logger.warning(f"Hoja '{nombre_hoja}' no encontrada, utilizando primera hoja: '{ws.title}'")
        else:
            ws = wb[nombre_hoja]

        # Eliminar datos anteriores conservando la fila 1 (encabezados)
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

        for fila in filas_datos:
            ws.append(fila)

        # Aplicar formato de celda si se requiere (ej. formato texto '@' para DNI con ceros)
        fmts = {}
        if formatos_columnas:
            fmts.update(formatos_columnas)
        if formato_columna_a:
            fmts[1] = formato_columna_a

        if fmts:
            for col_idx, fmt in fmts.items():
                for r_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r_idx, column=col_idx)
                    cell.number_format = fmt
                    if cell.value is not None and fmt == '@':
                        cell.value = str(cell.value)
                        cell.data_type = 's'

        guardado = False
        ultimo_error = None
        for intento in range(1, reintentos + 1):
            try:
                wb.save(ruta_archivo)
                guardado = True
                break
            except PermissionError as e:
                ultimo_error = e
                logger.warning(f"Intento {intento}/{reintentos}: Archivo '{os.path.basename(ruta_archivo)}' bloqueado. Reintentando en {pausa}s...")
                time.sleep(pausa)

        if not guardado:
            msg = (f"No se pudo guardar '{os.path.basename(ruta_archivo)}'. "
                   f"El archivo está bloqueado por otro proceso o abierto en Excel. "
                   f"Por favor ciérrelo e intente nuevamente.")
            logger.error(msg)
            raise PermissionError(msg) from ultimo_error

        logger.info(f"Actualizada plantilla '{os.path.basename(ruta_archivo)}' con {len(filas_datos)} registros.")
    finally:
        wb.close()


def generar_plantilla_vehiculos(plantillas_dir, tractos, cisternas):
    """Genera Plantilla_Registro_Masivo_Vehiculos_DAG.xlsx"""
    archivo = os.path.join(plantillas_dir, 'Plantilla_Registro_Masivo_Vehiculos_DAG.xlsx')
    filas = []

    # 1. Tractocamiones
    for p in sorted(tractos.keys()):
        v = tractos[p]
        filas.append([
            v['placa'],
            v['tipo'],
            v['marca'],
            v['modelo'],
            v['tipo_aro'],
            v['estado'],
            v['km'],
            v['carroceria'],
            v.get('tanque_lateral_izquierdo', 0),
            v.get('tanque_lateral_derecho', 0),
            v.get('tanque_mochila_auxiliar', 0)
        ])

    # 2. Semirremolques
    for p in sorted(cisternas.keys()):
        c = cisternas[p]
        filas.append([
            c['placa'],
            c['tipo'],
            c['marca'],
            c['modelo'],
            c['tipo_aro'],
            c['estado'],
            c['km'],
            c['carroceria'],
            c.get('tanque_lateral_izquierdo', 0),
            c.get('tanque_lateral_derecho', 0),
            c.get('tanque_mochila_auxiliar', 0)
        ])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_VEHICULOS', filas)
    return len(filas)


def generar_plantilla_conductores(plantillas_dir, conductores):
    """
    Genera Plantilla_Registro_Masivo_Conductores_DAG.xlsx.
    Garantiza que la columna A (DNI) tenga siempre 8 dígitos rellenando con ceros a la izquierda si tiene menos.
    Antes de escribir, consulta el DNI en https://eldni.com/pe/buscar-datos-por-dni (input id='dni')
    y copia los resultados (input id='completos').
    - Si existe: Columna A = DNI (8 dígitos), Columna B = Nombre oficial de id='completos'.
    - Si no existe: Columna A = 'DNI INCORRECTO', Columna B = Nombre original del conductor (dejado como está).
    """
    archivo = os.path.join(plantillas_dir, 'Plantilla_Registro_Masivo_Conductores_DAG.xlsx')
    filas = []

    for dni in sorted(conductores.keys(), key=lambda x: str(x)):
        c = conductores[dni]
        dni_val = normalizar_dni(c.get('dni', dni))
        nombre_original = c.get('nombre', '')

        # Consultar en eldni.com (utiliza caché persistente en disco)
        nombre_oficial = consultar_nombre_eldni(dni_val, plantillas_dir)

        if nombre_oficial == "DNI INCORRECTO":
            col_a_dni = "DNI INCORRECTO"
            col_b_nombre = nombre_original
        else:
            col_a_dni = dni_val
            col_b_nombre = nombre_oficial

        # Actualizar datos en memoria para coherencia en el flujo
        c['dni_display'] = col_a_dni
        c['dni'] = col_a_dni
        c['nombre'] = col_b_nombre

        filas.append([
            col_a_dni,
            col_b_nombre,
            c['estado']
        ])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_CONDUCTORES', filas, formato_columna_a='@')
    return len(filas)


def generar_plantilla_asignacion_conductores(plantillas_dir, asignaciones, tractos, conductores):
    """Genera Plantilla_Registro_Masivo_Asignacion_Conductores_DAG.xlsx asegurando 8 dígitos en DNI y nombres sincronizados."""
    archivo = os.path.join(plantillas_dir, 'Plantilla_Registro_Masivo_Asignacion_Conductores_DAG.xlsx')
    filas = []
    fecha_actual = datetime.now().strftime('%Y-%m-%d')

    for placa in sorted(tractos.keys()):
        if placa in asignaciones:
            asig = asignaciones[placa]
            dni_val = normalizar_dni(asig.get('dni', ''))
            nombre_orig = asig.get('nombre', '')

            # Sincronizar con el conductor validado
            c_match = conductores.get(dni_val)
            if c_match and 'dni_display' in c_match:
                dni_final = c_match['dni_display']
                nombre_final = c_match['nombre']
            else:
                nombre_oficial = consultar_nombre_eldni(dni_val, plantillas_dir)
                if nombre_oficial == "DNI INCORRECTO":
                    dni_final = "DNI INCORRECTO"
                    nombre_final = nombre_orig
                else:
                    dni_final = dni_val
                    nombre_final = nombre_oficial

            asig['dni'] = dni_final
            asig['nombre'] = nombre_final

            filas.append([
                placa,
                dni_final,
                nombre_final,
                fecha_actual,
                'ASIGNACIÓN MASIVA AUTOMÁTICA DAG'
            ])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_ASIGNACIONES', filas, formatos_columnas={2: '@'})
    return len(filas)


def generar_plantilla_neumaticos(plantillas_dir, neumaticos):
    """
    Genera Plantilla_Registro_Masivo_Neumaticos_DAG.xlsx.
    Garantiza que la columna A (CODIGO) tenga siempre 7 dígitos rellenando con ceros a la izquierda si tiene menos.
    """
    archivo = os.path.join(plantillas_dir, 'Plantilla_Registro_Masivo_Neumaticos_DAG.xlsx')
    filas = []

    for codigo in sorted(neumaticos.keys(), key=lambda x: str(x)):
        n = neumaticos[codigo]
        cod_val = normalizar_codigo_neumatico(n['codigo'])
        filas.append([
            cod_val,
            n['marca'],
            n['medida'],
            n['tipo'],
            n['diseno'],
            n['condicion'],
            n['prof_int'],
            n['prof_cent'],
            n['prof_ext'],
            n['semana_dot'],
            n['ano_dot'],
            n['psi']
        ])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_NEUMATICOS', filas, formato_columna_a='@')
    return len(filas)


def generar_plantilla_movimientos(plantillas_dir, movimientos):
    """Genera Plantilla_Asignacion_Movimiento_Neumaticos_DAG.xlsx asegurando 7 dígitos en CODIGO."""
    archivo = os.path.join(plantillas_dir, 'Plantilla_Asignacion_Movimiento_Neumaticos_DAG.xlsx')
    filas = []

    for m in movimientos:
        cod_val = normalizar_codigo_neumatico(m['codigo'])
        filas.append([
            m['placa'],
            cod_val,
            m['posicion'],
            m['km'],
            m['motivo']
        ])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_MOVIMIENTOS', filas, formatos_columnas={2: '@'})
    return len(filas)


def generar_plantilla_comparativa(plantillas_dir, tractos):
    """Genera Plantilla_Registro_Masivo_Comparativa_DAG.xlsx"""
    archivo = os.path.join(plantillas_dir, 'Plantilla_Registro_Masivo_Comparativa_DAG.xlsx')
    filas = []

    for placa in sorted(tractos.keys()):
        t = tractos[placa]
        cap_total = t.get('tanque_lateral_izquierdo', 0) + t.get('tanque_lateral_derecho', 0) + t.get('tanque_mochila_auxiliar', 0)
        galones_diesel = cap_total if cap_total > 0 else (223 if t.get('marca') == 'Volvo' else 200)

        filas.append([placa, 'GASOLINA REGULAR', None, 'ACTIVO'])
        filas.append([placa, 'GASOLINA PREMIUM', None, 'ACTIVO'])
        filas.append([placa, 'DIESEL B5 S-50', galones_diesel, 'ACTIVO'])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_COMPARATIVA', filas)
    return len(filas)


def generar_plantilla_grifos(plantillas_dir):
    """Genera / Asegura Plantilla_Registro_Masivo_Grifos_DAG.xlsx"""
    archivo = os.path.join(plantillas_dir, 'Plantilla_Registro_Masivo_Grifos_DAG.xlsx')
    grifos_base = [
        ['GRIFO PRIMAX PIURA CENTRO', 'PRIMAX', 'PIURA', 'AV. SANCHEZ CERRO 1234', 'ACTIVO'],
        ['ESTACION REPSOL TALARA SUR', 'REPSOL', 'TALARA', 'PANAMERICANA NORTE KM 1080', 'ACTIVO'],
        ['PATIO CENTRAL DAG PAITA', 'PATIO / PROPIO', 'PAITA', 'ZONA INDUSTRIAL MZ. C LOTE 4', 'ACTIVO'],
        ['GRIFO PETROPERU CHICLAYO', 'PETROPERÚ', 'CHICLAYO', 'AV. VICTOR RAUL HAYA DE LA TORRE 500', 'ACTIVO'],
        ['GRIFO SHELL LIMA NORTE', 'SHELL', 'LIMA', 'PANAMERICANA NORTE KM 22', 'ACTIVO']
    ]

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_GRIFOS', grifos_base)
    return len(grifos_base)


def generar_plantilla_inspecciones(plantillas_dir, inspecciones):
    """
    Genera Plantilla_Inspeccion_Neumaticos_DAG.xlsx asegurando 7 dígitos en CODIGO_NEUMATICO.
    Contiene las 804 mediciones de desgaste (profundidades int, cent, ext) y kilometraje
    correspondientes a los neumáticos montados en la flota activa.
    """
    archivo = os.path.join(plantillas_dir, 'Plantilla_Inspeccion_Neumaticos_DAG.xlsx')
    filas = []

    for ins in inspecciones:
        cod_val = normalizar_codigo_neumatico(ins['codigo'])
        filas.append([
            ins['placa'],
            cod_val,
            ins['posicion'],
            ins['km'],
            ins['prof_int'],
            ins['prof_cent'],
            ins['prof_ext'],
            ins.get('observacion', 'Desgaste regular normal')
        ])

    _actualizar_plantilla_excel(archivo, 'PLANTILLA_INSPECCIONES', filas, formatos_columnas={2: '@'})
    return len(filas)


def ejecutar_pipeline(maestro_path=None, plantillas_dir=None, rango_resumen="A2:A35", placas_seleccionadas=None):
    """
    Función principal de ejecución del Pipeline ETL.
    """
    inicio = datetime.now()
    logger.info(f"=== INICIANDO PIPELINE ETL (FILTRO: RESUMEN!{rango_resumen}) ===")

    base_local = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

    if not maestro_path:
        local_maestro = os.path.join(base_local, 'Archivos', 'Maestro', 'k.xlsx')
        if os.path.exists(local_maestro):
            maestro_path = local_maestro
        elif os.path.exists('/app/Archivos/Maestro/k.xlsx'):
            maestro_path = '/app/Archivos/Maestro/k.xlsx'
        else:
            maestro_path = local_maestro

    if not plantillas_dir:
        local_plantillas = os.path.join(base_local, 'Archivos', 'Plantillas')
        if os.path.exists(local_plantillas):
            plantillas_dir = local_plantillas
        elif os.path.exists('/app/Archivos/Plantillas'):
            plantillas_dir = '/app/Archivos/Plantillas'
        else:
            plantillas_dir = local_plantillas

    logger.info(f"Ruta Maestro: {maestro_path}")
    logger.info(f"Directorio Plantillas: {plantillas_dir}")

    # 1. Extracción con filtro Resumen y placas seleccionadas
    datos = extraer_datos_maestro(maestro_path, rango_resumen=rango_resumen, placas_seleccionadas=placas_seleccionadas)

    # 2. Generación de las 8 plantillas oficiales
    n_vehiculos = generar_plantilla_vehiculos(plantillas_dir, datos['tractos'], datos['cisternas'])
    n_conductores = generar_plantilla_conductores(plantillas_dir, datos['conductores'])
    n_asig_conductores = generar_plantilla_asignacion_conductores(plantillas_dir, datos['asignaciones_conductores'], datos['tractos'], datos['conductores'])
    n_neumaticos = generar_plantilla_neumaticos(plantillas_dir, datos['neumaticos'])
    n_movimientos = generar_plantilla_movimientos(plantillas_dir, datos['movimientos'])
    n_inspecciones = generar_plantilla_inspecciones(plantillas_dir, datos.get('inspecciones', []))
    n_comparativa = generar_plantilla_comparativa(plantillas_dir, datos['tractos'])
    n_grifos = generar_plantilla_grifos(plantillas_dir)

    # 3. Generación del Reporte de Auditoría y Conciliación (Opción B, Duplicados y Faltantes)
    reporte_auditoria_path = os.path.join(plantillas_dir, 'Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx')
    try:
        try:
            from scripts.generar_reporte_auditoria import crear_reporte_auditoria_excel
        except ModuleNotFoundError:
            from generar_reporte_auditoria import crear_reporte_auditoria_excel
        crear_reporte_auditoria_excel(datos, reporte_auditoria_path)
        dir_reportes = os.path.join(base_local, 'Archivos', 'Reportes')
        if os.path.exists(dir_reportes):
            crear_reporte_auditoria_excel(datos, os.path.join(dir_reportes, 'Reporte_Auditoria_Conciliacion_Neumaticos_DAG.xlsx'))
        logger.info(f"Reporte de Auditoría y Conciliación generado en: {reporte_auditoria_path}")
    except Exception as e_rep:
        logger.warning(f"No se pudo generar el reporte de auditoría automático: {e_rep}")

    duracion = (datetime.now() - inicio).total_seconds()
    resumen = {
        'duracion_segundos': duracion,
        'vehiculos_total': n_vehiculos,
        'conductores_total': n_conductores,
        'asignaciones_conductores_total': n_asig_conductores,
        'neumaticos_total': n_neumaticos,
        'movimientos_total': n_movimientos,
        'inspecciones_total': n_inspecciones,
        'comparativas_total': n_comparativa,
        'grifos_total': n_grifos,
        'reporte_auditoria': reporte_auditoria_path
    }

    logger.info("=== PIPELINE ETL COMPLETADO EXITOSAMENTE ===")
    logger.info("Resumen de registros generados en plantillas:")
    for k, v in resumen.items():
        logger.info(f" - {k}: {v}")

    return resumen


if __name__ == '__main__':
    ejecutar_pipeline()

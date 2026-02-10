import pandas as pd
import json
import os
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors

# ==========================================
#   CONFIGURACIÓN MAESTRA DEL SISTEMA
# ==========================================

# Rutas de archivos (Automáticas según donde esté el script)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(BASE_DIR, "INVENTARIO ACTIVOS FIJOS SURAKI 30-01.xlsx")
ARCHIVO_LOGO  = os.path.join(BASE_DIR, "logo_suraki.png") # Debe existir para salir en el PDF

# Salidas
SALIDA_JSON = os.path.join(BASE_DIR, "datos_inventario.json")   # Para la Web (GitHub)
SALIDA_PDF  = os.path.join(BASE_DIR, "Etiquetas_Suraki_Master.pdf") # Para Imprimir

# Tu Web de GitHub (El QR apuntará aquí)
URL_SISTEMA = "https://sissur-wq.github.io/Inventario_Suraki/"

# Configuración de Etiquetas (57x32mm)
ANCHO_ETIQUETA = 57 * mm
ALTO_ETIQUETA  = 32 * mm

# Hojas que NO son de activos
HOJAS_IGNORAR = ['NOMENCLATURA', 'MODELO DE ETIQUETADO', 'RESUMEN', 'INDICE', 'PORTADA', 'CODIFICACION']

# ==========================================
#   FUNCIONES DE AYUDA (GLOBALES)
# ==========================================

def detectar_fila_encabezados(df_raw):
    """Busca en qué fila están los títulos (DESCRIPCION, CODIGO, etc.)"""
    keywords = ['DESCRIPCION', 'DESCRIPCIÓN', 'MARCA', 'MODELO', 'SERIAL', 'CÓDIGO', 'CODIGO', 'BIEN']
    for i, row in df_raw.iterrows():
        fila_str = " ".join([str(val).upper() for val in row.values])
        if sum(1 for k in keywords if k in fila_str) >= 2: return i
    return None

def buscar_columna(df, posibles_nombres):
    """Encuentra la columna correcta ignorando mayúsculas o espacios"""
    cols_limpias = [str(c).strip().upper() for c in df.columns]
    for candidato in posibles_nombres:
        cand = candidato.upper()
        if cand in cols_limpias: return df.columns[cols_limpias.index(cand)]
        for i, col_real in enumerate(cols_limpias):
            if cand in col_real: return df.columns[i]
    return None

def cargar_nomenclatura(xls):
    """Carga la hoja de Nomenclatura para saber la categoría del activo (Ej: RE = Refrigeración)"""
    mapa = {}
    try:
        hoja_nom = next((s for s in xls.sheet_names if "codif" in s.lower() or "nomenclatura" in s.lower()), None)
        if hoja_nom:
            # Intentamos leer asumiendo fila 1 como header
            df_nom = pd.read_excel(xls, sheet_name=hoja_nom, header=1) 
            col_cod = next((c for c in df_nom.columns if "dig" in str(c).lower() or "XX" in str(c)), None)
            col_cat = next((c for c in df_nom.columns if "cat" in str(c).lower()), None)
            
            if col_cod and col_cat:
                for _, row in df_nom.iterrows():
                    c = str(row[col_cod]).strip().upper()
                    if c and c != 'NAN': mapa[c] = str(row[col_cat]).strip()
    except: pass
    return mapa

def obtener_categoria(id_val, mapa):
    """Busca qué categoría es según el prefijo del ID (Ej: RE-001 -> Refrigeración)"""
    if not mapa: return ""
    partes = id_val.replace('-', ' ').split()
    for p in partes:
        if p.strip().upper() in mapa: return mapa[p.strip().upper()]
    return ""

def dibujar_texto_dinamico(c, texto, x, y, max_w, font="Helvetica", max_s=8, min_s=4):
    """Reduce el texto para que quepa en el ancho disponible"""
    size = max_s
    c.setFont(font, size)
    while c.stringWidth(texto, font, size) > max_w and size > min_s:
        size -= 0.5
        c.setFont(font, size)
    c.drawString(x, y, texto)

# ==========================================
#   PARTE 1: GENERADOR DE BASE DE DATOS (JSON)
# ==========================================

def generar_base_datos_web(xls):
    print("🔄 Generando Base de Datos para la Web...")
    lista_activos = []

    for hoja in xls.sheet_names:
        if any(x in hoja.upper() for x in HOJAS_IGNORAR): continue
        
        try:
            df_temp = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=15)
            fila = detectar_fila_encabezados(df_temp)
            if fila is None: continue
            
            df = pd.read_excel(xls, sheet_name=hoja, header=fila)
            
            col_id = buscar_columna(df, ['CODIGO', 'ID', 'ETIQUETA'])
            col_desc = buscar_columna(df, ['DESCRIPCION', 'BIEN', 'NOMBRE'])
            col_marca = buscar_columna(df, ['MARCA', 'FABRICANTE'])
            col_modelo = buscar_columna(df, ['MODELO'])
            col_serial = buscar_columna(df, ['SERIAL', 'SERIE', 'S/N'])

            if col_id and col_desc:
                for _, row in df.iterrows():
                    id_val = str(row[col_id]).strip()
                    if id_val.lower() not in ['nan', '', '0', 'none']:
                        activo = {
                            "id": id_val,
                            "descripcion": str(row[col_desc]).strip(),
                            "marca": str(row[col_marca]).replace('nan','') if col_marca else "",
                            "modelo": str(row[col_modelo]).replace('nan','') if col_modelo else "",
                            "serial": str(row[col_serial]).replace('nan','') if col_serial else "S/N",
                            "sede": hoja
                        }
                        lista_activos.append(activo)
        except Exception as e:
            print(f"   ⚠️ Error leyendo hoja '{hoja}': {e}")

    # Guardar JSON
    with open(SALIDA_JSON, 'w', encoding='utf-8') as f:
        json.dump(lista_activos, f, indent=4, ensure_ascii=False)
    print(f"✅ JSON Creado: {len(lista_activos)} activos exportados.")

# ==========================================
#   PARTE 2: GENERADOR DE ETIQUETAS (PDF)
# ==========================================

def generar_pdf_etiquetas(xls):
    print("🖨️ Generando PDF de Etiquetas...")
    
    c = canvas.Canvas(SALIDA_PDF, pagesize=(ANCHO_ETIQUETA, ALTO_ETIQUETA))
    mapa_nomenclatura = cargar_nomenclatura(xls)
    tiene_logo = os.path.exists(ARCHIVO_LOGO)
    
    total = 0

    for hoja in xls.sheet_names:
        if any(x in hoja.upper() for x in HOJAS_IGNORAR): continue
        print(f"   📄 Procesando hoja: {hoja}")

        try:
            df_temp = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=15)
            fila = detectar_fila_encabezados(df_temp)
            if fila is None: continue
            
            df = pd.read_excel(xls, sheet_name=hoja, header=fila)
            col_id = buscar_columna(df, ['CODIGO', 'ID', 'ETIQUETA'])
            col_desc = buscar_columna(df, ['DESCRIPCION', 'BIEN'])
            col_marca = buscar_columna(df, ['MARCA'])
            col_modelo = buscar_columna(df, ['MODELO'])
            col_serial = buscar_columna(df, ['SERIAL', 'S/N'])

            if not col_id or not col_desc: continue

            for _, row in df.iterrows():
                id_val = str(row[col_id]).strip()
                if id_val.lower() in ['nan', '', '0', 'none']: continue
                
                total += 1

                # --- DISEÑO DE ETIQUETA ---
                
                # 1. LOGO Y TÍTULO
                y_cursor = ALTO_ETIQUETA - 2*mm
                if tiene_logo:
                    try:
                        c.drawImage(ARCHIVO_LOGO, 2*mm, ALTO_ETIQUETA - 12*mm, width=12*mm, height=10*mm, mask='auto', preserveAspectRatio=True)
                    except: pass
                
                # Título
                c.setFillColor(colors.black)
                c.setFont("Helvetica-Bold", 10)
                c.drawString(16*mm, ALTO_ETIQUETA - 7*mm, "HIPER SURAKI")
                
                # Categoría (Pequeña debajo del título)
                cat = obtener_categoria(id_val, mapa_nomenclatura)
                if cat:
                    c.setFillColor(colors.darkgrey)
                    c.setFont("Helvetica", 5)
                    c.drawString(16*mm, ALTO_ETIQUETA - 9.5*mm, cat[:25])

                # 2. DATOS DEL ACTIVO (Izquierda)
                x_content = 2*mm
                y_cursor = ALTO_ETIQUETA - 14*mm
                w_content = 36*mm # Espacio antes del QR

                # Descripción
                c.setFillColor(colors.black)
                c.setFont("Helvetica", 7.5)
                desc = str(row[col_desc]).replace('\n', ' ').strip()
                
                # Lógica de lineas para descripción
                palabras = desc.split()
                linea1, linea2 = "", ""
                for p in palabras:
                    if c.stringWidth(linea1 + " " + p, "Helvetica", 7.5) < w_content: linea1 += " " + p
                    else: linea2 += " " + p
                
                c.drawString(x_content, y_cursor, linea1.strip())
                if linea2:
                    y_cursor -= 3*mm
                    c.drawString(x_content, y_cursor, linea2.strip()[:30])

                # Marca / Modelo
                y_cursor -= 4*mm
                marca = str(row[col_marca]).replace('nan','') if col_marca else ""
                modelo = str(row[col_modelo]).replace('nan','') if col_modelo else ""
                txt_mod = f"{marca} / {modelo}".strip(" / ") or "GENÉRICO"
                
                c.setFont("Helvetica-Bold", 7)
                dibujar_texto_dinamico(c, txt_mod, x_content, y_cursor, w_content, "Helvetica-Bold", 7, 5)

                # Serial
                if col_serial:
                    ser = str(row[col_serial]).replace('nan','')
                    if ser:
                        y_cursor -= 3*mm
                        c.setFont("Helvetica", 6)
                        c.drawString(x_content, y_cursor, f"SN: {ser}")

                # 3. QR (Derecha) -> APUNTA A LA WEB
                qr_size = 17 * mm
                qr_x = ANCHO_ETIQUETA - 19*mm
                qr_y = ALTO_ETIQUETA - 22*mm
                
                # CONTENIDO QR = URL WEB
                qr_data = f"{URL_SISTEMA}?id={id_val}"
                
                qr_obj = qr.QrCodeWidget(qr_data)
                qr_obj.barWidth = qr_size
                qr_obj.barHeight = qr_size
                qr_obj.qrVersion = 1
                
                # Marco QR
                c.setLineWidth(0.5)
                c.setStrokeColor(colors.lightgrey)
                c.rect(qr_x - 0.5*mm, qr_y - 0.5*mm, qr_size + 1*mm, qr_size + 1*mm, stroke=1, fill=0)
                
                d = Drawing(qr_size, qr_size)
                d.add(qr_obj)
                d.drawOn(c, qr_x, qr_y)

                # ID Debajo del QR
                c.setFillColor(colors.black)
                center_x = qr_x + (qr_size/2)
                y_id = qr_y - 3*mm
                
                c.setFont("Helvetica-Bold", 8)
                size_id = 8
                while c.stringWidth(id_val, "Helvetica-Bold", size_id) > (qr_size + 4*mm) and size_id > 4:
                    size_id -= 0.5
                    c.setFont("Helvetica-Bold", size_id)
                
                c.drawCentredString(center_x, y_id, id_val)

                c.showPage()
        
        except Exception as e:
            print(f"   ⚠️ Error en hoja '{hoja}': {e}")

    c.save()
    print(f"✅ PDF Creado: {total} etiquetas generadas.")

# ==========================================
#   EJECUCIÓN PRINCIPAL
# ==========================================

if __name__ == "__main__":
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"❌ ERROR: No se encuentra el archivo '{ARCHIVO_EXCEL}'")
    else:
        xls_file = pd.ExcelFile(ARCHIVO_EXCEL)
        
        # 1. Generar JSON para la web
        generar_base_datos_web(xls_file)
        
        # 2. Generar Etiquetas para imprimir
        generar_pdf_etiquetas(xls_file)
        
        print("\n🚀 ¡PROCESO FINALIZADO CON ÉXITO!")
        print(f"1. Sube '{os.path.basename(SALIDA_JSON)}' a tu GitHub.")
        print(f"2. Imprime '{os.path.basename(SALIDA_PDF)}' en tu impresora de etiquetas.")
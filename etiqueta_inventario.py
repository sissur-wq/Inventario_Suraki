import pandas as pd
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
import os

# ==========================================
#   CONFIGURACIÓN (OFFLINE PREMIUM)
# ==========================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(BASE_DIR, "INVENTARIO ACTIVOS FIJOS SURAKI 30-01.xlsx")
ARCHIVO_SALIDA = os.path.join(BASE_DIR, "Etiquetas_Suraki_Premium_Offline.pdf")
ARCHIVO_LOGO = os.path.join(BASE_DIR, "logo_suraki.png") # Opcional, si tienes el logo

# Dimensiones exactas (57x32 mm)
ANCHO = 57 * mm
ALTO = 32 * mm

# Colores Corporativos
ROJO_SURAKI = colors.HexColor("#D32F2F")
NEGRO = colors.black
BLANCO = colors.white
GRIS_OSCURO = colors.HexColor("#333333")

# Hojas a ignorar
IGNORAR = ['NOMENCLATURA', 'MODELO', 'RESUMEN', 'INDICE', 'PORTADA', 'CODIFICACION']

# ==========================================
#   FUNCIONES DE AYUDA
# ==========================================

def detectar_encabezados(df_raw):
    """Busca dónde empieza la tabla real"""
    keywords = ['DESCRIPCION', 'MARCA', 'MODELO', 'SERIAL', 'CODIGO', 'BIEN']
    for i, row in df_raw.iterrows():
        texto = " ".join([str(val).upper() for val in row.values])
        if sum(1 for k in keywords if k in texto) >= 2: return i
    return None

def buscar_columna(df, candidatos):
    """Encuentra la columna correcta"""
    cols = [str(c).strip().upper() for c in df.columns]
    for cand in candidatos:
        cand = cand.upper()
        if cand in cols: return df.columns[cols.index(cand)]
        for i, real in enumerate(cols):
            if cand in real: return df.columns[i]
    return None

def texto_ajustable(c, texto, x, y, max_ancho, fuente="Helvetica", max_sz=8):
    """Reduce la letra si el texto es muy largo"""
    sz = max_sz
    c.setFont(fuente, sz)
    while c.stringWidth(texto, fuente, sz) > max_ancho and sz > 4:
        sz -= 0.5
        c.setFont(fuente, sz)
    c.drawString(x, y, texto)

# ==========================================
#   GENERADOR DE ETIQUETAS
# ==========================================

def generar_offline_premium():
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"❌ ERROR: Falta el archivo '{ARCHIVO_EXCEL}'")
        return

    print("✨ Generando etiquetas 'Premium Offline' (Con diseño visual en el QR)...")
    
    try:
        xls = pd.ExcelFile(ARCHIVO_EXCEL)
    except Exception as e:
        print(f"❌ Error Excel: {e}")
        return

    c = canvas.Canvas(ARCHIVO_SALIDA, pagesize=(ANCHO, ALTO))
    total = 0
    tiene_logo = os.path.exists(ARCHIVO_LOGO)

    for hoja in xls.sheet_names:
        if any(x in hoja.upper() for x in IGNORAR): continue
        
        try:
            # Leer excel inteligentemente
            df_temp = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=15)
            fila = detectar_encabezados(df_temp)
            if fila is None: continue
            
            df = pd.read_excel(xls, sheet_name=hoja, header=fila)
            
            # Buscar columnas
            col_id = buscar_columna(df, ['CODIGO', 'ID', 'ETIQUETA'])
            col_desc = buscar_columna(df, ['DESCRIPCION', 'BIEN', 'NOMBRE'])
            col_marca = buscar_columna(df, ['MARCA'])
            col_modelo = buscar_columna(df, ['MODELO'])
            col_serial = buscar_columna(df, ['SERIAL', 'SERIE', 'S/N'])

            if not col_id or not col_desc: continue
            print(f"   Procesando: {hoja}...")

            for _, row in df.iterrows():
                id_val = str(row[col_id]).strip()
                if id_val.lower() in ['nan', '', '0', 'none']: continue
                total += 1

                # --- 1. DISEÑO FÍSICO (Etiqueta Impresa) ---
                
                # Cabecera Roja
                c.setFillColor(ROJO_SURAKI)
                c.rect(0, ALTO - 6*mm, ANCHO, 6*mm, fill=1, stroke=0)
                
                # Título Blanco
                c.setFillColor(BLANCO)
                c.setFont("Helvetica-Bold", 8)
                c.drawString(2*mm, ALTO - 4.5*mm, "HIPER SURAKI")
                
                # Sucursal (Derecha)
                c.setFont("Helvetica", 5)
                sucursal = hoja.replace("Base de Datos Maestra", "").replace("Base Datos Maestra", "").strip()
                c.drawRightString(ANCHO - 2*mm, ALTO - 4.5*mm, sucursal[:25])

                # --- DATOS VISIBLES (Izquierda) ---
                c.setFillColor(NEGRO)
                
                # Descripción
                desc = str(row[col_desc]).replace('\n', ' ').strip()
                c.setFont("Helvetica", 7.5) # Fuente ligeramente más grande y legible
                
                x_text = 2*mm
                w_text = 35*mm # Espacio para texto
                
                # Ajuste de líneas
                palabras = desc.split()
                linea1, linea2 = "", ""
                for p in palabras:
                    if c.stringWidth(linea1 + " " + p, "Helvetica", 7.5) < w_text: linea1 += " " + p
                    else: linea2 += " " + p
                
                y_cursor = ALTO - 10*mm
                c.drawString(x_text, y_cursor, linea1.strip())
                if linea2:
                    y_cursor -= 3*mm
                    c.drawString(x_text, y_cursor, linea2.strip()[:28])

                # Marca / Modelo (Gris oscuro para contraste)
                y_cursor -= 4.5*mm
                marca = str(row[col_marca]).replace('nan','') if col_marca else ""
                modelo = str(row[col_modelo]).replace('nan','') if col_modelo else ""
                info_tec = f"{marca} {modelo}".strip() or "GENÉRICO"
                
                c.setFillColor(GRIS_OSCURO)
                c.setFont("Helvetica-Bold", 7)
                texto_ajustable(c, info_tec, x_text, y_cursor, w_text, "Helvetica-Bold", 7)

                # Serial
                if col_serial:
                    ser = str(row[col_serial]).replace('nan','')
                    if ser:
                        y_cursor -= 3.5*mm
                        c.setFillColor(NEGRO)
                        c.setFont("Courier", 6.5)
                        c.drawString(x_text, y_cursor, f"SN: {ser}")

                # --- 2. QR "OFFLINE PREMIUM" (EL TRUCO VISUAL) ---
                qr_x = ANCHO - 18*mm
                qr_y = 6*mm
                qr_size = 16 * mm
                
                # AQUÍ DISEÑAMOS EL "TEXTO BONITO" PARA EL CELULAR
                # Usamos emojis y líneas separadoras para simular una App
                contenido_qr = (
                    f"🔴 ACTIVO SURAKI\n"
                    f"━━━━━━━━━━━━━━━\n"
                    f"🆔 ID:    {id_val}\n"
                    f"📦 BIEN:  {desc[:25]}\n"
                    f"🏭 MARCA: {marca[:15]}\n"
                    f"🔢 MOD:   {modelo[:15]}\n"
                    f"📍 SEDE:  {sucursal[:15]}\n"
                    f"━━━━━━━━━━━━━━━\n"
                    f"✅ Inventario Verificado"
                )
                
                qr_obj = qr.QrCodeWidget(contenido_qr)
                qr_obj.barWidth = qr_size
                qr_obj.barHeight = qr_size
                qr_obj.qrVersion = 1 
                
                # Dibujar QR
                d = Drawing(qr_size, qr_size)
                d.add(qr_obj)
                d.drawOn(c, qr_x, qr_y)

                # --- ID VISIBLE (ROJO) ---
                c.setFillColor(ROJO_SURAKI)
                center_x = qr_x + (qr_size/2)
                y_id = 2.5*mm # Posición fija abajo
                
                c.setFont("Helvetica-Bold", 8)
                # Centrar y reducir si es muy largo
                sz_id = 8
                while c.stringWidth(id_val, "Helvetica-Bold", sz_id) > 18*mm: sz_id -= 0.5
                c.setFont("Helvetica-Bold", sz_id)
                c.drawCentredString(center_x, y_id, id_val)

                c.showPage()

        except Exception as e:
            print(f"⚠️  Advertencia en {hoja}: {e}")

    c.save()
    print(f"\n✅ ¡ÉXITO! PDF Premium generado en: {ARCHIVO_SALIDA}")
    print("   Imprime y escanea para ver el nuevo diseño 'Ficha Técnica' en tu celular.")

if __name__ == "__main__":
    generar_offline_premium()
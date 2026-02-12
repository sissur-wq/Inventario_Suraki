import pandas as pd
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
import os

# ==========================================
#   CONFIGURACIÓN FINAL (3x3 cm)
# ==========================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(BASE_DIR, "INVENTARIO ACTIVOS FIJOS SURAKI 30-01.xlsx")
ARCHIVO_SALIDA = os.path.join(BASE_DIR, "Etiquetas_Suraki_3x3_Texto.pdf")

# Dimensiones exactas (30x30 mm)
ANCHO = 30 * mm
ALTO = 30 * mm

# Colores Corporativos
ROJO_SURAKI = colors.HexColor("#D32F2F") # Rojo intenso
BLANCO = colors.white
NEGRO = colors.black

# Hojas a ignorar
IGNORAR = ['NOMENCLATURA', 'MODELO', 'RESUMEN', 'INDICE', 'PORTADA', 'CODIF']

# ==========================================
#   LÓGICA
# ==========================================

def detectar_encabezados(df_raw):
    keywords = ['DESCRIPCION', 'MARCA', 'MODELO', 'SERIAL', 'CODIGO', 'BIEN']
    for i, row in df_raw.iterrows():
        texto = " ".join([str(val).upper() for val in row.values])
        if sum(1 for k in keywords if k in texto) >= 2: return i
    return None

def buscar_columna(df, candidatos):
    cols = [str(c).strip().upper() for c in df.columns]
    for cand in candidatos:
        cand = cand.upper()
        if cand in cols: return df.columns[cols.index(cand)]
        for i, real in enumerate(cols):
            if cand in real: return df.columns[i]
    return None

# ==========================================
#   GENERADOR
# ==========================================

def generar_etiquetas_texto():
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"❌ FALTA EL ARCHIVO: {ARCHIVO_EXCEL}")
        return

    print("🔴 Generando etiquetas 3x3cm (Estilo Texto Corporativo)...")
    
    try:
        xls = pd.ExcelFile(ARCHIVO_EXCEL)
    except Exception as e:
        print(f"❌ Error Excel: {e}")
        return

    c = canvas.Canvas(ARCHIVO_SALIDA, pagesize=(ANCHO, ALTO))
    total = 0

    for hoja in xls.sheet_names:
        if any(x in hoja.upper() for x in IGNORAR): continue
        
        try:
            df_temp = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=15)
            fila = detectar_encabezados(df_temp)
            if fila is None: continue
            
            df = pd.read_excel(xls, sheet_name=hoja, header=fila)
            
            col_id = buscar_columna(df, ['CODIGO', 'ID', 'ETIQUETA'])
            col_desc = buscar_columna(df, ['DESCRIPCION', 'BIEN', 'NOMBRE'])
            col_marca = buscar_columna(df, ['MARCA'])
            col_modelo = buscar_columna(df, ['MODELO'])
            col_serial = buscar_columna(df, ['SERIAL', 'SERIE', 'S/N'])

            if not col_id: continue
            print(f"   Procesando: {hoja}...")

            for _, row in df.iterrows():
                id_val = str(row[col_id]).strip()
                if id_val.lower() in ['nan', '', '0', 'none']: continue
                total += 1

                # Datos para el QR (Ocultos a la vista, visibles al escanear)
                desc = str(row[col_desc]).strip() if col_desc else "Activo"
                marca = str(row[col_marca]).replace('nan','') if col_marca else ""
                modelo = str(row[col_modelo]).replace('nan','') if col_modelo else ""
                sucursal = hoja.replace("Base de Datos Maestra", "").replace("Base Datos Maestra", "").strip()

                # --- DISEÑO ---

                # 1. CABECERA ROJA
                c.setFillColor(ROJO_SURAKI)
                c.rect(0, ALTO - 6*mm, ANCHO, 6*mm, fill=1, stroke=0)

                # 2. TEXTO "HIPER SURAKI" (Blanco y Centrado)
                c.setFillColor(BLANCO)
                c.setFont("Helvetica-Bold", 7)
                # Centrado perfecto
                c.drawCentredString(ANCHO/2, ALTO - 4.5*mm, "HIPER SURAKI")

                # 3. CÓDIGO QR
                qr_size = 19 * mm
                qr_x = (ANCHO - qr_size) / 2
                qr_y = 5 * mm # Espacio abajo para el ID
                
                # Contenido del QR (Ficha Técnica Virtual)
                contenido_qr = (
                    f"🔴 ACTIVO SURAKI\n"
                    f"🆔 {id_val}\n"
                    f"📦 {desc[:20]}\n"
                    f"🏭 {marca[:10]} {modelo[:10]}\n"
                    f"📍 {sucursal[:15]}"
                )
                
                qr_obj = qr.QrCodeWidget(contenido_qr)
                qr_obj.barWidth = qr_size
                qr_obj.barHeight = qr_size
                qr_obj.qrVersion = 1 
                
                d = Drawing(qr_size, qr_size)
                d.add(qr_obj)
                d.drawOn(c, qr_x, qr_y)

                # 4. ID (Abajo, en Negro)
                c.setFillColor(NEGRO)
                c.setFont("Helvetica-Bold", 8)
                
                # Ajustar tamaño si el ID es muy largo
                sz = 8
                while c.stringWidth(id_val, "Helvetica-Bold", sz) > (ANCHO - 2*mm):
                    sz -= 0.5
                    c.setFont("Helvetica-Bold", sz)
                
                c.drawCentredString(ANCHO/2, 1.5*mm, id_val)

                # Borde gris muy suave (guía de corte)
                c.setLineWidth(0.1)
                c.setStrokeColor(colors.lightgrey)
                c.rect(0, 0, ANCHO, ALTO)

                c.showPage()

        except Exception as e:
            print(f"⚠️  Error en {hoja}: {e}")

    c.save()
    print(f"\n✅ ¡LISTO! PDF generado: {ARCHIVO_SALIDA}")

if __name__ == "__main__":
    generar_etiquetas_texto()
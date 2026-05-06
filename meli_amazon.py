import streamlit as st
import pandas as pd
import io
import re

# --- CONFIGURACIÓN DE LA APP ---
st.set_page_config(page_title="Generador Amazon - IconCase", page_icon="📦", layout="centered")

st.title("📦 Generador de Plantillas Amazon")
st.markdown("Sube tu archivo de Mercado Libre y la plantilla vacía de Amazon para generar el formato final listo para Copy-Paste.")

# --- ZONA DE SUBIDA DE ARCHIVOS ---
col1, col2 = st.columns(2)
with col1:
    archivo_ml = st.file_uploader("1. Archivo de Mercado Libre (Excel)", type=["xlsx", "xls"])
with col2:
    archivo_amazon = st.file_uploader("2. Plantilla de Amazon (Excel)", type=["xlsx", "xlsm"])

# --- FUNCIÓN DE LIMPIEZA ---
def clean_str(s):
    s = str(s).lower()
    s = re.sub(r'[áäâà]', 'a', s)
    s = re.sub(r'[éëêè]', 'e', s)
    s = re.sub(r'[íïîì]', 'i', s)
    s = re.sub(r'[óöôò]', 'o', s)
    s = re.sub(r'[úüûù]', 'u', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

# --- MOTOR PRINCIPAL ---
if archivo_ml and archivo_amazon:
    if st.button("🚀 Generar Archivo Maestro", use_container_width=True):
        with st.spinner("Procesando y cruzando datos..."):
            try:
                # 1. LEER MERCADO LIBRE
                df_ml = pd.read_excel(archivo_ml, sheet_name=0)
                
                def encontrar_columna_ml(palabras_clave):
                    for col in df_ml.columns:
                        col_clean = clean_str(col)
                        if any(clean_str(p) in col_clean for p in palabras_clave):
                            return col
                    return None

                c_titulo = encontrar_columna_ml(['titulo', 'título'])
                c_modelo = encontrar_columna_ml(['nombre del diseño'])
                c_material = encontrar_columna_ml(['materiales del exterior'])
                c_color = encontrar_columna_ml(['color'])
                c_desc = encontrar_columna_ml(['descripci'])
                c_sku = encontrar_columna_ml(['user product id'])
                c_family = encontrar_columna_ml(['family id'])
                c_stock = encontrar_columna_ml(['stock', 'cantidad'])
                c_precio = encontrar_columna_ml(['precio'])

                if c_precio:
                    df_ml[c_precio] = df_ml[c_precio].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False).str.strip()
                    df_ml[c_precio] = pd.to_numeric(df_ml[c_precio], errors='coerce')

                # 2. LEER PLANTILLA AMAZON
                xls_amazon = pd.ExcelFile(archivo_amazon)
                hoja_correcta = xls_amazon.sheet_names[-1]
                for sheet in xls_amazon.sheet_names:
                    if 'Template' in sheet or 'Plantilla' in sheet:
                        hoja_correcta = sheet
                        break
                        
                df_raw = pd.read_excel(xls_amazon, sheet_name=hoja_correcta, header=None)
                header_idx = 2
                for i in range(min(20, len(df_raw))):
                    fila = df_raw.iloc[i].astype(str).str.lower().values
                    if 'item_sku' in fila or 'sku' in fila:
                        header_idx = i
                        break
                        
                df_template = pd.read_excel(xls_amazon, sheet_name=hoja_correcta, header=header_idx)
                template_cols = [c for c in df_template.columns if not str(c).startswith('Unnamed')]
                amazon_map = {clean_str(c): c for c in template_cols}

                amazon_data = []

                def assign(data_dict, col_name, val):
                    clean_name = clean_str(col_name)
                    if clean_name in amazon_map:
                        data_dict[amazon_map[clean_name]] = val

                def assign_any(data_dict, col_names, val):
                    for c in col_names:
                        clean_name = clean_str(c)
                        if clean_name in amazon_map:
                            data_dict[amazon_map[clean_name]] = val
                            break

                familias = df_ml.groupby(c_family) if c_family else []

                # 3. PROCESAR DATOS
                for family_id, group in familias:
                    if pd.isna(family_id): continue

                    modelo_base = group.iloc[0]
                    
                    titulo = str(modelo_base[c_titulo]) if c_titulo and pd.notna(modelo_base[c_titulo]) else 'Funda'
                    modelo_diseno = str(modelo_base[c_modelo]).strip() if c_modelo and pd.notna(modelo_base[c_modelo]) else 'Celular'
                    material_ext = str(modelo_base[c_material]).strip() if c_material and pd.notna(modelo_base[c_material]) else ''
                    descripcion_ml = str(modelo_base[c_desc]).strip() if c_desc and pd.notna(modelo_base[c_desc]) else ''
                    
                    # Ajuste 1: SKU Padre toma el Family ID asegurando que no tenga ".0"
                    sku_padre = str(family_id).strip()
                    if sku_padre.endswith('.0'): 
                        sku_padre = sku_padre[:-2]

                    titulo_lower = titulo.lower()
                    if '3 en 1' in titulo_lower or '360' in titulo_lower:
                        v1 = f"PROTECCIÓN 360° ULTRA REFORZADA: Diseño 3 en 1 que envuelve y protege tu {modelo_diseno} por completo contra impactos extremos."
                        v2 = "SISTEMA DE 3 PIEZAS: Compuesta por 2 marcos y una funda central. (Nota: Se envía preensamblada, desármala antes de instalar y vuelve a cerrarla para un ajuste perfecto)."
                        v3 = "USO RUDO Y ANTIGOLPES: Materiales de alta resistencia que blindan tu equipo ante caídas accidentales."
                        v4 = "AJUSTE PERFECTO: Diseñada milimétricamente para liberar todos los puertos y botones de tu celular."
                        v5 = "ESTILO Y SEGURIDAD: Gran variedad de colores para que personalices tu equipo sin sacrificar protección."
                    elif 'hidrogel' in titulo_lower:
                        v1 = f"KIT DE PROTECCIÓN TOTAL: Incluye una Funda de Uso Rudo + 1 Mica de Hidrogel premium para tu {modelo_diseno}."
                        v2 = "MICA DE 4 CAPAS: Tecnología avanzada. (Importante: La capa instalable es la central. Recomendamos instalación por un profesional para evitar burbujas)."
                        v3 = "CASE ANTIGOLPES: Estructura resistente y gruesa que disipa la fuerza de los impactos de forma eficiente."
                        v4 = "DISEÑO ÚNICO: Variedad de colores vibrantes y diseños para que tu celular refleje tu estilo personal."
                        v5 = "ORILLAS REFORZADAS: Biseles elevados que protegen la pantalla y la cámara al colocar el equipo en superficies planas."
                    else:
                        v1 = f"PROTECCIÓN DE USO RUDO: Funda antigolpes resistente y gruesa, ideal para proteger tu {modelo_diseno} en el ajetreo diario."
                        v2 = "ORILLAS REFORZADAS: Esquinas con sistema de amortiguación que reducen el riesgo de daños por caídas laterales."
                        v3 = "AGARRE SEGURO: Materiales diseñados para evitar que el equipo se te resbale de las manos."
                        v4 = "ACCESO TOTAL: Recortes precisos que respetan al 100% las funciones, botones y cámara de tu dispositivo."
                        v5 = "PERSONALIZA TU EQUIPO: Variedad de colores y diseños para que se adapten a tu estilo a la perfección."

                    parent = {c: '' for c in template_cols}
                    
                    assign(parent, 'SKU', sku_padre)
                    assign(parent, 'Tipo de producto', 'CELLULAR_PHONE_CASE')
                    assign(parent, 'Acción de listado', 'Crear o reemplazar')
                    assign(parent, 'Nivel de relación', 'Principal.')
                    assign(parent, 'Nombre del tema de variación', 'COLOR')
                    assign(parent, 'Nombre del producto', titulo)
                    assign(parent, 'Marca', 'Icon Case')
                    assign(parent, 'Tipo de ID del producto', 'Exento de GTIN')
                    assign(parent, 'Numero de modelo', modelo_diseno)
                    assign(parent, 'Nombre Modelo', modelo_diseno)
                    assign(parent, 'Fabricante', 'Icon Case')
                    assign(parent, 'Descripción Producto', descripcion_ml)
                    assign(parent, 'Viñeta', v1)
                    assign(parent, 'Viñeta.1', v2)
                    assign(parent, 'Viñeta.2', v3)
                    assign(parent, 'Viñeta.3', v4)
                    assign(parent, 'Viñeta.4', v5)
                    assign(parent, 'Palabra clave genérica', f"Funda {modelo_diseno}, Case {modelo_diseno}, Protector {modelo_diseno}")
                    assign(parent, 'Material', material_ext)
                    assign(parent, 'Número de Artículos', '1')
                    assign(parent, 'Numero de pieza', modelo_diseno)
                    assign(parent, 'Modelos de teléfono móvil compatibles', modelo_diseno)
                    assign(parent, 'Dispositivos Compatibles', modelo_diseno)
                    assign(parent, 'Conteo de unidades', '1.0')
                    assign(parent, 'Tipo de conteo de unidades', 'unidad')
                    assign(parent, 'Saltar oferta', 'No')
                    assign(parent, 'Estado del producto', 'Nuevo')
                    assign(parent, 'Moneda del precio de venta recomendado', 'MXN')
                    assign(parent, 'Cumplimiento de código de canal (MX)', 'Gestionado por el vendedor')
                    assign(parent, 'Plantilla de envío (MX)', 'Plantilla de Amazon')
                    assign(parent, 'Garantía de Producto', 'Garantia de 30 dias ')
                    assign(parent, 'País de origen', 'China')
                    
                    # Ajuste 4: Baterías SOLO en el padre
                    assign_any(parent, ['¿Se necesitan baterías?', 'se necesitan baterias', '¿se necesitan baterias?'], 'No')
                    
                    amazon_data.append(parent)

                    for _, row in group.iterrows():
                        child = parent.copy()
                        
                        # Limpiamos todo rastro de "baterías" para los hijos
                        assign_any(child, ['¿Se necesitan baterías?', 'se necesitan baterias', '¿se necesitan baterias?'], '')
                        
                        sku_hijo_raw = str(row[c_sku]).strip() if c_sku and pd.notna(row[c_sku]) else ''
                        sku_hijo = sku_hijo_raw.replace('MLM', '') if sku_hijo_raw.startswith('MLM') else sku_hijo_raw
                        
                        assign(child, 'SKU', sku_hijo)
                        assign(child, 'Nivel de relación', 'Niños')
                        assign(child, 'SKU principal', sku_padre)
                        
                        color_val = str(row[c_color]).strip() if c_color and pd.notna(row[c_color]) else ''
                        assign(child, 'Color', color_val)
                        
                        precio_val = row[c_precio] if c_precio and pd.notna(row[c_precio]) else ''
                        assign(child, 'Precio de venta recomendado (PVPR)', precio_val)
                        assign(child, 'Su precio MXN (Vender en Amazon, MX)', precio_val)
                        
                        # Ajuste 2: Cantidad e Inventario SOLO para los hijos, jalando del Stock
                        stock_val = row[c_stock] if c_stock and pd.notna(row[c_stock]) else ''
                        assign(child, 'Cantidad (MX)', stock_val)
                        assign(child, 'Inventario siempre disponible (MX)', 'Deshabilitado')
                        
                        # Ajuste 3: Medidas SOLO para los hijos, con los nuevos datos de paquete
                        assign(child, 'Valor decimal del grosor del artículo', '1.0')
                        assign(child, 'Valor descriptivo del grosor del artículo', '1')
                        assign(child, 'Unidad del grosor del artículo', 'Centímetros')
                        
                        assign_any(child, ['Longitud del artículo', 'Longitud del artículo.1', 'item_package_dimensions_length'], '18.0')
                        assign_any(child, ['longitud del artículo', 'item_package_dimensions_length_unit'], 'Centímetros')
                        assign_any(child, ['ancho del articulo'], '12.0')
                        assign_any(child, ['Unidad de ancho de artículo'], 'Centímetros')
                        assign_any(child, ['Altura del artículo'], '1.0')
                        assign_any(child, ['Unidad de altura del artículo'], 'Centímetros')
                        
                        assign_any(child, ['Longitud Paquete'], '19.0')
                        assign_any(child, ['Unidad de longitud del paquete'], 'Centímetros')
                        assign_any(child, ['Ancho Paquete'], '15.0')
                        assign_any(child, ['Unidad de anchura del paquete'], 'Centímetros')
                        assign_any(child, ['Altura Paquete'], '1.0')
                        assign_any(child, ['Unidad de altura del paquete'], 'Centímetros')
                        
                        assign(child, 'Peso del paquete', '75.0')
                        assign(child, 'Unidad del peso del paquete', 'Gramos')

                        def get_img(num):
                            c_img = encontrar_columna_ml([f'imagen {num}', f'image {num}'])
                            return str(row[c_img]).strip() if c_img and pd.notna(row[c_img]) else ''

                        assign(child, 'URL de la imagen principal', get_img(1))
                        assign(child, 'Otra URL de la imagen', get_img(2))
                        assign(child, 'Otra URL de la imagen.1', get_img(3))
                        assign(child, 'Otra URL de la imagen.2', get_img(4))
                        assign(child, 'Otra URL de la imagen.3', get_img(5))
                        assign(child, 'Otra URL de la imagen.4', get_img(6))
                        assign(child, 'Otra URL de la imagen.5', get_img(7))
                        assign(child, 'Otra URL de la imagen.6', get_img(8))
                        assign(child, 'Otra URL de la imagen.7', get_img(9))

                        amazon_data.append(child)

                # 4. PREPARAR DESCARGA
                df_final = pd.DataFrame(amazon_data)
                df_final = df_final[template_cols]
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                processed_data = output.getvalue()
                
                st.success("¡Archivo generado con éxito! El Equipo de Trabajo ya lo puede subir directo.")
                st.download_button(
                    label="📥 Descargar Archivo para Amazon",
                    data=processed_data,
                    file_name="Carga_Amazon_Perfect_CopyPaste.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Ocurrió un error: {e}")

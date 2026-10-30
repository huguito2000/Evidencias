import os
import re
import unicodedata
import openpyxl
import pandas as pd
from docxtpl import DocxTemplate


def normalizar_texto(texto):
    """Normaliza texto eliminando acentos, caracteres especiales y convirtiendo a minúsculas."""
    if not isinstance(texto, str):
        texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return re.sub(r'[^a-zA-Z0-9]', '', texto).lower()


def buscar_columna(cols, aliases):
    """
    Busca de forma flexible la columna correspondiente dentro de una lista de nombres de columnas.
    Primero busca coincidencia exacta con algún alias normalizado, y luego coincidencia por inclusión.
    """
    norm_cols = {col: normalizar_texto(col) for col in cols}
    # 1. Coincidencia exacta
    for alias in aliases:
        na = normalizar_texto(alias)
        for col, ncol in norm_cols.items():
            if ncol == na:
                return col
    # 2. Coincidencia por subcadena
    for alias in aliases:
        na = normalizar_texto(alias)
        for col, ncol in norm_cols.items():
            if na in ncol:
                # Evitar falsos positivos como asociar 'resultado esperado' con 'resultado de prueba/estado'
                if na in ['resultado', 'estado'] and 'esperado' in ncol:
                    continue
                return col
    return None


def detectar_fila_cabecera(df_raw):
    """
    Detecta de forma inteligente la fila que contiene las cabeceras reales de los casos de prueba.
    Evalúa cada fila contando cuántas categorías estándar de columnas (ID, Título, Pasos, Descripción, etc.)
    están presentes, evitando falsos positivos con títulos superiores o encabezados institucionales.
    """
    HEADER_CATEGORIES = {
        'id': ['numcp', 'numdecp', 'nodecp', 'cp', 'tc', 'id', 'casodeprueba', 'testcase'],
        'titulo': ['titulo', 'title', 'nombre', 'escenario', 'flujo', 'summary', 'nombredelcaso'],
        'desc': ['descripcion', 'description', 'detalle', 'descripciondelcp'],
        'pasos': ['pasos', 'steps', 'pasosdeejecucion', 'pasosdeprueba', 'procedimiento'],
        'resultado': ['resultadoesperado', 'expectedresult', 'resultado', 'dictamen'],
        'datos': ['datos', 'testdata', 'datosdeprueba'],
        'estado': ['estadodelaprueba', 'estadodeprueba', 'estado', 'estatus', 'status', 'resultado'],
        'prioridad': ['prioridad', 'priority'],
    }
    
    best_row = 0
    best_score = 0
    
    for i, row in df_raw.iterrows():
        # Las celdas de encabezado suelen ser textos cortos (< 60 caracteres)
        cells = [normalizar_texto(c) for c in row if pd.notna(c) and len(str(c).strip()) < 60]
        matched_cats = set()
        for cat, aliases in HEADER_CATEGORIES.items():
            for c in cells:
                if any(alias == c or (len(alias) > 3 and alias in c) for alias in aliases):
                    matched_cats.add(cat)
                    break
        score = len(matched_cats)
        if score > best_score:
            best_score = score
            best_row = i
            
    # Si ninguna fila tiene al menos 3 categorías, devolver la primera fila no vacía o 0 por defecto
    return best_row if best_score >= 2 else 0


def extraer_metadatos_caratula(ruta_excel, pestana, header_row):
    """
    Extrae los metadatos de la cabecera del Excel (filas previas a la tabla de CPs)
    para rellenar la carátula inicial (Tabla 0) del documento Word.
    """
    meta = {
        'proyecto': None,
        'tester': None,
        'folio': None,
        'fecha_inicio': None,
        'fecha_fin': None,
        'estatus': 'Exitoso'
    }
    try:
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
        if pestana not in wb.sheetnames:
            return meta
        ws = wb[pestana]
        
        for r in range(1, header_row + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(r, c).value
                if not val:
                    continue
                val_str = str(val).strip()
                
                # Extraer Folio si viene en formato "Folio: SP001"
                m_folio = re.search(r'folio\s*:\s*([A-Za-z0-9_-]+)', val_str, re.IGNORECASE)
                if m_folio:
                    meta['folio'] = m_folio.group(1)
                
                val_norm = normalizar_texto(val_str)
                # Buscar el valor contiguo no vacío en la misma fila
                next_val = None
                for c2 in range(c + 1, ws.max_column + 1):
                    if ws.cell(r, c2).value is not None:
                        next_val = str(ws.cell(r, c2).value).strip()
                        break
                
                if next_val:
                    if 'proyecto' in val_norm and not meta['proyecto']:
                        meta['proyecto'] = next_val
                    elif ('tester' in val_norm or 'qa' in val_norm) and not meta['tester']:
                        meta['tester'] = next_val
                    elif ('creacion' in val_norm or 'fechainicio' in val_norm) and not meta['fecha_inicio']:
                        meta['fecha_inicio'] = next_val
                        if not meta['fecha_fin']:
                            meta['fecha_fin'] = next_val
                    elif ('fechafin' in val_norm or 'entrega' in val_norm) and not meta['fecha_fin']:
                        meta['fecha_fin'] = next_val
    except Exception as e:
        print(f"    [!] Advertencia al extraer metadatos de la carátula: {e}")
        
    return meta


def actualizar_caratula_word(doc, meta, pestana):
    """
    Actualiza la Tabla 0 de la plantilla de Word con los metadatos extraídos,
    preservando el formato visual, fuentes y estilos de celda.
    """
    def set_cell_text(cell, new_text):
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = new_text
            for run in cell.paragraphs[0].runs[1:]:
                run.text = ''
        else:
            cell.text = new_text

    if not doc.docx.tables:
        return

    t0 = doc.docx.tables[0]
    if len(t0.rows) >= 6:
        if meta.get('proyecto'):
            set_cell_text(t0.rows[0].cells[1], meta['proyecto'])
        if meta.get('tester'):
            set_cell_text(t0.rows[1].cells[1], meta['tester'])
            
        folio_str = meta.get('folio')
        mp_nombre = f"{pestana} ({folio_str})" if folio_str else pestana
        set_cell_text(t0.rows[2].cells[1], mp_nombre)
        
        if meta.get('fecha_inicio'):
            set_cell_text(t0.rows[3].cells[1], meta['fecha_inicio'])
        if meta.get('fecha_fin'):
            set_cell_text(t0.rows[4].cells[1], meta['fecha_fin'])
        if meta.get('estatus'):
            set_cell_text(t0.rows[5].cells[1], meta['estatus'])


def main():
    ruta_documentos = "documentos"
    ruta_salida = "evidencias"
    ruta_plantilla = os.path.join(ruta_documentos, "Evidencias - Template.docx")

    # Verificar existencia de carpetas y plantilla
    if not os.path.exists(ruta_documentos):
        print(f"Error: No se encontró la carpeta de origen '{ruta_documentos}'")
        return

    if not os.path.exists(ruta_plantilla):
        print(f"Error: No se encontró la plantilla '{ruta_plantilla}'")
        print(f"Asegúrate de que el archivo 'Evidencias - Template.docx' esté dentro de la carpeta '{ruta_documentos}'.")
        return

    if not os.path.exists(ruta_salida):
        os.makedirs(ruta_salida)
        print(f"Carpeta de salida '{ruta_salida}' creada.")

    # Listar archivos de Excel en la carpeta documentos, ignorando archivos temporales (~$)
    archivos_excel = [
        f for f in os.listdir(ruta_documentos) 
        if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')
    ]

    if not archivos_excel:
        print(f"No se encontraron archivos de Excel en '{ruta_documentos}'.")
        return

    print(f"Archivos encontrados: {archivos_excel}\n")

    for archivo in archivos_excel:
        ruta_excel = os.path.join(ruta_documentos, archivo)
        print(f"--- Procesando archivo: {archivo} ---")

        try:
            xls = pd.ExcelFile(ruta_excel)
        except Exception as e:
            print(f"Error al leer el Excel {archivo}: {e}")
            continue

        pestanas_excluidas = ["Portada", "Graficas", "Bugs"]
        pestanas_validas = [p for p in xls.sheet_names if p not in pestanas_excluidas]
        
        for pestana in pestanas_validas:
            print(f"  Procesando pestaña: {pestana}...")

            # 1. Leer el Excel completo sin cabecera para detectar la fila exacta de encabezados
            df_temp = pd.read_excel(xls, sheet_name=pestana, header=None)
            header_row = detectar_fila_cabecera(df_temp)
            print(f"    -> Cabecera detectada en fila índice: {header_row} (Fila Excel: {header_row + 1})")

            # 2. Cargar el DataFrame a partir de la fila de cabecera detectada
            df = pd.read_excel(xls, sheet_name=pestana, header=header_row)
            df.columns = [str(col).strip() for col in df.columns]
            cols = list(df.columns)

            col_num = buscar_columna(cols, ['num_cp', 'num de cp', '# de cp', 'no de cp', 'nro de cp', 'cp', 'id', 'caso de prueba', 'test case', 'tc'])
            col_titulo = buscar_columna(cols, ['titulo', 'title', 'escenario / flujo', 'escenario', 'flujo', 'nombre del caso', 'nombre de la prueba', 'nombre', 'caso de uso', 'summary'])
            col_pasos = buscar_columna(cols, ['pasos de ejecucion', 'pasos de prueba', 'pasos', 'steps', 'execution steps', 'pasos a seguir', 'procedimiento'])
            col_estado = buscar_columna(cols, ['estado de la prueba', 'estado de prueba', 'estado', 'estatus', 'status', 'dictamen', 'resultado'])
            col_desc = buscar_columna(cols, ['descripcion del cp', 'descripcion', 'description', 'detalle del caso', 'detalle'])
            col_pre = buscar_columna(cols, ['precondiciones', 'precondicion', 'pre-condiciones', 'prerequisites', 'prerrequisitos'])
            col_datos = buscar_columna(cols, ['datos de prueba', 'test data', 'datos', 'data'])
            col_res = buscar_columna(cols, ['resultado esperado', 'expected result'])
            col_obs = buscar_columna(cols, ['observaciones', 'observacion', 'comments', 'notas'])

            if not col_num:
                print(f"    [!] No se encontró columna de número de CP en '{pestana}'.")
                continue

            print(f"    Columnas mapeadas:")
            print(f"      - ID Caso: '{col_num}'")
            print(f"      - Título:  '{col_titulo}'")
            print(f"      - Pasos:   '{col_pasos}'")
            print(f"      - Estado:  '{col_estado}'")

            registros = []
            for _, row in df.iterrows():
                num_cp_val = str(row[col_num]).strip() if pd.notna(row[col_num]) else ''
                # Ignorar filas vacías o con valores nulos en el ID de CP
                if not num_cp_val or num_cp_val.lower() in ['nan', 'none', '']:
                    continue

                titulo_val = str(row[col_titulo]).strip() if col_titulo and pd.notna(row[col_titulo]) else ''
                pasos_val = str(row[col_pasos]).strip() if col_pasos and pd.notna(row[col_pasos]) else ''
                
                # Estado: si no viene en el Excel o está vacío, asignar "Exitoso" por defecto
                if col_estado and pd.notna(row[col_estado]) and str(row[col_estado]).strip() and str(row[col_estado]).strip().lower() not in ['nan', 'none']:
                    estado_val = str(row[col_estado]).strip()
                else:
                    estado_val = 'Exitoso'

                # Armar la Descripción completa combinando Descripción, Precondiciones, Datos, Resultado Esperado y Observaciones
                partes_desc = []
                if col_desc and pd.notna(row[col_desc]) and str(row[col_desc]).strip():
                    partes_desc.append(str(row[col_desc]).strip())
                if col_pre and pd.notna(row[col_pre]) and str(row[col_pre]).strip():
                    partes_desc.append(f"Precondiciones:\n{str(row[col_pre]).strip()}")
                if col_datos and pd.notna(row[col_datos]) and str(row[col_datos]).strip():
                    partes_desc.append(f"Datos de Prueba:\n{str(row[col_datos]).strip()}")
                if col_res and pd.notna(row[col_res]) and str(row[col_res]).strip():
                    partes_desc.append(f"Resultado Esperado:\n{str(row[col_res]).strip()}")
                if col_obs and pd.notna(row[col_obs]) and str(row[col_obs]).strip():
                    partes_desc.append(f"Observaciones:\n{str(row[col_obs]).strip()}")
                
                desc_val = '\n\n'.join(partes_desc)

                r = {
                    'Num_CP': num_cp_val,
                    'Título': titulo_val,
                    'Titulo': titulo_val,
                    ' Título': titulo_val,
                    ' Titulo': titulo_val,
                    'estado_de_la_prueba': estado_val,
                    'Estado_de_la_prueba': estado_val,
                    'Descripción': desc_val,
                    'Descripcion': desc_val,
                    'pasos': pasos_val,
                    'Pasos': pasos_val,
                }
                registros.append(r)

            if not registros:
                print(f"    [!] No se encontraron registros válidos en '{pestana}'.")
                continue

            print(f"    -> {len(registros)} casos de prueba procesados.")

            # Cargar y renderizar la plantilla docx
            doc = DocxTemplate(ruta_plantilla)
            contexto = {'filas': registros}
            # autoescape=True es indispensable para que caracteres especiales (<, >, &)
            # se escapen como XML válido y no corrompan la estructura interna de Word.
            doc.render(contexto, autoescape=True)

            # Extraer metadatos de carátula del Excel y actualizar la Tabla 0
            metadatos = extraer_metadatos_caratula(ruta_excel, pestana, header_row)
            actualizar_caratula_word(doc, metadatos, pestana)

            # Quitar la extensión del nombre original para el nombre de salida
            nombre_base = os.path.splitext(archivo)[0]
            nombre_salida = os.path.join(ruta_salida, f"{nombre_base}_{pestana}.docx")
            
            try:
                doc.save(nombre_salida)
                print(f"    -> Guardado exitosamente: {nombre_salida}")
            except PermissionError:
                print(f"    [!] Error: No se pudo guardar '{nombre_salida}'. Asegúrate de cerrar el archivo si está abierto.")
            except Exception as e:
                print(f"    [!] Error al guardar '{nombre_salida}': {e}")

    print("\n¡Proceso de todos los archivos completado!")


if __name__ == "__main__":
    main()

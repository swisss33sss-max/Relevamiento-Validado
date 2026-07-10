import os
import shutil
import win32com.client
from collections import defaultdict

def main():
    file_dir = r"c:\Users\hesalinas\Desktop\Etiquetas P-touch\TODO\Aplicacion"
    file_name = "Etiquetas Laboratorio2.xls"
    file_path = os.path.join(file_dir, file_name)
    backup_path = os.path.join(file_dir, "Etiquetas Laboratorio2_Backup.xls")
    
    if not os.path.exists(file_path):
        print(f"Error: El archivo original no existe en {file_path}")
        return

    # 1. Abrir Excel
    print("Iniciando Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = None
    try:
        print(f"Abriendo {file_name}...")
        wb = excel.Workbooks.Open(file_path)
        
        # Seleccionar Hoja1 (primera hoja) y Hoja2 (segunda hoja)
        sheet_source = wb.Sheets(1)
        print(f"Hoja de origen seleccionada: {sheet_source.Name}")
        
        # 2. Buscar Hoja2 y limpiar/asegurar su estado
        sheet_dest = None
        for s in wb.Sheets:
            if s.Name == "Hoja2":
                sheet_dest = s
                break
        
        if not sheet_dest:
            # Si no existe Hoja2, la creamos
            sheet_dest = wb.Sheets.Add(After=sheet_source)
            sheet_dest.Name = "Hoja2"
            print("Hoja2 no existía. Se ha creado.")
        else:
            print("Hoja2 encontrada. Limpiando contenidos previos...")
            sheet_dest.Cells.Clear() # Limpia contenido, formatos e hipervínculos
            
        # 3. Eliminar la hoja 'Repetidos' de ejecuciones previas (si existe)
        repetidos_sheet = None
        for s in wb.Sheets:
            if s.Name == "Repetidos":
                repetidos_sheet = s
                break
        if repetidos_sheet:
            print("Eliminando hoja 'Repetidos' de ejecuciones anteriores...")
            repetidos_sheet.Delete()
            
        # Copiar encabezado de Hoja1 a Hoja2
        sheet_source.Rows(1).Copy(Destination=sheet_dest.Rows(1))
        
        # 4. Escanear filas en Hoja1
        last_row = sheet_source.Cells(sheet_source.Rows.Count, "A").End(-4162).Row # xlUp
        print(f"Filas encontradas en total (incluyendo cabecera): {last_row}")
        
        serial_map = defaultdict(list)
        
        for r in range(2, last_row + 1):
            ser_val = sheet_source.Cells(r, 4).Value # Columna D (Nº Serie)
            ser_str = str(ser_val).strip() if ser_val is not None else ""
            
            # Comprobar si hay link en la columna G (Codigo de barras)
            cell_g = sheet_source.Cells(r, 7)
            has_link = cell_g.Hyperlinks.Count > 0
            val_g = cell_g.Value
            if not has_link and val_g:
                val_str = str(val_g).strip().lower()
                if val_str.startswith("http") or "sharepoint" in val_str:
                    has_link = True
            
            serial_map[ser_str].append({
                'row_idx': r,
                'has_link': has_link
            })
            
        # 5. Clasificar filas a mantener y a eliminar
        rows_to_keep = set()
        rows_to_remove = []
        
        for ser, rows in serial_map.items():
            if ser == "":
                # Si por alguna razón el número de serie está vacío, lo dejamos pasar y no lo consideramos duplicado
                for row_info in rows:
                    rows_to_keep.add(row_info['row_idx'])
                continue
                
            if len(rows) == 1:
                rows_to_keep.add(rows[0]['row_idx'])
            else:
                # Duplicados detectados!
                # Buscar la fila preferida (la primera que tenga link)
                best_row = None
                for row_info in rows:
                    if row_info['has_link']:
                        best_row = row_info['row_idx']
                        break
                        
                # Si ninguno tiene link, conservar el primero
                if best_row is None:
                    best_row = rows[0]['row_idx']
                    
                rows_to_keep.add(best_row)
                
                # Las demás se marcan para eliminar
                for row_info in rows:
                    if row_info['row_idx'] != best_row:
                        rows_to_remove.append(row_info['row_idx'])
                        
        print(f"Filas a conservar en {sheet_source.Name}: {len(rows_to_keep)}")
        print(f"Filas duplicadas a eliminar y mover a {sheet_dest.Name}: {len(rows_to_remove)}")
        
        # 6. Copiar filas repetidas a Hoja2
        dest_row_idx = 2
        for r in sorted(rows_to_remove):
            sheet_source.Rows(r).Copy(Destination=sheet_dest.Rows(dest_row_idx))
            dest_row_idx += 1
        print(f"Copiado de filas duplicadas a {sheet_dest.Name} completado.")
        
        # 7. Eliminar filas duplicadas de la hoja origen (de abajo hacia arriba)
        for r in sorted(rows_to_remove, reverse=True):
            sheet_source.Rows(r).Delete()
        print("Eliminación de filas duplicadas en la hoja origen completada.")
        
        # Guardar y cerrar
        print("Guardando cambios en el archivo...")
        wb.Save()
        print("Archivo guardado con éxito.")
        
    except Exception as e:
        print(f"Ocurrió un error durante el proceso: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.Close(False)
            except:
                pass
        excel.Quit()
        print("Excel cerrado.")

if __name__ == "__main__":
    main()

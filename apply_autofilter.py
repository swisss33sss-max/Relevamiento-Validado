import os
import win32com.client

def clean_float_to_str(val):
    if val is None:
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()

def main():
    file_dir = r"c:\Users\hesalinas\Desktop\Etiquetas P-touch\TODO\Aplicacion"
    file_name = "Etiquetas Laboratorio2.xls"
    file_path = os.path.join(file_dir, file_name)
    
    if not os.path.exists(file_path):
        print(f"Error: El archivo no existe en {file_path}")
        return

    print("Iniciando Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = None
    try:
        print(f"Abriendo {file_name}...")
        wb = excel.Workbooks.Open(file_path)
        
        for s_idx in [1, 2]:
            sheet = wb.Sheets(s_idx)
            sheet.Activate()
            
            # 1. Si la hoja contiene Tablas de Excel (ListObjects), las convertimos a Rangos normales.
            # Esto es vital para poder aplicar un filtro que cubra todas las filas actuales y evitar
            # que Excel bloquee el método AutoFilter a nivel de hoja.
            if sheet.ListObjects.Count > 0:
                print(f"\nDetectadas {sheet.ListObjects.Count} tablas de Excel en {sheet.Name}. Convirtiendo a Rangos...")
                for lo_idx in range(sheet.ListObjects.Count, 0, -1):
                    lo = sheet.ListObjects(lo_idx)
                    print(f"  Convirtiendo tabla '{lo.Name}' a rango normal...")
                    lo.Unlist()
            
            last_row = sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row
            print(f"\nProcesando hoja: {sheet.Name} (filas: {last_row})")
            
            # 2. Convertir la Columna D (Nº Serie) a formato texto y normalizar
            print("  Convirtiendo columna D a formato de Texto y limpiando decimales...")
            for r in range(2, last_row + 1):
                cell = sheet.Cells(r, 4) # Col D
                val = cell.Value
                if val is not None:
                    cleaned_str = clean_float_to_str(val)
                    cell.NumberFormat = "@"
                    cell.Value = cleaned_str
                    
            # 3. Habilitar AutoFiltro cubriendo todo el rango activo
            print("  Habilitando AutoFiltro...")
            if sheet.AutoFilterMode:
                sheet.AutoFilterMode = False
            sheet.Cells(1, 1).AutoFilter()
            
        print("\nGuardando cambios en el archivo...")
        wb.Save()
        print("Filtros aplicados y tipos de datos normalizados con éxito.")
        
    except Exception as e:
        print(f"Ocurrió un error: {e}")
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

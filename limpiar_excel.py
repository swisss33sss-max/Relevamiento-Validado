import pandas as pd
import numpy as np
import sys

def process_file():
    input_file = 'Etiquetas Laboratorio2.xls'
    output_file = 'Etiquetas Laboratorio2_Limpio.xlsx'
    
    print(f"Leyendo {input_file}...")
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error leyendo el archivo: {e}")
        sys.exit(1)
        
    print(f"Total de filas originales: {len(df)}")
    
    # Asumimos que la columna D es la 4ta (índice 3) y la G es la 7ma (índice 6)
    if len(df.columns) < 7:
        print("El archivo no tiene suficientes columnas.")
        sys.exit(1)
        
    col_d = df.columns[3]
    col_g = df.columns[6]
    
    print(f"Columna D identificada como: '{col_d}'")
    print(f"Columna G identificada como: '{col_g}'")
    
    # Crear una columna temporal para saber si tiene link (no nulo y no vacío)
    # Rellenamos NA con string vacío para evaluar fácilmente
    temp_g = df[col_g].fillna('').astype(str).str.strip()
    df['_has_link'] = temp_g != ''
    
    # Separamos las filas donde la columna D está vacía (si las hay, tal vez queramos conservarlas todas, o no)
    # El usuario dijo "que no se repitan los valores de la columna D". 
    # Para estar seguros de que no borramos filas sin serie, podríamos separarlas, pero usualmente drop_duplicates borra las vacías extra.
    # Vamos a seguir la regla estricta: borrar duplicados de D.
    
    # Ordenamos: primero por la columna D, luego por _has_link (True primero, es decir, descendente)
    # Así, para cada valor de D, la fila que sí tiene link queda primero.
    df_sorted = df.sort_values(by=[col_d, '_has_link'], ascending=[True, False], na_position='last')
    
    # Eliminamos duplicados basados en la columna D, conservando el primero (que será el que tiene link)
    df_clean = df_sorted.drop_duplicates(subset=[col_d], keep='first')
    
    # Eliminamos la columna temporal
    df_clean = df_clean.drop(columns=['_has_link'])
    
    # Restauramos el orden original (opcional, pero buena práctica) usando el índice original
    df_clean = df_clean.sort_index()
    
    print(f"Total de filas después de limpiar: {len(df_clean)}")
    print(f"Se eliminaron {len(df) - len(df_clean)} filas repetidas.")
    
    print(f"Guardando en {output_file}...")
    df_clean.to_excel(output_file, index=False)
    print("¡Proceso completado exitosamente!")

if __name__ == '__main__':
    process_file()

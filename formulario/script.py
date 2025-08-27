import pandas as pd
from openpyxl import load_workbook

# --- Configuración ---
ruta_excel = "plantilla.xlsx"
ruta_csv = "datos.csv"
hoja = "Hoja1"

# 1. Cargar datos del CSV
df = pd.read_csv("ruta csv")

# 2. Abrir el Excel existente
wb = load_workbook("ruta hoja de calculo")
ws = wb["Hoja 1"]

# Mapear cada mes a sus columnas en el Excel (ajusta según tu archivo)
columnas = {
    "marzo": {"capital": 2, "interes": 3, "mora": 4},
    "abril": {"capital": 5, "interes": 6, "mora": 7},
    "mayo": {"capital": 8, "interes": 9, "mora": 10},
    "junio": {"capital": 11, "interes": 12, "mora": 13}
}

# 3. Rellenar datos
for index, row in df.iterrows():
    asociado = row["asociado"]
    mes = row["mes"].lower()
    capital, interes, mora = row["capital"], row["interes"], row["mora"]

    # Buscar la fila del asociado
    for fila in range(3, ws.max_row + 1):  # arranca en 3 porque las 2 primeras son encabezados
        if ws.cell(row=fila, column=1).value == asociado:
            # Insertar valores en las columnas correctas
            ws.cell(row=fila, column=columnas[mes]["capital"], value=capital)
            ws.cell(row=fila, column=columnas[mes]["interes"], value=interes)
            ws.cell(row=fila, column=columnas[mes]["mora"], value=mora)

# 4. Guardar
wb.save("plantilla_actualizada1.xlsx")
print("✅ Excel actualizado con los datos del CSV")

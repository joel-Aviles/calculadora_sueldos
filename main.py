from funcs.converter import converter, check_valid_dir
from openpyxl import Workbook
from datetime import datetime as dt
import os

def calculate_amounts(percep_ord, percep_extra, deductions, formula_type):
    formulas = {
        '1': lambda: (percep_ord + percep_extra) - deductions,  # Liquido
        '2': lambda: (percep_ord + percep_extra) - deductions,  # Neto
        '3': lambda: percep_ord + percep_extra,                 # Solo percepciones
        '4': lambda: percep_ord - deductions,                   # Percepciones ordinarias - deducciones de ley
        '5': lambda: percep_extra - deductions,                 # Percepciones extraordinarias - deducciones de ley
        '6': lambda: percep_ord,                                # Solo percepciones ordinarias
        '7': lambda: percep_extra,                              # Solo percepciones extraordinarias
        '8': lambda: (percep_ord + percep_extra) - deductions,  # Sueldo (percepciones 07)
    }
    return formulas.get(formula_type, lambda: 0)()

def get_personal_data(df, personal):
    rfc = df["rfc"].unique()[0]
    person = personal.loc[personal["rfc"] == rfc].iloc[0]
    name = f"{person['nombre'].strip()} {person['paterno'].strip()} {person['materno'].strip()}"
    return rfc, name

def write_excel(ws, section, data, start_row, total_label):
    row = start_row
    ws[f"B{row}"] = section
    row += 1
    ws[f"B{row}"] = "Consecutivo"
    ws[f"C{row}"] = "Tipo Nómina"
    ws[f"D{row}"] = "Tipo"
    ws[f"E{row}"] = "Clave"
    ws[f"F{row}"] = "Concepto"
    ws[f"G{row}"] = "Importe"
    row += 1

    total_sum = 0
    for consec, item in enumerate(data, start=1):
        if item["suma"] != 0:
            ws[f"B{row}"] = consec
            ws[f"C{row}"] = "Ordinario" if "Ordinaria" in section else "Extraordinario"
            ws[f"D{row}"] = item["tipo"]
            ws[f"E{row}"] = item["concepto"]
            ws[f"F{row}"] = item["descrip"]
            ws[f"G{row}"] = "{0:.2f}".format(item["suma"])
            total_sum += item["suma"]
            row += 1

    ws[f"B{row}"] = total_label
    ws[f"G{row}"] = total_sum
    return row + 2, total_sum

def main():
    # 1) Leer archivo
    process_type = input("¿Qué proceso deseas llevar a cabo?\n 1) Pensión alimenticia\n 2) Juicios mercantiles\nIngrese su elección (1 ó 2): ")
    
    if process_type == '1':
        discount_percent = int(input("¿Qué porcentaje se le va a descontar al trabajador? ")) * 0.01
        money_formula = input("¿Qué fórmula se usará?\n 1) Líquido\n 2) Neto\n 3) Solo percepciones\n 4) Percepciones ordinarias - deducciones de ley\n 5) Percepciones extraordinarias - deducciones de ley\n 6) Solo ordinarias\n 7) Solo extraordinarias\n 8) Percepciones 07\nIngrese el número de su elección: ")
        payment_period = int(input("¿Cuántas quincenas se le va a cobrar? "))
    else:
        discount_percent = 1
        money_formula = '1'
        payment_period = 1

    df = converter(input("Nombre del archivo: "))

    # 2) Leer percepciones
    percep = converter("percepciones.xlsx")
    perord = df[(df["tipoconcepto"] == "Percepción") & (df["conceptosiapsep"].isin(percep["clave"]))]
    perext = df[(df["tipoconcepto"] == "Percepción") & (~df["conceptosiapsep"].isin(percep["clave"]))]

    # 3) Leer deducciones
    deducs = converter("deducciones.xlsx")
    deducs["concepto"] = deducs["concepto"].astype(str)
    gended = df[(df["tipoconcepto"] == "Deducción") & (df["conceptosiapsep"].isin(deducs["concepto"]))]
    outded = df[(df["tipoconcepto"] == "Deducción") & (~df["conceptosiapsep"].isin(deducs["concepto"]))]

    # 4) Obtener datos del empleado
    fstqnaproc = df["qnaproc"].min()
    lstqnaproc = df["qnaproc"].max()
    personal = converter("Personal.xlsx")
    rfc, name = get_personal_data(df, personal)

    # 5) Crear y escribir en archivo Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Prueba"

    ws["B2"] = "Datos del Empleado"
    ws["B3"] = "Nombre"
    ws["B4"] = "RFC"
    ws["B5"] = "Qna Devengada"
    ws["C3"] = name
    ws["C4"] = rfc
    ws["C5"] = lstqnaproc

    # Percepciones Ordinarias
    data = [
        {
            "concepto": concepto,
            "descrip": percep.loc[percep["clave"] == concepto, "descripcion"].iloc[0],
            "suma": perord[(perord["conceptosiapsep"] == concepto) & (perord["qnaproc"] == lstqnaproc)]["importe"].sum(),
            "tipo": "Percepción"
        }
        for concepto in perord["conceptosiapsep"].unique()
    ]
    xindex, total_percep_ord = write_excel(ws, f"Percepciones Ordinarias {lstqnaproc}", data, 7, "Total Percepciones Ordinarias")

    # Deducciones Ordinarias
    data = [
        {
            "concepto": concepto,
            "descrip": deducs.loc[deducs["concepto"] == concepto, "descripcion"].iloc[0],
            "suma": gended[(gended["conceptosiapsep"] == concepto) & (gended["qnaproc"] == lstqnaproc)]["importe"].sum(),
            "tipo": "Deducción"
        }
        for concepto in gended["conceptosiapsep"].unique()
    ]
    xindex, total_deduc_ord = write_excel(ws, f"Deducciones Ordinarias {lstqnaproc}", data, xindex, "Total Deducciones Ordinarias")

    # Percepciones Extraordinarias
    nocont = converter("PercepExtra_NoContarPensiones.xlsx")
    data = [
        {
            "concepto": concepto,
            "descrip": perext.loc[perext["conceptosiapsep"] == concepto, "descripciondeconcepto"].iloc[0],
            "suma": perext[(perext["conceptosiapsep"] == concepto) & (perext["qnapago"] >= fstqnaproc)]["importe"].sum(),
            "tipo": "Percepción"
        }
        for concepto in perext["conceptosiapsep"].unique()
        if concepto not in nocont[nocont["cuenta"].str.lower() == "no"]["concepto"].to_list()
    ]
    xindex, total_percep_extra = write_excel(ws, "Percepciones extraordinarias anuales", data, xindex, "Total Percepciones Extraordinarias")

    # 6) Formato para pensiones alimenticias
    if process_type == '1':
        formula_result = calculate_amounts(total_percep_ord, total_percep_extra, total_deduc_ord, money_formula)
        mount_to_discount = formula_result * discount_percent

        ws[f"B{xindex}"] = "Fórmula usada"
        ws[f"C{xindex}"] = money_formula
        xindex += 1
        ws[f"B{xindex}"] = "Descuento del " + str(discount_percent * 100) + "%"
        ws[f"C{xindex}"] = mount_to_discount
        xindex += 2

        # Pagos retroactivos
        ws[f"B{xindex}"] = "Periodo"
        ws[f"C{xindex}"] = str(payment_period) + " Quincenas"
        xindex += 1
        ws[f"B{xindex}"] = "Consecutivo"
        ws[f"C{xindex}"] = "Número de quincena"
        ws[f"D{xindex}"] = "Monto"
        xindex += 1

        mount_per_period = mount_to_discount / payment_period
        for period in range(payment_period):
            lstqnaproc += 1  # Ajuste para saltar de la quincena 24 a la 01, si es necesario
            ws[f"C{xindex}"] = lstqnaproc
            ws[f"D{xindex}"] = "{0:.2f}".format(mount_per_period)
            xindex += 1

    # 7) Guardar archivo Excel
    userpath = os.path.expanduser(os.getenv('USERPROFILE'))
    dirpath = check_valid_dir(f"{userpath}/OneDrive - Secretaría de Educación de Guanajuato/tmp/Pensiones/{rfc}")
    counters = len([file for file in os.listdir(dirpath) if f"{rfc}-" in file and ".xlsx" in file])
    filename = f"{dirpath}/{rfc}-{dt.now().strftime('%d%m%Y')}_{counters + 1}.xlsx"

    try:
        wb.save(filename=filename)
        print(f"Archivo guardado exitosamente con el nombre '{filename.split('/')[-1]}'")
    except Exception as ex:
        print(f"Error = {ex}")
    finally:
        print("Programa Finalizado")

if __name__ == '__main__':
    main()

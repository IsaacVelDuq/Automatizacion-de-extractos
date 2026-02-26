import pandas as pd
from utils import table_utils
import win32com.client as win32
import pythoncom,os
from openpyxl import load_workbook
import numpy as np
def read_db():
    df =  pd.read_excel(r"\\fsapp\archivos_dya\Analitica\Repositorio\Archivos Reportes PWBI\Corporativo\Auditoria interna\Tesoreria\Control tarjetas de credito 2024 (1.0)-LPJUANRB1-2.xlsm", sheet_name="BD (NO MODIFICAR)", engine="openpyxl")
    df = df.drop(df.columns[[0,1,2]], axis=1)
    df = df.drop([0,1,2,3,4], axis=0) 
    headers=[]
    seen={}
    for col in df.iloc[0,:]:
        if str(col) in seen:
            seen[col] += 1 
            headers.append(f"{col}_{seen[col]}") 
        else: 
            seen[col] = 0 
            headers.append(col)    
    df= df.iloc[1:,:]
    headers[0]="ID"
    df.columns= headers
    df = df.drop(["ESTADO"], axis=1)
    df=df.rename(columns={"ESTADO_1":"ESTADO"})
    df= df.astype(str)
    df["NUMERO DE TARJETA"] = df["NUMERO DE TARJETA"].str.strip()
    return df



def create_details(data):
    months = { 1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
    details= pd.DataFrame(data)
    path= details.loc[0,"details"]
    db= read_db()
    details=details.drop(["client"], axis=1)
    df= db.merge(details,left_on="NUMERO DE TARJETA",right_on="account",how="inner")
    df["Mes"] = df["period"].dt.month.map(months)
    df["Año"] = df["period"].dt.year
    df["Valor(USD)"]=0
    df["Legalizaciòn"]= None
    df= df[["NOMBRE","EMPRESA","CARGO/TIPO","TIPO","UNIDAD DE NEGOCIO","BANCO EMISOR","NUMERO DE TARJETA","MONEDA","Mes","CUPO TC","ESTADO","Legalizaciòn","total_balance","Valor(USD)","Año"]]
    df.loc[df["NOMBRE"].str.contains(r"\d"), "NOMBRE"] = df["EMPRESA"]
    columns_dict = {
    "NOMBRE": "Nombre",
    "EMPRESA": "Empresa",
    "CARGO/TIPO": "Cargo",
    "TIPO": "Tipo",
    "UNIDAD DE NEGOCIO": "Unidad de Negocio",
    "BANCO EMISOR": "Banco Emisor",
    "NUMERO DE TARJETA": "Número de Tarjeta",
    "MONEDA": "Moneda",
    "Mes": "Mes",
    "CUPO TC": "Cupo TC",
    "ESTADO": "Estado",
    "Legalizaciòn": "Legalizaciòn",
    "total_balance": "Valor (COP)",
    "Valor(USD)": "Valor(USD)",
    "Año": "Año"
    }

    df.rename(columns=columns_dict, inplace=True)
    df["Cupo TC"]= pd.to_numeric(df["Cupo TC"])
    df["Valor(USD)"]= pd.to_numeric(df["Valor(USD)"])
    df["Año"]= pd.to_numeric(df["Año"])
    df["Valor (COP)"]= pd.to_numeric(df["Valor (COP)"])
    table_utils.format_excel(df,"cuadro auditoria","Control",path)
    return df

def emails():
    df =  pd.read_excel(r"\\fsapp\archivos_dya\Analitica\Repositorio\Archivos Reportes PWBI\Corporativo\Auditoria interna\Tesoreria\Control tarjetas de credito 2024 (1.0)-LPJUANRB1-2.xlsm", sheet_name="BD Mails", engine="openpyxl")
    df= df.rename(columns={"NOMBRE":"CLIENTE"})
    df = df.reset_index(drop=True)
    df = df.astype(str)
    df["TARJETA"]= df["TARJETA"].str.strip()
    return df


def process_email_report(data, df_details):
    path = data[0]["details"]
    df_mails = emails()
    df_report = None
    df = pd.DataFrame(data)
    if os.path.exists(path):
        # Abrir el archivo y listar las hojas
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names
        # Validar existencia de las hojas
        if ("Automatización Envío" in sheets) or ("Error de envío" in sheets):
            if("Automatización Envío" in sheets):
                df_aut = pd.read_excel(xls, sheet_name="Automatización Envío")
            
            if "Error de envío" in sheets:
                df_err = pd.read_excel(xls, sheet_name="Error de envío")

            if("Automatización Envío" in sheets) and ("Error de envío" in sheets):
                df_report = pd.concat([df_aut, df_err], ignore_index=True)
            elif ("Automatización Envío" in sheets):
                df_report=df_aut.copy()
            elif ("Error de envío" in sheets):
                df_report= df_err.copy()
            df_aux= df_report[(df_report["ESTADO"]!="Enviado")]

            if len(df_aux)>1:
                df_aux = df_aux.merge(df[["account","pdf","client"]], left_on="PDF", right_on="pdf", how="inner")
                df_aux["TARJETA"]=df_aux["account"]
                df_aux= df_aux[[
                "CLAVE","account","client", "DESTINATARIO", "UNIDAD DE NEGOCIO",
                "COPIA", "COPIA OCULTA", "ESTADO", "Año", "Mes", "pdf"
                ]]
                df_aux=df_aux.rename(columns={"account":"TARJETA","pdf":"PDF","client":"CLIENTE"})
            df_report = pd.concat([df_report, df_aux], ignore_index=True)
            df_report = df_mails.merge(df_report[["TARJETA", "UNIDAD DE NEGOCIO","Año","Mes","PDF","ESTADO"]], on="TARJETA", how="right")
            df_report = df_report.astype(str)
            df_report = df_report[~df_report["TARJETA"].isnull()]
        else:
            df_db = read_db()
            df = df.merge(df_db, left_on="account", right_on="NUMERO DE TARJETA", how="left")
            df = df.merge(df_mails, left_on="account", right_on="TARJETA", how="left")
            df = df.merge(df_details, left_on="NUMERO DE TARJETA", right_on="Número de Tarjeta", how="inner")
            df = df.rename(columns={"pdf": "PDF"})         
            df["ESTADO"] = "Pendiente"
            df["CLIENTE"] = df["CLIENTE"].fillna(df["NOMBRE"])
            df_report = df.copy()
            
    df_report = df_report[[
        "CLAVE","TARJETA", "CLIENTE", "DESTINATARIO", "UNIDAD DE NEGOCIO",
        "COPIA", "COPIA OCULTA", "ESTADO", "Año", "Mes", "PDF"
        ]]
    df_report.drop_duplicates(inplace=True)

    return df_report


def get_db():
    df= pd.read_excel(r"\\fsapp\archivos_dya\Analitica\Repositorio\Archivos Reportes PWBI\Corporativo\Auditoria interna\Tesoreria\Control tarjetas de credito 2024 (1.0)-LPJUANRB1-2.xlsm",
                       sheet_name="Registros",
                       header=4)
    df = df.drop(df.columns[[0,1]], axis=1)
    df = df.dropna(how="all")
    return df



def insert_in_control(df):
    excel = None
    wb = None
    db= get_db()
    try:
        excel = win32.Dispatch("Excel.Application")       
        excel.Visible = False
        wb = excel.Workbooks.Open(
            r"\\fsapp\archivos_dya\Analitica\Repositorio\Archivos Reportes PWBI\Corporativo\Auditoria interna\Tesoreria\Control tarjetas de credito 2024 (1.0)-LPJUANRB1-2.xlsm"  )

        if wb.ReadOnly:
            raise RuntimeError("El archivo de auditoría ya está abierto por otro usuario")

        ws = wb.Sheets("Registros")
        table = ws.ListObjects("Tabla1")
        table.AutoFilter.ShowAllData()
        db_values = set(zip(db["Nombre"], db["Empresa"], db["Número de Tarjeta"], db["Valor (COP)"], db["Mes"], db["Año"]))

        for i in range(len(df)):
            row = df.iloc[i]
            row_value = (row["Nombre"], row["Empresa"],row["Número de Tarjeta"],row["Valor (COP)"], row["Mes"],row["Año"])
            if (row_value not in db_values):
                table.ListRows.Add(1)
                new_row = table.DataBodyRange.Rows(1)
                previous_row = table.DataBodyRange.Rows(2)
                if (len(db_values)>=2):
                    previous_row.Copy()
                    new_row.PasteSpecial(Paste=-4122)
                new_row.Cells(1, 1).Value = row["Nombre"]
                new_row.Cells(1, 2).Value = row["Empresa"]
                new_row.Cells(1, 3).Value = row["Cargo"]
                new_row.Cells(1, 4).Value = row["Tipo"]
                new_row.Cells(1, 5).Value = row["Unidad de Negocio"]
                new_row.Cells(1, 6).Value = row["Banco Emisor"]
                new_row.Cells(1, 7).Value = row["Número de Tarjeta"]
                new_row.Cells(1, 8).Value = row["Moneda"]
                new_row.Cells(1, 9).Value = row["Mes"]
                new_row.Cells(1, 10).Value = row["Cupo TC"]
                new_row.Cells(1, 11).Value = row["Estado"]
                new_row.Cells(1, 12).Value = row["Legalizaciòn"]
                new_row.Cells(1, 13).Value = row["Valor (COP)"]
                new_row.Cells(1, 14).Value = row["Valor(USD)"]
                new_row.Cells(1, 15).Value = row["Año"]

        wb.Save()

    except PermissionError:
        raise PermissionError("No tienes permisos de escritura sobre el archivo de auditoría")
    except Exception as e:
        raise RuntimeError(f"Error inesperado al guardar el archivo: {e}")

    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
         


def send_emails(path,df):
    df["DESTINATARIO"] = df["DESTINATARIO"].replace("nan", np.nan) 
    try:
        outlook = win32.gencache.EnsureDispatch("Outlook.Application")

        for index, row in df.iterrows():
            try:
                if row["ESTADO"] != "Enviado":
                    mail = outlook.CreateItem(0)

                    if row["UNIDAD DE NEGOCIO"].strip() == "Global Brands":
                        body = f"""Buen día
                        
Adjunto extracto de TC. Por favor, realizar la legalización correspondiente y enviarla con su respectiva autorización.
Recuerden que estos gastos corresponden al mes de {row["Mes"]}.
De no ser entregada en esta fecha, estos gastos deben ser autorizados para registrarlos como no deducibles.
Notas:
    1- Confirmar los gastos.
    2- Al momento de enviar los soportes de los extractos diligenciar debidamente la plantilla de rendición (marca, ceco, firma de la persona y firma del aprobador).

Feliz tarde.

Muchas gracias
Cualquier inquietud quedamos atentos.
"""
                    else:
                        body = f"""Buen día

Adjunto extracto de TC, por favor realizar la solicitud del fondo y su debida legalización en la herramienta de RINDEGASTOS.
Recuerden que estos gastos corresponden al mes de {row["Mes"]}.
De no ser entregada en esta fecha, estos gastos deben ser autorizados para registrarlos como no deducibles.

Feliz tarde.

Muchas gracias
Cualquier inquietud quedamos atentos.
"""
                    if pd.notna(row["DESTINATARIO"]) and str(row["DESTINATARIO"]).strip() != "":
                        mail.To= row["DESTINATARIO"]
                    else:
                        raise Exception("No existe un destinatario")
                    account = row["CLAVE"].strip()
                    title = f"Extracto TC ****{account[-4:]} {row['CLIENTE']} - Periodo {row['Mes']}/{row['Año']}"
                    mail.Subject = title
                    mail.Body = body

                    if pd.notna(row.get("COPIA")):
                        mail.CC = row["COPIA"]
                    if pd.notna(row.get("COPIA OCULTA")):
                        mail.BCC = row["COPIA OCULTA"]

                    # Adjuntos
                    if pd.notna(row.get("PDF")):
                        mail.Attachments.Add(row["PDF"])               
                        mail.Send()
                        df.at[index, "ESTADO"] = "Enviado"
                    

            except Exception as e:
                df.at[index, "ESTADO"] = f"Error: Revisar BD"
        df= df.drop_duplicates()
        table_utils.format_excel(df, "Automatización Envío", "automatizacion", path)

    except Exception as e:
        raise Exception(f"Error inicializando Outlook: {e}")

    finally:
        df_report = df.loc[df["ESTADO"] == "Enviado"] 
        df_null = df.loc[df["ESTADO"] != "Enviado"]
        table_utils.format_excel(df_report, "Automatización Envío", "automatizacion", path)
        table_utils.format_excel(df_null, "Error de envío", "error", path)
        pythoncom.CoUninitialize()
    return len(df_null)>0



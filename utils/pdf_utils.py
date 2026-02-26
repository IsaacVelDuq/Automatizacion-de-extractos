from PyPDF2 import PdfReader,PdfWriter
from datetime import datetime
import re,os


def clean_name(name): 
    # Reemplaza caracteres inválidos en nombres de archivo/carpeta de Windows 
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def save_subpdf(writer,data,folder):
    # Crear carpeta de la empresa dentro de la carpeta principal 
    months = { 1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
    year = data["period"].year
    month = data["period"].month
    year_folder= os.path.join(folder,f"EXTRACTOS {str(year)}")
    month_folder = os.path.join(year_folder,f"{str(month)}-{months[month]}")
    data["details"]= os.path.join(month_folder,f"EXTRACTOS-{months[month]}-{str(year)}.xlsx")
    company_folder = os.path.join(month_folder, clean_name(data["company"])) 
    os.makedirs(company_folder, exist_ok=True)
    #crear la carpeta unica en el que se ubicara pdf y excel
    unique_folder= os.path.join(company_folder,f"{data['account'][-4:]}_{clean_name(data['company'])}_{clean_name(data['client'])}")
    os.makedirs(unique_folder,exist_ok=True)
    # Construir ruta del archivo 
    pdf_name = f"{data['account'][-4:]}_{clean_name(data['company'])}_{clean_name(data['client'])}.pdf"
    path = os.path.join(
    unique_folder,
    pdf_name)
    data['pdf_name']= pdf_name
    
    with open(path, "wb") as f: 
        writer.write(f) 
        data["pdf"] = os.path.abspath(path)
    return data

def get_person(text,company):
    #la persona está en dos linea justo despues de 'Apreciado cliente'
    if company:
        match = re.search(r"Apreciado Cliente\s*\n.+\n(.+)", text)
        name= match.group(1).strip()
        if str(name).isdigit(): name = company
        return name if match else None

    
def get_company(text):
    #Se asume que está en la linea justo despues de 'Apreciado Cliente'
    match = re.search(r"Apreciado Cliente\s*\n(.+)", text)
    return match.group(1).strip() if match else  None

def get_account(text):
#ultimos 4 digitos en el formato '# 0000 0000 0000 0000'
    match = re.search(r"#\s*(\d{4}\s\d{4}\s\d{4}\s\d{4})", text)
    if match:
        num = match.group(1)
        return str(num).strip()
    return None

def get_period(text):
    match = re.search(r"Periodo liquidado\s+([A-Z]{3}\.\d{2}/\d{2}\s*-\s*[A-Z]{3}\.\d{2}/\d{2})", text)
    if match:
        months = { "ENE": "01", "FEB": "02", "MAR": "03",
                   "ABR": "04", "MAY": "05", "JUN": "06", 
                   "JUL": "07", "AGO": "08", "SEP": "09", 
                   "OCT": "10", "NOV": "11", "DIC": "12" }
        date= match.group(1)
        char= "-"
        if(char in date):
            dat= date.split(char)
            date = dat[-1].strip()
            dat= date.split(".")
            month = dat[0]
            aux = dat[-1]
            dat= aux.split("/")
            day= dat[0]
            year= dat[-1]

            date= f"{day}/{months[month]}/20{year}"
        return datetime.strptime(date, "%d/%m/%Y")
    return None

def get_total_balance(text):
    match = re.search(r"Saldo total\s*\$\s*([\d.,]+)", text)
    if match:
        total_balance= match.group(1)
        total_balance= total_balance.split("$")[-1].strip()
        return total_balance.replace(",","")
    return None

def split_pdf(file, folder="output/extractos"):
    """  
    El PDF se segmenta en subdocumentos utilizando como criterio la frase “Apreciado cliente”,
    ya que esta aparece una sola vez en cada archivo.
    Durante el proceso de división también se extraen y registran datos clave: 
    el nombre del cliente, la company y los últimos cuatro dígitos de la account bancaria.
    """
    try:
        reader = PdfReader(file)
        subpdfs = []
        writer= None
        data = {}

        for  page in (reader.pages):
            text = page.extract_text()
            if "Apreciado Cliente" in text:

                if writer and all(data.values()):
                    subpdfs.append(save_subpdf(writer,data,folder))      

                writer= PdfWriter()
                data={"account": None, "company": None, "client": None, "period":None , "total_balance":None}
                
                if not(data["account"]):
                    data["account"]= get_account(text)
                if not(data["company"]):
                    data["company"] = get_company(text)
                if not(data["client"]):
                    data["client"] = get_person(text,data["company"])
                if not(data["period"]):
                    data["period"] = get_period(text)
                if not (data["total_balance"]):
                    data["total_balance"]= get_total_balance(text)

                
            if writer:
                writer.add_page(page)

        if writer and all(data.values()):
            subpdfs.append(save_subpdf(writer,data,folder))
            
        return subpdfs  
    
    except FileNotFoundError:
        print(f"Error, no se encontró el archivo {file}")

    except Exception as e:
        print(f"Ocurrio un error \n {e}")


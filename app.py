import tkinter as tk
from tkinter import filedialog, messagebox
import os

from utils import pdf_utils, table_utils, db_utils


class PDFProcessorApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Extractos Bancarios")
        self.root.geometry("560x300")
        self.root.resizable(False, False)

        self.pdf_path = tk.StringVar()
        self.status_text = tk.StringVar(value="")

        # Título
        tk.Label(
            root,
            text="Procesamiento de Extractos Bancarios",
            font=("Arial", 14, "bold"),
            fg="#2C3E50"
        ).pack(pady=10)

        # Frame selección archivo
        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Entry(
            frame,
            textvariable=self.pdf_path,
            width=45,
            state="readonly"
        ).grid(row=0, column=0, padx=5)

        tk.Button(
            frame,
            text="Seleccionar PDF",
            command=self.select_file,
            bg="#3498DB",
            fg="white"
        ).grid(row=0, column=1, padx=5)

        # Botón ejecutar
        self.run_button = tk.Button(
            root,
            text="Ejecutar proceso",
            width=25,
            height=2,
            command=self.run_process,
            bg="#27AE60",
            fg="white",
            font=("Arial", 10, "bold")
        )
        self.run_button.pack(pady=10)

        # Label de estado
        tk.Label(
            root,
            textvariable=self.status_text,
            fg="blue",
            font=("Arial", 10, "italic")
        ).pack()

    def select_file(self):
        file = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if file:
            self.pdf_path.set(file)

    def run_process(self):
        pdf = self.pdf_path.get()

        if not pdf:
            messagebox.showwarning("Advertencia", "Seleccione un archivo PDF primero")
            return

        # Mostrar estado de carga
        self.status_text.set("⏳ Procesando... por favor espere")
        self.run_button.config(state="disabled")
        self.root.update_idletasks()

        # Lista de pasos
        pasos = [
            "Crear carpeta por cada extracto",
            "Crear documento Excel con movimientos",
            "Crear consolidado mensual",
            "Insertar datos en auditoría",
            "Preparar datos para envío",
            "Enviar extractos por correo"
        ]
        completed = []
        pending = pasos.copy()

        try:
            output_folder = os.path.abspath(r"O:\Finanzas\Tesoreria\PagosDoc\TARJETAS DE CREDITO\EXTRACTOS BANCO DAVIVIENDA")
            os.makedirs(output_folder, exist_ok=True)

            data = pdf_utils.split_pdf(pdf, output_folder)
            completed.append(pending.pop(0))

            data = table_utils.create_all_excels(data)
            completed.append(pending.pop(0))

            path = data[0]["details"]
            df = db_utils.create_details(data)
            completed.append(pending.pop(0))

            db_utils.insert_in_control(df)
            completed.append(pending.pop(0))

            df=db_utils.process_email_report(data, df)
            completed.append(pending.pop(0))

            no_sent=db_utils.send_emails(path,df)
            completed.append(pending.pop(0))

            # Si todo salió bien
            self.status_text.set("✅ Proceso finalizado con éxito")
            messagebox.showinfo(
                "Proceso completado",
                f"✅ El procesamiento se completó correctamente.\n\n"
                f"✅ Pasos realizados:\n- " + "\n- ".join(completed) +
                f"\n\nArchivos guardados en:\n{output_folder}"
            )
            if no_sent:
                messagebox.showwarning(

                    "Extractos no enviados\n",
                    f"Algunos extractos no pudieron ser enviados correctamente,\npor favor revise la hoja 'Error de envío' en el archivo excel resumen del mes correspondiente.\n\n"

            )


        except Exception as e:
            # Mostrar qué pasos quedaron pendientes y cuáles sí se hicieron
            self.status_text.set("❌ Error durante el proceso")
            mensaje = (
                f"❌ Ocurrió un error durante el proceso:\n\n{e}\n\n"
                f"✅ Pasos completados:\n- " + "\n- ".join(completed) +
                ("\n\n❌ Pasos pendientes:\n- " + "\n- ".join(pending) if pending else "")
            )
            messagebox.showerror("Error", mensaje)

        finally:
            self.run_button.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()

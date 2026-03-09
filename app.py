import tkinter as tk
from tkinter import filedialog, messagebox
import os
import threading
import time

try:
    from utils import pdf_utils, table_utils, db_utils
except ImportError:
    pass

BG         = "#0F1923"
CARD       = "#162130"
BORDER     = "#1E3048"
ACCENT     = "#00C896"
ACCENT_DIM = "#00785A"
TEXT_HI    = "#E8F1F8"
TEXT_MID   = "#7A99B8"
TEXT_LO    = "#3A5570"
WARN       = "#F4A823"
ERROR      = "#E85D5D"
SUCCESS    = "#00C896"


class StepRow(tk.Frame):
    SIZE = 24
    def __init__(self, parent, number, label, **kwargs):
        super().__init__(parent, bg=CARD, **kwargs)
        self.number = number
        self.canvas = tk.Canvas(self, width=self.SIZE, height=self.SIZE,
                                bg=CARD, highlightthickness=0)
        self.canvas.pack(side="left", padx=(0, 12))
        self.lbl = tk.Label(self, text=label, bg=CARD,
                            fg=TEXT_LO, font=("Courier New", 9))
        self.lbl.pack(side="left", anchor="w")
        self.set_state("idle")

    def set_state(self, state):
        c = self.canvas
        c.delete("all")
        s = self.SIZE
        cfg = {
            "idle":   (BORDER,  TEXT_LO,  str(self.number), TEXT_LO,  "normal"),
            "active": (WARN,    "#0F1923","▶",               WARN,     "bold"),
            "done":   (ACCENT,  "#0F1923","✓",               TEXT_HI,  "bold"),
            "error":  (ERROR,   "#ffffff","✕",               ERROR,    "bold"),
        }
        fill, sfg, sym, lfg, wt = cfg[state]
        c.create_oval(1, 1, s-1, s-1, fill=fill, outline=fill)
        c.create_text(s//2, s//2, text=sym, fill=sfg,
                      font=("Courier New", 8, "bold"))
        self.lbl.config(fg=lfg, font=("Courier New", 9, wt))


class AnimatedBar(tk.Canvas):
    H = 6
    def __init__(self, parent, total, **kwargs):
        super().__init__(parent, height=self.H, bg=CARD,
                         highlightthickness=0, **kwargs)
        self.total = total
        self._val = 0.0
        self._target = 0.0
        self.bind("<Configure>", lambda e: self._draw(self._val))

    def _rrect(self, x1, y1, x2, y2, r, fill):
        r = min(r, (y2-y1)//2, max(1,(x2-x1)//2))
        pts = [x1+r,y1,x2-r,y1,x2,y1,x2,y1+r,
               x2,y2-r,x2,y2,x2-r,y2,x1+r,y2,
               x1,y2,x1,y2-r,x1,y1+r,x1,y1]
        self.create_polygon(pts, smooth=True, fill=fill, outline=fill)

    def _draw(self, val):
        self.delete("all")
        w = self.winfo_width() or 460
        h = self.H
        self._rrect(0, 0, w, h, h//2, BORDER)
        if val > 0:
            fw = max(h, int(w * val / self.total))
            self._rrect(0, 0, fw, h, h//2, ACCENT)

    def animate_to(self, target):
        self._target = float(target)
        self._tick()

    def _tick(self):
        diff = self._target - self._val
        if abs(diff) < 0.04:
            self._val = self._target
            self._draw(self._val)
            return
        self._val += diff * 0.22
        self._draw(self._val)
        self.after(16, self._tick)


class PDFProcessorApp:
    STEPS = [
        "Separar extractos PDF",
        "Crear Excels por cuenta",
        "Crear consolidado mensual",
        "Insertar datos en auditoría",
        "Preparar reporte de envío",
        "Enviar extractos por correo",
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Extractos")
        self.root.geometry("520x640")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)
        self._full_path = ""
        self.pdf_var = tk.StringVar(value="Ningún archivo seleccionado")
        self._build()

    def _card(self, parent, builder, pady=(0, 0)):
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill="x", padx=30, pady=pady)
        inner = tk.Frame(outer, bg=CARD)
        inner.pack(fill="x", padx=1, pady=1)
        builder(inner)

    def _build(self):
        h = tk.Frame(self.root, bg=BG)
        h.pack(fill="x", padx=30, pady=(30, 0))
        tk.Label(h, text="PROCESADOR", bg=BG, fg=ACCENT,
                 font=("Courier New", 10, "bold")).pack(anchor="w")
        tk.Label(h, text="Extractos Bancarios", bg=BG, fg=TEXT_HI,
                 font=("Georgia", 22, "bold")).pack(anchor="w")
        tk.Label(h, text="GCO  ·  Tesorería", bg=BG, fg=TEXT_MID,
                 font=("Courier New", 8)).pack(anchor="w", pady=(2, 0))

        rule = tk.Canvas(self.root, height=2, bg=BG, highlightthickness=0)
        rule.pack(fill="x", padx=30, pady=(16, 22))
        def draw_rule(e):
            rule.delete("all")
            rule.create_rectangle(0, 0, e.width//3, 2, fill=ACCENT, outline="")
            rule.create_rectangle(e.width//3, 0, e.width, 2, fill=BORDER, outline="")
        rule.bind("<Configure>", draw_rule)

        self._card(self.root, self._file_picker)
        self._card(self.root, self._steps_panel, pady=(12, 0))
        self._card(self.root, self._progress_panel, pady=(12, 0))

        self.run_btn = tk.Button(
            self.root, text="EJECUTAR  →", command=self.run_process,
            bg=ACCENT, fg="#0F1923", bd=0, relief="flat",
            font=("Courier New", 11, "bold"), cursor="hand2",
            activebackground=ACCENT_DIM, activeforeground="#0F1923"
        )
        self.run_btn.pack(fill="x", padx=30, pady=(16, 30), ipady=14)
        self.run_btn.bind("<Enter>", lambda e: self.run_btn.config(bg=ACCENT_DIM))
        self.run_btn.bind("<Leave>", lambda e: self.run_btn.config(bg=ACCENT))

    def _file_picker(self, f):
        tk.Label(f, text="ARCHIVO PDF", bg=CARD, fg=TEXT_LO,
                 font=("Courier New", 7, "bold")).pack(anchor="w", padx=16, pady=(14, 6))
        row = tk.Frame(f, bg=CARD)
        row.pack(fill="x", padx=16, pady=(0, 14))
        nf = tk.Frame(row, bg=BORDER)
        nf.pack(side="left", fill="x", expand=True)
        self._file_lbl = tk.Label(nf, textvariable=self.pdf_var,
                                   bg="#111C28", fg=TEXT_MID,
                                   font=("Courier New", 8), anchor="w",
                                   padx=10, pady=7)
        self._file_lbl.pack(fill="x", padx=1, pady=1)
        btn = tk.Button(row, text="Examinar", command=self.select_file,
                        bg=BORDER, fg=TEXT_HI, bd=0, relief="flat",
                        font=("Courier New", 8, "bold"), cursor="hand2",
                        activebackground="#273D56", padx=14, pady=7)
        btn.pack(side="left", padx=(8, 0))
        btn.bind("<Enter>", lambda e: btn.config(bg="#273D56"))
        btn.bind("<Leave>", lambda e: btn.config(bg=BORDER))

    def _steps_panel(self, f):
        tk.Label(f, text="PIPELINE", bg=CARD, fg=TEXT_LO,
                 font=("Courier New", 7, "bold")).pack(anchor="w", padx=16, pady=(14, 8))
        self.step_rows = []
        for i, label in enumerate(self.STEPS):
            sr = StepRow(f, i+1, label)
            sr.pack(fill="x", padx=16, pady=2)
            self.step_rows.append(sr)
        tk.Frame(f, height=14, bg=CARD).pack()

    def _progress_panel(self, f):
        inner = tk.Frame(f, bg=CARD)
        inner.pack(fill="x", padx=16, pady=14)
        top = tk.Frame(inner, bg=CARD)
        top.pack(fill="x", pady=(0, 6))
        self.status_lbl = tk.Label(top, text="En espera", bg=CARD,
                                    fg=TEXT_LO, font=("Courier New", 8))
        self.status_lbl.pack(side="left")
        self.pct_lbl = tk.Label(top, text="", bg=CARD,
                                 fg=ACCENT, font=("Courier New", 9, "bold"))
        self.pct_lbl.pack(side="right")
        self.bar = AnimatedBar(inner, total=len(self.STEPS))
        self.bar.pack(fill="x")

    # ── UI helpers (siempre llamar desde hilo principal vía root.after) ─────────

    def _ui(self, fn, *args):
        """Encola una actualización de UI en el hilo principal."""
        self.root.after(0, fn, *args)

    def _start_step(self, idx):
        """Marca un paso como activo."""
        self.step_rows[idx].set_state("active")
        self.bar.animate_to(idx)
        pct = int(idx / len(self.STEPS) * 100)
        self.pct_lbl.config(text=f"{pct}%")
        self.status_lbl.config(
            text=f"{idx+1}/{len(self.STEPS)}  ·  {self.STEPS[idx]}", fg=WARN)

    def _complete_step(self, idx):
        """Marca un paso como completado y avanza la barra."""
        self.step_rows[idx].set_state("done")
        self.bar.animate_to(idx + 1)
        pct = int((idx + 1) / len(self.STEPS) * 100)
        self.pct_lbl.config(text=f"{pct}%")
        self.status_lbl.config(text=f"✓  {self.STEPS[idx]}", fg=ACCENT)

    def _finish(self):
        for sr in self.step_rows:
            sr.set_state("done")
        self.bar.animate_to(len(self.STEPS))
        self.pct_lbl.config(text="100%")
        self.status_lbl.config(text="Completado exitosamente", fg=SUCCESS)

    def _mark_error(self, idx):
        self.step_rows[idx].set_state("error")
        self.status_lbl.config(text=f"Error en paso {idx+1}", fg=ERROR)

    def _reset(self):
        for sr in self.step_rows:
            sr.set_state("idle")
        self.bar.animate_to(0)
        self.pct_lbl.config(text="")
        self.status_lbl.config(text="En espera", fg=TEXT_LO)

    def _step(self, idx, fn, *args):
        """
        Ejecuta fn(*args) en el hilo de trabajo,
        actualizando la UI antes y después del paso.
        Retorna el resultado de fn.
        """
        # Señala inicio y espera que el hilo principal procese el cambio
        done_event = threading.Event()
        self._ui(lambda: (self._start_step(idx), done_event.set()))
        done_event.wait()

        result = fn(*args)

        # Señala fin y espera confirmación
        done_event.clear()
        self._ui(lambda: (self._complete_step(idx), done_event.set()))
        done_event.wait()

        return result

    # ── Actions ────────────────────────────────────────────────────────────────

    def select_file(self):
        file = filedialog.askopenfilename(
            title="Seleccionar extracto PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if file:
            self._full_path = file
            self.pdf_var.set(os.path.basename(file))
            self._file_lbl.config(fg=TEXT_HI)

    def run_process(self):
        if not self._full_path:
            messagebox.showwarning("Sin archivo", "Selecciona un PDF primero.")
            return
        self._reset()
        self.run_btn.config(state="disabled", bg=BORDER, fg=TEXT_MID,
                            text="PROCESANDO…")
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        pdf = self._full_path
        completed, pending = [], self.STEPS.copy()
        try:
            control_path  = os.path.abspath(
                r"\\fsapp\archivos_dya\Analitica\Repositorio\Archivos Reportes PWBI\Corporativo\Auditoria interna\Tesoreria\Control tarjetas de credito 2024 (1.0)-LPJUANRB1-2.xlsm")
            output_folder = os.path.abspath(
                r"O:\Finanzas\Tesoreria\PagosDoc\TARJETAS DE CREDITO\EXTRACTOS BANCO DAVIVIENDA")
            os.makedirs(output_folder, exist_ok=True)

            data = self._step(0, pdf_utils.split_pdf, pdf, output_folder)
            completed.append(pending.pop(0))

            data = self._step(1, table_utils.create_all_excels, data)
            completed.append(pending.pop(0))

            df = self._step(2, db_utils.create_details, data, control_path)
            completed.append(pending.pop(0))

            self._step(3, db_utils.insert_in_control, df, control_path)
            completed.append(pending.pop(0))

            df = self._step(4, db_utils.process_email_report, data, df, control_path)
            completed.append(pending.pop(0))

            # Paso 6: envío (comentado)
            done_event = threading.Event()
            self._ui(lambda: (self._start_step(5), done_event.set()))
            done_event.wait()
            no_sent = db_utils.send_emails(data[0]["details"], df)
            done_event.clear()
            self._ui(lambda: (self._complete_step(5), done_event.set()))
            done_event.wait()
            completed.append(pending.pop(0))

            self._ui(self._finish)
            self.root.after(300, lambda: messagebox.showinfo(
                "Proceso completado",
                "El procesamiento finalizó correctamente.\n\n"
                "Pasos realizados:\n  · " + "\n  · ".join(completed) +
                f"\n\nArchivos guardados en:\n{output_folder}"
            ))

        except Exception as e:
            idx = len(completed)
            self._ui(self._mark_error, idx)
            self.root.after(200, lambda err=e: messagebox.showerror(
                "Error en el proceso",
                f"Ocurrió un error:\n\n{err}\n\n"
                "Completados:\n  · " + "\n  · ".join(completed) +
                (f"\n\nPendientes:\n  · " + "\n  · ".join(pending) if pending else "")
            ))
        finally:
            self.root.after(0, lambda: self.run_btn.config(
                state="normal", bg=ACCENT, fg="#0F1923", text="EJECUTAR  →"))


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()
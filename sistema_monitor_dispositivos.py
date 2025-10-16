
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime

ARQUIVO = "monitor_multidispositivo.xlsx"
DEVICES = ["Pirometro","Temporizador","Dinometro"]

def enviar_email_outlook(para, cc, assunto, corpo_html):
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError("Instale pywin32: pip install pywin32") from e
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = para
    if cc:
        mail.CC = cc
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    mail.Send()

def contar_decimais(valor):
    if valor is None:
        return None
    s = str(valor).strip().replace(",", ".")
    m = re.match(r"^-?\d+(?:\.(\d+))?$", s)
    if not m:
        return None
    dec = m.group(1)
    return len(dec) if dec else 0

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Monitor de Dispositivos — Pirometro / Temporizador / Dinometro")
        self.geometry("880x540")

        self.maquina_var = tk.StringVar()

        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=8)

        ttk.Label(top, text="Máquina:").pack(side="left")
        self.maquina_entry = ttk.Entry(top, width=30, textvariable=self.maquina_var)
        self.maquina_entry.pack(side="left", padx=6)

        ttk.Button(top, text="Salvar Máquina na Principal", command=self.save_principal).pack(side="left", padx=8)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=12, pady=8)

        self.frames = {}
        for dev in DEVICES:
            f = DeviceFrame(self.nb, device=dev)
            self.nb.add(f, text=dev)
            self.frames[dev] = f

    def save_principal(self):
        path = Path(ARQUIVO)
        if not path.exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado: {path.resolve()}")
            return
        try:
            df = pd.DataFrame([{"Maquina": self.maquina_var.get().strip(),
                                "Ultima_atualizacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}])
            with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
                df.to_excel(xw, index=False, sheet_name="Principal")
            messagebox.showinfo("OK", "Máquina salva na tela Principal.")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

class DeviceFrame(ttk.Frame):
    def __init__(self, parent, device):
        super().__init__(parent)
        self.device = device

        # Left form
        form = ttk.Frame(self)
        form.pack(side="left", fill="y", padx=10, pady=10)

        self.vars = {k: tk.StringVar() for k in [
            "Data_registro","Maquina","Codigo","Data_vencimento",
            "Amostra_leitura","Unidade",
            "Tecnico_nome","Tecnico_WWID",
            "Emails_ajuste_1","Emails_ajuste_2","Emails_ajuste_3","Emails_ajuste_4","Emails_ajuste_5","Emails_ajuste_6"
        ]}

        row=0
        def add_row(lbl, key, width=30):
            nonlocal row
            ttk.Label(form, text=lbl).grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=self.vars[key], width=width).grid(row=row, column=1, sticky="w", pady=3)
            row+=1

        add_row("Data registro (dd/mm/aaaa):", "Data_registro", 20)
        add_row("Máquina:", "Maquina", 25)
        add_row("Código:", "Codigo", 20)
        add_row("Data vencimento (dd/mm/aaaa):", "Data_vencimento", 20)
        add_row("Amostra leitura:", "Amostra_leitura", 20)
        add_row("Unidade:", "Unidade", 10)
        add_row("Técnico:", "Tecnico_nome", 25)
        add_row("WWID:", "Tecnico_WWID", 20)
        ttk.Label(form, text="E-mails para ajuste (até 6):").grid(row=row, column=0, sticky="w", pady=(10,3)); row+=1
        for i in range(1,7):
            add_row(f"Email {i}:", f"Emails_ajuste_{i}", 30)

        btns = ttk.Frame(form)
        btns.grid(row=row, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="Registrar + Enviar e-mail (ajuste)", command=self.registrar_enviar).pack(side="left", padx=6)
        ttk.Button(btns, text="Somente Registrar", command=self.registrar).pack(side="left", padx=6)

    def registrar(self):
        try:
            self._registrar_core(enviar=False)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def registrar_enviar(self):
        try:
            self._registrar_core(enviar=True)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _registrar_core(self, enviar=False):
        path = Path(ARQUIVO)
        if not path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {path.resolve()}")
        xls = pd.read_excel(path, sheet_name=None)
        df = xls[self.device]

        # Build new row
        r = {k: v.get().strip() for k, v in self.vars.items()}
        if not r["Maquina"]:
            raise ValueError("Preencha a Máquina.")
        if not r["Data_registro"]:
            r["Data_registro"] = datetime.now().strftime("%d/%m/%Y")
        dec = contar_decimais(r["Amostra_leitura"])

        # Fill other fields
        r["Decimais_detectados"] = dec
        r["Decimais_esperados"] = ""  # preenchido pelo monitor diário via Cadastro
        r["Status_validacao"] = ""
        r["Certificado_calibracao"] = ""
        r["Avisar_dias_antes"] = ""
        r["Status_calibracao"] = ""
        r["Alerta_decimais_enviado_em"] = ""
        r["Alerta_calibracao_enviado_em"] = ""
        r["Ajuste_email_enviado_em"] = ""

        df = pd.concat([df, pd.DataFrame([r])], ignore_index=True)

        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
            df.to_excel(xw, index=False, sheet_name=self.device)
            pd.read_excel(path, sheet_name="Cadastro").to_excel(xw, index=False, sheet_name="Cadastro")
            pd.read_excel(path, sheet_name="Principal").to_excel(xw, index=False, sheet_name="Principal")

        if enviar:
            emails = [r.get(f"Emails_ajuste_{i}") for i in range(1,7)]
            emails = [e for e in emails if e]
            if not emails:
                messagebox.showwarning("Aviso", "Nenhum e-mail informado para ajuste. Registro salvo sem envio.")
                return
            para = ";".join(emails)
            assunto = f"[AJUSTE] {r['Maquina']} / {self.device} — ajuste registrado"
            corpo = f"""
            <p>Olá,</p>
            <p>Um ajuste foi registrado no <b>{self.device}</b> da máquina <b>{r['Maquina']}</b>.</p>
            <ul>
              <li><b>Data registro:</b> {r['Data_registro']}</li>
              <li><b>Código:</b> {r['Codigo']}</li>
              <li><b>Amostra:</b> {r['Amostra_leitura']} {r['Unidade']}</li>
              <li><b>Vencimento calibração:</b> {r['Data_vencimento']}</li>
              <li><b>Técnico/WWID:</b> {r['Tecnico_nome']} / {r['Tecnico_WWID']}</li>
              <li><b>Decimais detectados:</b> {dec}</li>
            </ul>
            <p>— Envio automático pela tela de ajuste</p>
            """
            enviar_email_outlook(para, "", assunto, corpo)
            # Atualiza flag de envio no último registro
            xls2 = pd.read_excel(path, sheet_name=None)
            df2 = xls2[self.device]
            df2.loc[len(df2)-1, "Ajuste_email_enviado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
                df2.to_excel(xw, index=False, sheet_name=self.device)
                pd.read_excel(path, sheet_name="Cadastro").to_excel(xw, index=False, sheet_name="Cadastro")
                pd.read_excel(path, sheet_name="Principal").to_excel(xw, index=False, sheet_name="Principal")
            messagebox.showinfo("OK", "Registro salvo e e-mail enviado.")
        else:
            messagebox.showinfo("OK", "Registro salvo.")

if __name__ == "__main__":
    app = App()
    app.mainloop()

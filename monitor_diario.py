
import re
import pandas as pd
from pathlib import Path
from datetime import datetime, date

ARQUIVO = "monitor_multidispositivo.xlsx"
SHEETS_DEVICES = ["Pirometro","Temporizador","Dinometro"]
SHEET_CAD = "Cadastro"

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
    if pd.isna(valor):
        return None
    s = str(valor).strip().replace(",", ".")
    m = re.match(r"^-?\d+(?:\.(\d+))?$", s)
    if not m:
        return None
    dec = m.group(1)
    return len(dec) if dec else 0

def parse_data(d):
    if pd.isna(d) or str(d).strip() == "":
        return None
    if isinstance(d, (datetime, date)):
        return datetime(d.year, d.month, d.day)
    s = str(d).strip().replace("\\n"," ").replace("\\r"," ")
    for fmt in ("%Y-%m-%d","%d/%m/%Y","%d-%m-%Y","%Y/%m/%d","%d.%m.%Y","%m/%d/%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt
        except ValueError:
            continue
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return datetime(dt.year, dt.month, dt.day)
    except Exception:
        return None

def run():
    path = Path(ARQUIVO)
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path.resolve()}")
    cad = pd.read_excel(path, sheet_name=SHEET_CAD)
    cad_idx = cad.set_index(["Maquina","Dispositivo"])
    xls = pd.read_excel(path, sheet_name=None)

    hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    enviados_dec, enviados_cal = 0, 0

    for device in SHEETS_DEVICES:
        df = xls[device].copy()

        # Ensure columns exist
        needed_cols = [
            "Decimais_esperados","Decimais_detectados","Status_validacao",
            "Avisar_dias_antes","Status_calibracao",
            "Alerta_decimais_enviado_em","Alerta_calibracao_enviado_em"
        ]
        for c in needed_cols:
            if c not in df.columns:
                df[c] = ""

        for i, row in df.iterrows():
            maq = row.get("Maquina")
            chave = (maq, device if device!="Temporizador" else "Temporizador")
            email_para = email_cc = ""
            esperado = None
            aviso = 15

            if chave in cad_idx.index:
                esperado = cad_idx.loc[chave, "Decimais_esperados"]
                email_para = str(cad_idx.loc[chave, "Email_gestor"] or "").strip()
                email_cc = str(cad_idx.loc[chave, "Email_copia"] or "").strip()
                try:
                    aviso = int(cad_idx.loc[chave, "Avisar_dias_antes"])
                except Exception:
                    aviso = 15

            # ---- Decimais ----
            leitura = row.get("Amostra_leitura")
            dec = contar_decimais(leitura)
            df.loc[i, "Decimais_detectados"] = dec
            if esperado is not None:
                df.loc[i, "Decimais_esperados"] = int(esperado)

            ja_dec = str(row.get("Alerta_decimais_enviado_em") or "").strip()
            if esperado is None:
                df.loc[i, "Status_validacao"] = "SEM_CADASTRO"
            elif dec is None:
                df.loc[i, "Status_validacao"] = "LEITURA_INVALIDA"
            elif dec != int(esperado):
                df.loc[i, "Status_validacao"] = "ALERTA_DECIMAIS"
                if not ja_dec and email_para:
                    assunto = f"[ALERTA] {maq} / {device}: casas decimais divergentes (esp: {esperado}, det: {dec})"
                    corpo = f"""
                    <p>Olá,</p>
                    <p>Detectamos divergência na precisão decimal do <b>{device}</b> da máquina <b>{maq}</b>.</p>
                    <ul>
                      <li><b>Esperado:</b> {esperado}</li>
                      <li><b>Detectado:</b> {dec}</li>
                      <li><b>Amostra:</b> {row.get("Amostra_leitura")} {row.get("Unidade") or ""}</li>
                      <li><b>Código:</b> {row.get("Codigo") or ""}</li>
                      <li><b>Técnico/WWID:</b> {row.get("Tecnico_nome") or ""} / {row.get("Tecnico_WWID") or ""}</li>
                    </ul>
                    <p>Favor ajustar/validar conforme procedimento.</p>
                    <p>— Alerta automático</p>
                    """
                    enviar_email_outlook(email_para, email_cc, assunto, corpo)
                    df.loc[i, "Alerta_decimais_enviado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    enviados_dec += 1
            else:
                df.loc[i, "Status_validacao"] = "OK"

            # ---- Calibracao ----
            validade = parse_data(row.get("Data_vencimento"))
            ja_cal = str(row.get("Alerta_calibracao_enviado_em") or "").strip()
            if validade is None:
                df.loc[i, "Status_calibracao"] = "SEM_DATA_VALIDACAO"
            else:
                dias = (validade - hoje).days
                if dias < 0:
                    status = "CALIBRACAO_VENCIDA"
                elif dias <= aviso:
                    status = "CALIBRACAO_A_VENCER"
                else:
                    status = "OK"
                df.loc[i, "Status_calibracao"] = status

                if status in ("CALIBRACAO_VENCIDA","CALIBRACAO_A_VENCER") and not ja_cal and email_para:
                    if status == "CALIBRACAO_VENCIDA":
                        assunto = f"[ALERTA] {maq} / {device}: CALIBRAÇÃO VENCIDA (venceu {validade.strftime('%d/%m/%Y')})"
                        extra = f"Dias em atraso: {abs(dias)}"
                    else:
                        assunto = f"[ALERTA] {maq} / {device}: calibração a vencer em {dias} dia(s) (validade {validade.strftime('%d/%m/%Y')})"
                        extra = f"Dias restantes: {dias}"
                    corpo = f"""
                    <p>Olá,</p>
                    <p>A calibração do <b>{device}</b> da máquina <b>{maq}</b> exige atenção.</p>
                    <ul>
                      <li><b>Validade:</b> {validade.strftime('%d/%m/%Y')}</li>
                      <li>{extra}</li>
                      <li><b>Código:</b> {row.get("Codigo") or ""}</li>
                      <li><b>Técnico/WWID:</b> {row.get("Tecnico_nome") or ""} / {row.get("Tecnico_WWID") or ""}</li>
                      <li><b>Certificado:</b> {row.get("Certificado_calibracao") or ""}</li>
                    </ul>
                    <p>— Alerta automático</p>
                    """
                    enviar_email_outlook(email_para, email_cc, assunto, corpo)
                    df.loc[i, "Alerta_calibracao_enviado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    enviados_cal += 1

        xls[device] = df

    # Save all back and update Principal
    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as xw:
        for device in SHEETS_DEVICES:
            xls[device].to_excel(xw, index=False, sheet_name=device)
        principal = pd.DataFrame([{"Maquina":"", "Ultima_atualizacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}])
        principal.to_excel(xw, index=False, sheet_name="Principal")

    print(f"Concluído. E-mails enviados - Decimais: {enviados_dec} | Calibração: {enviados_cal}")

if __name__ == "__main__":
    run()

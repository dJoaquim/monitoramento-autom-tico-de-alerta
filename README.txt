
SISTEMA DE ALERTA — PIROMETRO / TEMPORIZADOR / DINOMETRO

Arquivos principais:
- monitor_multidispositivo.xlsx  -> planilha com abas: Principal, Cadastro, Pirometro, Temporizador, Dinometro
- sistema_monitor_dispositivos.py -> GUI (tela principal) para cadastrar/ajustar e enviar e-mail (até 6 pessoas)
- monitor_diario.py              -> verificação automática (decimais divergentes + calibração a vencer/vencida)
- start_gui.bat                  -> abre a GUI (instala libs e executa)
- run_diario.bat                 -> roda o monitor diário (instala libs e executa)
- requirements.txt               -> dependências (pandas, openpyxl, pywin32)

COMO USAR
1) Abra a planilha 'monitor_multidispositivo.xlsx':
   - Aba 'Cadastro': defina Decimais_esperados por Máquina+Dispositivo, e e-mails padrão (gestor/cópia) e Avisar_dias_antes.
   - Aba 'Principal': apenas informativo (última atualização).
   - Abas 'Pirometro', 'Temporizador', 'Dinometro': registros de trocas/ajustes.

2) GUI — Cadastro/Ajuste com envio a 6 pessoas:
   - Execute 'start_gui.bat' (Windows): abrirá a tela principal.
   - Informe a Máquina no topo (opcional) e utilize as abas para registrar.
   - Clique 'Registrar + Enviar e-mail (ajuste)' para disparar um e-mail para até 6 destinatários informados nos campos.
   - O sistema salva o registro na aba correspondente e marca 'Ajuste_email_enviado_em'.

3) Monitor Diário — execução a cada 24h:
   - Execute 'run_diario.bat' manualmente ou agende no Agendador de Tarefas do Windows para rodar 1x ao dia.
   - O script compara quantas casas decimais há na 'Amostra_leitura' vs 'Decimais_esperados' (Cadastro) e envia alerta se diferente.
   - Verifica 'Data_vencimento' e envia alerta quando faltar <= Avisar_dias_antes ou se já estiver vencido.
   - Flags 'Alerta_decimais_enviado_em' e 'Alerta_calibracao_enviado_em' evitam e-mails duplicados.

REQUISITOS
- Windows + Outlook instalado/logado
- Python 3 no PATH
- Bibliotecas: pandas, openpyxl, pywin32 (instaladas automaticamente pelos .bat)

DICAS
- Separe máquinas e responsáveis no 'Cadastro' — permite roteamento de e-mails por área.
- Use ponto OU vírgula na 'Amostra_leitura' (ambos são aceitos para contar decimais).
- Campos de técnico e WWID são gravados e repassados nos e-mails.
- Você pode anexar certificado à mensagem customizando o 'monitor_diario.py' (campo 'Certificado_calibracao').

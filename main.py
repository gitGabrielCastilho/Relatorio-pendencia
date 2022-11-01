import time
import schedule
import win32com.client as win32
import dateutil.utils
import pandas as pd
import fdb



def gerar_pendencia():
    dst_path = r'C:/Users/Gabriel/Desktop/MSYSDADOS.FDB'
    excel_path = r'C:/Users/Gabriel/Desktop/pendencia.xlsx'
    a = dateutil.utils.today()

    print(a)
    TABLE_NAME = 'RECEBER_TITULOS'
    TABLE_NAME1 = 'CLIENTES'

    SELECT = 'select REC_PEDIDO, REC_VALOR, REC_REP_CODIGO, REC_PLA_CAIXA, REC_DATA, ' \
             'REC_VALORPAGO, REC_VENCIMENTO, REC_CLI_CODIGO ' \
             'from %s WHERE (REC_VALORPAGO = 0) ORDER BY REC_DATA' % (TABLE_NAME)

    SELECT1 = 'select CLI_CODIGO, CLI_NOME from %s' % (TABLE_NAME1)

    con = fdb.connect(dsn=dst_path, user='SYSDBA', password='masterkey', charset='UTF8')

    cur = con.cursor()
    cur.execute(SELECT)

    table_rows = cur.fetchall()

    df = pd.DataFrame(table_rows)

    for y in df.loc[2]:
        df[2] = df[2].replace([1,2,3,4,5,6,7],["Leid","Castilho","Loja","Site","Samuel", "Chico", "Zefs"])


    for x in df.loc[3]:
         df[3] = df[3].replace([4,5,6,7,8,9,10,22], ['Dinheiro', 'Cheque','Boleto','Cr crédito',
                                                     'Cr débito','Pix','Depósito','Ajuste Saldo'])


    dft = df.loc[(df[6] <= a)]
    dft = dft.drop(columns=(5))
    dft = dft.rename(columns={0: 'PEDIDO', 1: 'VALOR',2: 'VENDEDOR', 3: 'FORMA PAGAMENTO',
                              4: 'DATA PEDIDO', 6: 'DATA VENCIMENTO',7: 'COD CLIENTE'})


    cur.execute(SELECT1)

    table_rows2 = cur.fetchall()

    dfx = pd.DataFrame(table_rows2)

    dfx = dfx.rename(columns={0:'COD CLIENTE',1:'NOME'})


    m = pd.merge(dft,dfx, how='inner', on='COD CLIENTE')
    m = m.sort_values(by=['FORMA PAGAMENTO','VENDEDOR','DATA VENCIMENTO'])
    m = m.drop(columns=('COD CLIENTE'))

    print(m.head(10))

    m.to_excel(excel_path, index=False)


    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)

    email.To = 'gabrielcastilho111@gmail.com'
    email.Subject = 'Email automático pendências'
    email.HTMLBody = f"""
    <p>Olá financeiro,<p> 
    
    <p>isso é um código automático,<p>
    
    <p>Favor não responder, segue anexo pendências do dia<p>
    """

    email.Attachments.Add(excel_path)

    email.Send()

    print("email enviado")


schedule.every().day.at("08:00").do(gerar_pendencia)

while True:
    schedule.run_pending()
    time.sleep(1)
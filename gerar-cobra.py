import dateutil.utils
import pandas as pd
import fdb

dst_path = r'MTK:C:/Microsys/MsysIndustrial/Dados/MSYSDADOS.FDB'
excel_path = r'C:/Users/Gabriel/Desktop/pendencia.xlsx'
a = dateutil.utils.today()

TABLE_NAME = 'RECEBER_TITULOS'
TABLE_NAME1 = 'CLIENTES'

SELECT = 'select REC_PEDIDO, REC_VALOR, REC_REP_CODIGO, REC_PLA_CAIXA, REC_DATA, ' \
         'REC_VALORPAGO, REC_VENCIMENTO, REC_CLI_CODIGO, REC_SALDO ' \
         'from %s WHERE (REC_VALORPAGO = 0 OR REC_SALDO > 0) ORDER BY REC_DATA' % (TABLE_NAME)

SELECT1 = 'select CLI_CODIGO, CLI_NOME from %s' % (TABLE_NAME1)

con = fdb.connect(dsn=dst_path, user='SYSDBA', password='masterkey', charset='UTF8')

cur = con.cursor()
cur.execute(SELECT)

table_rows = cur.fetchall()

df = pd.DataFrame(table_rows)

for y in df.loc[2]:
    df[2] = df[2].replace([1,2,3,4,5,6,7,12,14,15],["Leid","Castilho","Loja","Site","Samuel", "Chico", "Zefs",
                                              "Michele", "Isabelly", "Alicia"])


for x in df.loc[3]:
     df[3] = df[3].replace([4,5,6,7,8,9,10,22], ['Dinheiro', 'Cheque','Boleto','Cr crédito',
                                                 'Cr débito','Pix','Depósito','Ajuste Saldo'])


dft = df.loc[(df[6] <= a)]
dft = dft.drop(columns=(5))
dft = dft.rename(columns={0: 'PEDIDO', 1: 'VALOR',2: 'VENDEDOR', 3: 'FORMA PAGAMENTO',
                          4: 'DATA PEDIDO', 6: 'DATA VENCIMENTO',7: 'COD CLIENTE', 8:'SALDO'})


cur.execute(SELECT1)

table_rows2 = cur.fetchall()

dfx = pd.DataFrame(table_rows2)

dfx = dfx.rename(columns={0:'COD CLIENTE',1:'NOME'})


m = pd.merge(dft,dfx, how='inner', on='COD CLIENTE')
m = m.sort_values(by=['FORMA PAGAMENTO','VENDEDOR','DATA VENCIMENTO'])
m = m.drop(columns=('COD CLIENTE'))


def clicar():
    m.to_excel(excel_path, index=False)

clicar()




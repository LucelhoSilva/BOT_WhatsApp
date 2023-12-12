import pandas as pd

df = pd.read_excel('Contatos.xlsx')


df['Telefones'] = df['Telefones'].str.replace(r'\s+', '', regex=True)     # Remove espaços em branco
df['Telefones'] = df['Telefones'].str.replace(r'[^0-9]', '', regex=True)  # Remove caracteres não numéricos

df_sem_duplicatas = df.drop_duplicates(subset='Telefones')

def categorizar_pais(numero):
    if numero.startswith('55'):
        return 'Brasil'
    else:
        return 'Exterior'

df['Pais'] = df['Telefones'].apply(lambda x: categorizar_pais(x))

df_brasil = df[df['Pais'] == 'Brasil']
df_exterior = df[df['Pais'] == 'Exterior']

df_brasil.to_excel('Contatos_Brasil.xlsx', index=False)
df_exterior.to_excel('Contatos_Exterior.xlsx', index=False)

import win32com.client as win32 
import pythoncom
import datetime 
import pandas as pd 

def enviar_email(email_destinos, caminho_pasta, executivos, nome):
    data_atual = datetime.datetime.now() 
    dia_semana = data_atual.strftime('%A')
    print(f'Hoje é dia: {dia_semana}')
    hora_atual = data_atual.hour 
    if 6 <= hora_atual < 12:
        cumprimento = 'bom dia!'
    elif 12 <= hora_atual < 18:
        cumprimento = 'boa tarde!'
    else:
        cumprimento = 'boa noite!'
    pythoncom.CoInitialize() 
    outlook = win32.Dispatch('outlook.application')
    
    for email_destino, executivo, nome in zip(email_destinos, executivos, nome):
        anexo1 = caminho_pasta + '\\' + executivo + ' - RESUMO - ' + data_atual.strftime('%d-%m-%Y') + '.pdf'
        anexo2 = caminho_pasta + '\\' + executivo + ' - DETALHE - ' + data_atual.strftime('%d-%m-%Y') + '.pdf'
        if dia_semana.lower() == 'monday':
            lista_anexo = [anexo1, anexo2]
        else:
            lista_anexo = [anexo1]
        if email_destino == '-':    
            print(f'Endereço de email do executivo: {executivo} é inválido.') 
            continue 
        email = outlook.CreateItem(0)
        email.To = email_destino 
        email.Subject = f'Relatório de vendas - {data_atual.strftime('%d/%m')}'
        email.HTMLBody = f"""
        <p>{nome}, {cumprimento}</p>
        <p>Segue em anexo o relatório de vendas.</p>
        <p>At.te,</p>
        <p>Wuesley</p>
        """
        try:
            for anexo in lista_anexo:
                email.Attachments.Add(anexo)
            # email.Display()
            email.Send()
        except Exception as e:
            print(f'Erro ao anexar o arquivo do executivo: {executivo}. {e}')
            print(f'Arquivos: {anexo}')
        print('Email enviado')

executivos_dict = {
    'CD_EXECUTIVO': [3707391, 195949, 3342139, 4769, 329796, 3450235, 3669259, 291545, 1996982, 667920, 1203394, 3317282, 3028217, 4794, 4696, 680155, 0, 3844360, 3978814],
    'EXECUTIVO': ['ANDRÉ FERREIRA SOUSA LEMES', 'AUGUSTO CESAR RIBEIRO', 'CELSO RICARDO BARROS DE SOUZA', 'CLEITON DA SILVA NASCIMENTO', 'EMMANUEL DA SILVA ARAUJO', 'FABIO NOQUELE', 'FORÇA DE VENDAS INTERNA', 
                  'GERLEIDE BARROZO MEIRA', 'GILSON ALDO MEIRA', 'ISRAEL FRANKLIN GARCIA DE LIMA', 'JOABE MENEZES DA SILVA', 'MARIA EDUARDA MOREIRA DA SILVA', 'RAFAEL BRUNO RODRIGUES', 'RONALDO ALVES CONDE', 'SILVESTRE CIRILO SOBRINHO', 
                  'TIAGO DA CONCEIÇÃO SANTOS', 'VENDA INTERNA', 'RAFAEL AUGUSTO MAGIERA RIBEIRO', 'VINICIUS CARRA'],
    'NOME': ['André',' Augusto', 'Celso','Cleiton', 'Emmanuel', 'Fabio', '-', 'Gerleide', 'Gilson', 'Israel', 'Joabe', 'Maria Eduarda', 'Rafael', 'Ronaldo', 'Silvestre', 'Tiago', '-', 'Rafael', 'Vinicius'],
    'EMAIL': ['representantenorte03@laboratoriosodre.com.br', 'regionalnordeste@laboratoriosodre.com.br', 'celso.souza@laboratoriosodre.com.br', 'cleiton.nascimento@laboratoriosodre.com.br', 
              'emmanuel.araujo@laboratoriosodre.com.br', 'fabio.noquele@laboratoriosodre.com.br', '-', 'gerleide.meira@laboratoriosodre.com.br', 'gilson.meira@laboratoriosodre.com.br', 
              'israel.lima@laboratoriosodre.com.br', 'joabe.silva@laboratoriosodre.com.br', 'maria.silva@laboratoriosodre.com.br', 'rafael.rodrigues@laboratoriosodre.com.br', 'ronaldo.alves@laboratoriosodre.com.br', 
              'silvestre.cirilo@laboratoriosodre.com.br', 'tiago.conceicao@laboratoriosodre.com.br', '-', 'rafael.ribeiro@laboratoriosodre.com.br', 'vinicius.carra@laboratoriosodre.com.br']
}
dic = pd.DataFrame(executivos_dict)
executivos_para_enviar = ['RONALDO ALVES CONDE', 'RAFAEL AUGUSTO MAGIEIRA RIBEIRO', 'RAFAEL BRUNO RODRIGUES', 'SILVESTRE CIRILO SOBRINHO', 'FABIO NOQUELE', ' VINICIUS CARRA' ]
filtro_executivos = dic[dic['EXECUTIVO'].isin(executivos_para_enviar)]
lista_executivos = filtro_executivos['EXECUTIVO'].tolist()
lista_email = filtro_executivos['EMAIL'].tolist()
nome_executivos = filtro_executivos['NOME'].tolist()
pasta = r'C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Felipe Corradi\Relatórios Executivos\Relatório Segunda-Feira'
enviar_email(lista_email, pasta, lista_executivos, nome_executivos)

import os
import pandas as pd
import matplotlib.pyplot as plt
from sqlalchemy import create_engine
import datetime
from dateutil.relativedelta import relativedelta
import locale
import calendar
from matplotlib.backends.backend_pdf import PdfPages
import time

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

usuario = 'joao.ruiz'
senha = 'pSLOCgBgy5tB'
servidor = 'sodreproducao.public.608eda075a62.database.windows.net,3342'
banco_de_dados = 'toxicologico'
conexao_sql = f'mssql+pyodbc://{usuario}:{senha}@{servidor}/{banco_de_dados}?driver=ODBC+Driver+17+for+SQL+Server'

consulta_sql = """
DECLARE @DiaAtual INT = DAY(GETDATE());

DECLARE @MesPassado DATE;
DECLARE @MesAtual DATE;

IF @DiaAtual = 1
BEGIN    
    SET @MesPassado = CONVERT(DATE, CONCAT(MONTH(DATEADD(MONTH, -2, GETDATE())), '/', 01, '/', YEAR(GETDATE())));
    SET @MesAtual = EOMONTH(DATEADD(MONTH, -1, GETDATE()))
END
ELSE
BEGIN
    SET @MesPassado = CONVERT(DATE, CONCAT(MONTH(DATEADD(MONTH, -1, GETDATE())), '/', 01, '/', YEAR(GETDATE())));
    SET @MesAtual = EOMONTH(GETDATE())
END

SELECT 
    CASE 
        WHEN v.cd_empresa IS NULL 
        THEN v.cd_laboratorio 
        ELSE v.cd_empresa 
    END AS CD_ENTIDADE,
    e.DS_FANTASIA, 
    CAST(CONCAT(MONTH(v.dt_venda), '/', 01, '/', YEAR(GETDATE())) AS DATE) AS DATA,
    COUNT(v.ds_etiqueta) AS VOLUME, 
    CASE 
        WHEN DATEDIFF(DAY, v.dt_venda, GETDATE()) <= 90
            AND e.dt_cadastro >= '2022-12-31' 
            AND v.ds_finalidade <> 'TERCEIRO'
        THEN 'VENDA NOVA'
        ELSE 'VENDA MANUTENÇÃO'
    END AS STATUS_VENDA,
    f.ds_fantasia AS EXECUTIVO
FROM powerbi.consultavendas v 
LEFT JOIN morales.dbo.tbl_entidades e 
    ON (CASE 
            WHEN v.cd_empresa IS NULL 
            THEN v.cd_laboratorio 
            ELSE v.cd_empresa
        END) = e.cd_entidade
LEFT JOIN powerbi.executivos f 
    ON e.cd_vendedor = f.cd_vendedor 
WHERE v.dt_venda BETWEEN @MesPassado AND @MesAtual
GROUP BY 
    CASE 
        WHEN v.cd_empresa IS NULL 
        THEN v.cd_laboratorio 
        ELSE v.cd_empresa 
    END, 
    e.DS_FANTASIA,
    CAST(CONCAT(MONTH(v.dt_venda), '/', 01, '/', YEAR(GETDATE())) AS DATE),
    f.ds_fantasia,
    CASE 
        WHEN DATEDIFF(DAY, v.dt_venda, GETDATE()) <= 90 
            AND e.dt_cadastro >= '2022-12-31'
            AND v.ds_finalidade <> 'TERCEIRO'
        THEN 'VENDA NOVA'
        ELSE 'VENDA MANUTENÇÃO'
    END
ORDER BY 
    COUNT(v.ds_etiqueta) DESC
"""

def consulta_sql_dados(conexao_sql, consulta_sql):
    try:
        engine = create_engine(conexao_sql)
        df = pd.read_sql(consulta_sql, engine)
        return df
    except Exception as e:
        print(f'Ocorreu um erro ao executar a consulta SQL: {e}')
        return None

def contar_dias_uteis():
    ano_atual, mes_atual, dia_atual = time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday
    ano_passado, mes_passado = obter_mes_passado()
    dias_uteis_mes_passado = contar_dias_uteis_mes(ano_passado, mes_passado)
    dias_uteis_mes_atual = contar_dias_uteis_mes(ano_atual, mes_atual, dia_limite=dia_atual)

    return dias_uteis_mes_passado, dias_uteis_mes_atual

def obter_mes_passado():
    ano_atual, mes_atual = time.localtime().tm_year, time.localtime().tm_mon
    if mes_atual == 1:
        return ano_atual - 1, 12
    else:
        return ano_atual, mes_atual - 1
    
def contar_dias_uteis_mes(ano, mes, dia_limite=None):
    if dia_limite is None:
        _, ultimo_dia_mes = calendar.monthrange(ano, mes)
        dia_limite = ultimo_dia_mes
    
    dias_uteis = 0
    for dia in range(1, dia_limite + 1):
        dia_semana = calendar.weekday(ano, mes, dia)
        if dia_semana < 5:
            dias_uteis += 1
            
    return dias_uteis

dias_uteis_anterior, dias_uteis_atual = contar_dias_uteis()
day_py = datetime.datetime.now().day
if day_py == 1:
    mes_passado_py = (datetime.datetime.now() - relativedelta(months=2)).strftime('%B/%y')
    mes_atual_py = (datetime.datetime.now() - relativedelta(months=1)).strftime('%B/%y')
else:
    mes_passado_py = (datetime.datetime.now() - relativedelta(months=1)).strftime('%B/%y')
    mes_atual_py = datetime.datetime.now().strftime('%B/%y')

def gerar_pdf(df_geral, nome_arquivo_pdf):
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.axis('tight')
    ax.axis('off')
    tabela = ax.table(cellText=df_geral.values, cellLoc='center', loc='center')
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(10)
    tabela.scale(1.2, 1.2)
    try:
        with PdfPages(nome_arquivo_pdf) as pdf:
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
        print(f"PDF salvo como: {nome_arquivo_pdf}")
    except Exception as e:
        print(f"Erro ao salvar PDF: {e}")

def gerar_relatorios_executivo(df, pasta_destino):
    executivos = df['EXECUTIVO'].unique()
    for executivo in executivos:
        df_executivo = df[df['EXECUTIVO'] == executivo]    
        try:
            vendas_novas_lm = df_executivo[(df_executivo['STATUS_VENDA'] == 'VENDA NOVA') & 
                                           (df_executivo['DATA'].dt.strftime('%B/%y') == mes_passado_py)]['VOLUME'].sum()
            vendas_novas_gd = df_executivo[(df_executivo['STATUS_VENDA'] == 'VENDA NOVA') & 
                                           (df_executivo['DATA'].dt.strftime('%B/%y') == mes_atual_py)]['VOLUME'].sum()
            vendas_manutencao_lm = df_executivo[(df_executivo['STATUS_VENDA'] == 'VENDA MANUTENÇÃO') & 
                                                (df_executivo['DATA'].dt.strftime('%B/%y') == mes_passado_py)]['VOLUME'].sum()
            vendas_manutencao_gd = df_executivo[(df_executivo['STATUS_VENDA'] == 'VENDA MANUTENÇÃO') & 
                                                (df_executivo['DATA'].dt.strftime('%B/%y') == mes_atual_py)]['VOLUME'].sum()
            med_vnd_lm = (vendas_manutencao_lm + vendas_novas_lm) / dias_uteis_anterior
            med_vnd_gd = (vendas_novas_gd + vendas_manutencao_gd) / dias_uteis_atual
            media_venda = (vendas_manutencao_gd + vendas_novas_gd) / day_py
            dias_no_mes_atual = calendar.monthrange(datetime.datetime.now().year, datetime.datetime.now().month)[1]
            prj = media_venda * dias_no_mes_atual
            
        except KeyError as e:
            print(f"Erro ao acessar as colunas: {e}")

        data = {
            'A': [
                f"VENDAS NOVAS {mes_passado_py.upper()}",
                f"{round(vendas_novas_lm):,.0f}".replace(',', '.'),
                "",
                f"VENDAS MANUTENÇÃO {mes_passado_py.upper()}",
                f"{round(vendas_manutencao_lm):,.0f}".replace(',', '.'),
                "",
                f"MÉDIA AMOSTRAS DIAS ÚTEIS {mes_passado_py.upper()}",
                f"{round(med_vnd_lm):,.0f}".replace(',', '.')
            ],
            'B': [
                f"VENDAS NOVAS {mes_atual_py.upper()}",
                f"{round(vendas_novas_gd):,.0f}".replace(',', '.'),
                "",
                f"VENDAS MANUTENÇÃO {mes_atual_py.upper()}",
                f"{round(vendas_manutencao_gd):,.0f}".replace(',', '.'),
                "",
                f"MÉDIA AMOSTRAS DIAS ÚTEIS {mes_atual_py.upper()}",
                f"{round(med_vnd_gd):,.0f}".replace(',', '.')
            ],
            'C': [
                f"PROJEÇÃO VENDAS {mes_atual_py.upper()}", 
                f"{round(prj):,.0f}".replace(',', '.'),
                "",
                "",
                "",   
                "",
                "",
                ""
        ]}
        
        df_geral = pd.DataFrame(data)
        data_atual = datetime.datetime.now().strftime('%d-%m-%Y')
        nome_arquivo_excel = f'{executivo} - RESUMO - {data_atual}.xlsx'
        caminho_arquivo_excel = os.path.join(pasta_destino, nome_arquivo_excel)
        try:
            df_geral.to_excel(caminho_arquivo_excel, index=False, header=False)
            print(f"Arquivo Excel salvo para {executivo}: {caminho_arquivo_excel}")
        except Exception as e:
            print(f"Erro ao salvar arquivo Excel para {executivo}: {e}")
        nome_arquivo_pdf = f'{executivo} - RESUMO - {data_atual}.pdf'
        caminho_pdf = os.path.join(pasta_destino, nome_arquivo_pdf)
        try:
            gerar_pdf(df_geral, caminho_pdf)
            print(f"PDF salvo para {executivo}: {caminho_pdf}")
        except Exception as e:
            print(f"Erro ao gerar PDF para {executivo}: {e}")

df = consulta_sql_dados(conexao_sql, consulta_sql)

if df is not None:
    df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce')  # 'errors=coerce' lida com valores inválidos
    pasta_destino = r'C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Felipe Corradi\Relatórios Executivos\Relatório Segunda-Feira'
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    gerar_relatorios_executivo(df, pasta_destino)
else:
    print("Erro ao carregar os dados da consulta SQL.")
print('Código executado 👀')

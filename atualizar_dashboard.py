"""
Atualizar Dashboard - Brasil Botoes
=====================================
O que este script faz, passo a passo:
  1. Abre a planilha Acompanhamento_Clientes_-_Geral.xlsm no Excel (invisivel)
  2. Manda o Excel atualizar as consultas ODBC (o mesmo que clicar em "Atualizar Tudo")
  3. Le os dados atualizados das abas "Vendas" e "Amostras"
  4. Salva tudo em um arquivo data.json dentro da pasta do dashboard
  5. Fecha o Excel

Depois disso, o "publicar.bat" envia o data.json para o GitHub, atualizando o site.

Pre-requisitos (rodar UMA vez no terminal/prompt de comando):
    pip install xlwings openpyxl

Configuracao: ajuste as 2 linhas em "CONFIGURACOES" abaixo.
"""

import json
import datetime
import sys
from pathlib import Path

import xlwings as xw

# ======================= CONFIGURACOES =======================
# Caminho completo da planilha original (a que tem os ODBCs)
CAMINHO_PLANILHA = r"C:\Users\SEU_USUARIO\Documents\Acompanhamento_Clientes_-_Geral.xlsm"

# Pasta onde fica o dashboard (index.html + data.json), dentro do seu repositorio git
PASTA_DASHBOARD = r"C:\Users\SEU_USUARIO\Documents\dashboard-brasil-botoes"
# ===============================================================

COLUNAS = [
    "Situação do pedido", "Cod do Cliente", "Cliente", "Cod do Repres",
    "Representante", "Pedido Consistem", "Pedido Multiplier", "Pedido Cliente",
    "Data Emissão Pedido", "Data Previsão Faturamento", "Data Faturamento",
    "Tipo de NF", "Valor Nota Fiscal", "Valor Pedido", "Valor Saldo Pedido",
    "Observações do Pedido",
]


def formatar_data(valor):
    """Converte datas do Excel para texto AAAA-MM-DD (ou None se vazio)."""
    if isinstance(valor, datetime.datetime):
        return valor.strftime("%Y-%m-%d")
    if isinstance(valor, datetime.date):
        return valor.strftime("%Y-%m-%d")
    if valor in (None, ""):
        return None
    return str(valor)


def extrair_aba(sheet):
    """Le a tabela de uma aba (Vendas ou Amostras) a partir da linha do cabecalho."""
    usado = sheet.used_range
    valores = usado.value  # lista de listas

    linha_cabecalho = None
    for i, linha in enumerate(valores):
        if linha and linha[0] == "Situação do pedido":
            linha_cabecalho = i
            break
    if linha_cabecalho is None:
        raise RuntimeError(f"Não encontrei o cabeçalho na aba {sheet.name}")

    cabecalho = valores[linha_cabecalho]
    registros = []
    for linha in valores[linha_cabecalho + 1:]:
        if not linha or (linha[0] is None and linha[2] is None):
            continue
        registro = {}
        for col_nome, val in zip(cabecalho, linha):
            if col_nome is None:
                continue
            if "Data" in str(col_nome):
                val = formatar_data(val)
            registro[col_nome] = val
        registros.append(registro)
    return registros


def main():
    planilha_path = Path(CAMINHO_PLANILHA)
    saida_path = Path(PASTA_DASHBOARD) / "data.json"

    if not planilha_path.exists():
        print(f"ERRO: planilha não encontrada em {planilha_path}")
        sys.exit(1)

    print("Abrindo Excel (invisível)...")
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(str(planilha_path), update_links=False, read_only=False)

        print("Atualizando consultas ODBC (Atualizar Tudo)... isso pode levar alguns minutos.")
        wb.api.RefreshAll()
        app.api.CalculateUntilAsyncQueriesDone()  # espera as consultas terminarem

        print("Lendo dados atualizados...")
        vendas = extrair_aba(wb.sheets["Vendas"])
        amostras = extrair_aba(wb.sheets["Amostras"])

        print(f"  Vendas: {len(vendas)} pedidos | Amostras: {len(amostras)} pedidos")

        wb.save()
        wb.close()

    finally:
        app.quit()

    saida_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "gerado_em": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "vendas": vendas,
        "amostras": amostras,
    }
    with open(saida_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)

    print(f"OK! data.json atualizado em: {saida_path}")


if __name__ == "__main__":
    main()

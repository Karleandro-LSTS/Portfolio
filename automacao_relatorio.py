"""
Exemplo de automação: gera um resumo de vendas a partir de uma planilha.

Uso:
    python exemplo_automacao_relatorio.py vendas.xlsx

O script lê a planilha, calcula totais por vendedor e exporta um novo arquivo
com o resumo. Serve como modelo para automatizar relatórios recorrentes.
"""

import sys
from pathlib import Path

import pandas as pd


def gerar_resumo(caminho_entrada: Path) -> Path:
    df = pd.read_excel(caminho_entrada)

    # Validação básica das colunas esperadas
    colunas_necessarias = {"vendedor", "valor"}
    faltando = colunas_necessarias - set(df.columns)
    if faltando:
        raise ValueError(f"Colunas ausentes na planilha: {', '.join(sorted(faltando))}")

    # Remove linhas sem valor e garante tipo numérico
    df = df.dropna(subset=["valor"])
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df = df.dropna(subset=["valor"])

    resumo = (
        df.groupby("vendedor", as_index=False)
          .agg(total_vendas=("valor", "sum"), quantidade=("valor", "count"))
          .sort_values("total_vendas", ascending=False)
    )
    resumo["ticket_medio"] = (resumo["total_vendas"] / resumo["quantidade"]).round(2)

    caminho_saida = caminho_entrada.with_name("resumo_vendas.xlsx")
    resumo.to_excel(caminho_saida, index=False)
    return caminho_saida


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python exemplo_automacao_relatorio.py <arquivo.xlsx>")
        sys.exit(1)

    saida = gerar_resumo(Path(sys.argv[1]))
    print(f"Relatório gerado em: {saida}")

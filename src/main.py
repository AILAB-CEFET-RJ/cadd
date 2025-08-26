#!/usr/bin/env python3
"""
Script para executar em sequência os módulos:
- main_faixas_criticidade
- main_integralizacoes
- main_reprovacoes

Uso:
------
python executar_modulos.py -p 2017.2 -i ../data -o ../resultados

Parâmetros:
  -p, --periodo     Período a ser processado (ex: 2017.2)
  -i, --input       Caminho para os dados de entrada
  -o, --output      Caminho para salvar os resultados
"""

import argparse
import main_faixas_criticidade
import main_integralizacoes
import main_reprovacoes

def main():
    parser = argparse.ArgumentParser(
        description="Executa os módulos de análise (faixas de criticidade, integralizações e reprovações)."
    )
    parser.add_argument("-p", "--periodo", required=True, help="Período a ser processado (ex: 2017.2)")
    parser.add_argument("-i", "--input", required=True, help="Diretório com dados de entrada")
    parser.add_argument("-o", "--output", required=True, help="Diretório para salvar os resultados")

    args = parser.parse_args()

    # Constrói a lista de argumentos no formato esperado pelos módulos
    module_args = [
        "-p", args.periodo,
        "-i", args.input,
        "-o", args.output
    ]

    # Executa os três módulos
    main_faixas_criticidade.main(module_args)
    main_integralizacoes.main(module_args)
    main_reprovacoes.main(module_args)

if __name__ == "__main__":
    main()

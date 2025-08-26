# coding=UTF-8

import collections
import glob, os
import sys, getopt

import plotly.graph_objs as go
from plotly.offline import plot

from os.path import basename

from analisador_desempenho_academico import AnalisadorDesempenhoAcademico
from parametros import Parametros
import xlsxwriter

from periodo_letivo import PeriodoLetivo


def plot_agregado_integralizacoes(sigla_curso,
                                  outputdir,
                                  lista_qtd_periodos,
                                  lista_qtd_alunos_por_qtd_periodos_cursados,
                                  qtd_total_alunos,
                                  qtd_alunos_irregulares):
    trace_quantidades = go.Bar(
        x=lista_qtd_periodos,
        y=lista_qtd_alunos_por_qtd_periodos_cursados,
        name='# periodos cursados',
        marker=dict(color='rgb(55, 83, 109)')
    )
    data_quantidades = [trace_quantidades]

    layout_quantidades = go.Layout(
        title=(
            '<b>Gráfico: "Quantidades de alunos" versus "Qtd. de períodos cursados"</b>'
            f'<br>Curso: {sigla_curso}'
            f'<br>Período base: {Parametros.ANO_BASE}.{Parametros.PERIODO_BASE}'
            f'<br>Total/irregulares: {qtd_total_alunos}/{qtd_alunos_irregulares}'
        ),
        xaxis=dict(
            tickfont=dict(size=14, color='rgb(107, 107, 107)')
        ),
        yaxis=dict(
            title='qtd de alunos',
            titlefont=dict(size=16, color='rgb(107, 107, 107)'),
            tickfont=dict(size=14, color='rgb(107, 107, 107)')
        ),
        legend=dict(
            x=0,
            y=1.0,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        )
    )
    fig = go.Figure(data=data_quantidades, layout=layout_quantidades)
    filename = os.path.join(outputdir, f"{sigla_curso}-periodos-cursados.html")
    plot(fig, filename=filename, auto_open=False)


def processar_integralizacoes(planilha, outputdir):
    sigla_curso = os.path.splitext(basename(planilha))[0]

    mapa_alunos = AnalisadorDesempenhoAcademico.construir_mapa_alunos(planilha)
    mapas = AnalisadorDesempenhoAcademico.construir_mapas_periodos(planilha, mapa_alunos)
    mapa_periodos_cursados_por_aluno = mapas[0]
    mapa_trancamentos_totais_por_aluno = mapas[1]

    # Create a workbook and add a worksheet.
    filename = os.path.splitext(planilha)[0] + '-periodos-cursados.xlsx'
    filename = os.path.join(outputdir, basename(filename))
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Detalhado')

    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'matricula_aluno', bold)
    worksheet.write('B1', 'nome_aluno', bold)
    worksheet.write('C1', 'qtd_periodos_cursados', bold)
    worksheet.write('D1', 'qtd_trancamentos', bold)

    row = 1
    col = 0
    width_coluna_nome = 0

    for matr_aluno, qtd_periodos_cursados in mapa_periodos_cursados_por_aluno.items():
        qtd_trancamentos = mapa_trancamentos_totais_por_aluno.get(matr_aluno, 0)

        worksheet.write(row, col, matr_aluno)

        aluno = mapa_alunos[matr_aluno]
        worksheet.write(row, col + 1, aluno.nome)
        width_coluna_nome = max(width_coluna_nome, len(aluno.nome))

        worksheet.write(row, col + 2, qtd_periodos_cursados)
        worksheet.write(row, col + 3, qtd_trancamentos)
        row += 1

    worksheet.set_column(col, col, len('matricula_aluno'))
    worksheet.set_column(col + 1, col + 1, width_coluna_nome)
    worksheet.set_column(col + 2, col + 2, len('qtd_periodos_cursados'))
    worksheet.set_column(col + 3, col + 3, len('qtd_trancamentos'))

    print(f"Quantidade de alunos com matrícula ativa: {len(mapa_periodos_cursados_por_aluno)}")

    mapa_qtd_alunos_por_qtd_periodos_cursados = {}
    for matr_aluno, qtd_periodos_cursados in mapa_periodos_cursados_por_aluno.items():
        mapa_qtd_alunos_por_qtd_periodos_cursados[qtd_periodos_cursados] = \
            mapa_qtd_alunos_por_qtd_periodos_cursados.get(qtd_periodos_cursados, 0) + 1

    qtd_alunos_irregulares = 0
    qtd_total_alunos = 0
    for lista_qtd_periodos, qtd_alunos in mapa_qtd_alunos_por_qtd_periodos_cursados.items():
        qtd_total_alunos += qtd_alunos
        if lista_qtd_periodos > Parametros.mapa_qtd_maxima_periodos_integralizacao[sigla_curso]:
            qtd_alunos_irregulares += qtd_alunos

    print(f"Qtd. total alunos: {qtd_total_alunos}.")
    print(f"Qtd. alunos irregulares: {qtd_alunos_irregulares}.")

    worksheet = workbook.add_worksheet('Agregado')
    worksheet.write('A1', 'qtd_periodos_cursados', bold)
    worksheet.write('B1', 'qtd_alunos', bold)

    row = 1
    col = 0
    for key, value in mapa_qtd_alunos_por_qtd_periodos_cursados.items():
        worksheet.write(row, col, key)
        worksheet.write(row, col + 1, value)
        row += 1

    worksheet.set_column(col, col, len('qtd_periodos_cursados'))
    worksheet.set_column(col + 1, col + 1, len('qtd_alunos'))
    workbook.close()

    mapa_qtd_alunos_por_qtd_periodos_cursados = collections.OrderedDict(
        sorted(mapa_qtd_alunos_por_qtd_periodos_cursados.items())
    )

    lista_qtd_periodos = list(mapa_qtd_alunos_por_qtd_periodos_cursados.keys())
    lista_qtd_alunos_por_qtd_periodos_cursados = list(mapa_qtd_alunos_por_qtd_periodos_cursados.values())

    plot_agregado_integralizacoes(sigla_curso,
                                  outputdir,
                                  lista_qtd_periodos,
                                  lista_qtd_alunos_por_qtd_periodos_cursados,
                                  qtd_total_alunos,
                                  qtd_alunos_irregulares)


def main(argv):
    print(f"Período letivo base: {Parametros.ANO_BASE}/{Parametros.PERIODO_BASE}")

    inputdir = ''
    outputdir = ''

    try:
        opts, args = getopt.getopt(argv, "hp:i:o:", ["idir=", "odir="])
    except getopt.GetoptError:
        print("main.py -i <inputdir> -o <outputdir>")
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-h':
            print("main.py -i <inputdir> -o <outputdir>")
            sys.exit()
        elif opt in ("-i", "--idir"):
            inputdir = arg
        elif opt in ("-o", "--odir"):
            outputdir = arg
        elif opt == "-p":
            componentes = arg.split('.')
            if len(componentes) == 2:
                Parametros.ANO_BASE = int(componentes[0])
                Parametros.PERIODO_BASE = int(componentes[1])
                print(PeriodoLetivo.fromstring(arg))
            else:
                sys.exit()

    print("Input dir:", inputdir)
    print("Output dir:", outputdir)

    os.chdir(inputdir)
    for planilha in glob.glob("*.csv"):
        print(f"Iniciando processamento do arquivo {planilha}...")
        processar_integralizacoes(planilha, outputdir)


if __name__ == "__main__":
    main(sys.argv[1:])

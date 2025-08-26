# coding=UTF-8

import plotly.graph_objs as go
import sys, getopt
import glob, os
import collections
from os.path import basename

from plotly.offline import plot

from periodo_letivo import PeriodoLetivo
from analisador_desempenho_academico import AnalisadorDesempenhoAcademico
import xlsxwriter
from parametros import Parametros


def iniciar_mapa(sigla_curso):
    qtd = Parametros.mapa_qtd_reprovacoes_jubilacao_por_curso[sigla_curso]
    mapa = {}
    for i in range(0, qtd + 1):
        mapa[i] = 0
    return mapa


def processar_reprovacoes(planilha, outputdir):
    mapa_alunos = AnalisadorDesempenhoAcademico.construir_mapa_alunos(planilha)

    mapas = AnalisadorDesempenhoAcademico.construir_mapas_reprovacoes(planilha, mapa_alunos)
    mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina = mapas[0]
    mapa_aluno_e_disciplina_para_qtd_reprovacoes = mapas[1]
    mapa_de_cod_disciplina_para_nome_disciplina = mapas[2]

    sigla_curso = os.path.splitext(basename(planilha))[0]

    mapa_qtd_alunos_por_maximo_reprovacoes_em_uma_disciplina = iniciar_mapa(sigla_curso)

    qtd = Parametros.mapa_qtd_reprovacoes_jubilacao_por_curso[sigla_curso]
    for (matr_aluno, qtd_total_reprovacoes) in mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina.items():
        if qtd_total_reprovacoes < qtd:
            mapa_qtd_alunos_por_maximo_reprovacoes_em_uma_disciplina[qtd_total_reprovacoes] += 1
        else:
            mapa_qtd_alunos_por_maximo_reprovacoes_em_uma_disciplina[qtd] += 1

    filename = os.path.splitext(planilha)[0] + '-reprovacoes.xlsx'
    filename = os.path.join(outputdir, basename(filename))

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Detalhado')

    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'cod_disciplina', bold)
    worksheet.write('B1', 'nome_disciplina', bold)
    worksheet.write('C1', 'matr_aluno', bold)
    worksheet.write('D1', 'nome_aluno', bold)
    worksheet.write('E1', 'qtds_reprovacoes', bold)

    row = 1
    col = 0

    width_coluna_nome = 0
    width_coluna_nome_disciplina = 0

    for key, value in mapa_aluno_e_disciplina_para_qtd_reprovacoes.items():
        partes = key.rsplit(';')
        cod_disciplina = partes[0]
        nome_disciplina = mapa_de_cod_disciplina_para_nome_disciplina[cod_disciplina]
        matr_aluno = partes[1]

        worksheet.write(row, col, cod_disciplina)
        worksheet.write(row, col + 1, nome_disciplina)
        width_coluna_nome_disciplina = max(width_coluna_nome_disciplina, len(nome_disciplina))

        worksheet.write(row, col + 2, matr_aluno)

        aluno = mapa_alunos[matr_aluno]
        worksheet.write(row, col + 3, aluno.nome)
        width_coluna_nome = max(width_coluna_nome, len(aluno.nome))

        worksheet.write(row, col + 4, value)
        row += 1

    worksheet.set_column(col, col, len('cod_disciplina'))
    worksheet.set_column(col + 1, col + 1, width_coluna_nome_disciplina)
    worksheet.set_column(col + 2, col + 2, len('matr_aluno'))
    worksheet.set_column(col + 3, col + 3, width_coluna_nome)
    worksheet.set_column(col + 4, col + 4, len('qtds_reprovacoes'))

    worksheet = workbook.add_worksheet('Agregado')
    worksheet.write('A1', 'qtd_max_reprovacoes', bold)
    worksheet.write('B1', 'qtd_alunos', bold)

    row = 1
    col = 0
    for key, value in mapa_qtd_alunos_por_maximo_reprovacoes_em_uma_disciplina.items():
        worksheet.write(row, col, key)
        worksheet.write(row, col + 1, value)
        row += 1

    worksheet.set_column(col, col, len('qtd_max_reprovacoes'))
    worksheet.set_column(col + 1, col + 1, len('qtd_alunos'))

    workbook.close()

    dados = collections.OrderedDict(sorted(mapa_qtd_alunos_por_maximo_reprovacoes_em_uma_disciplina.items()))

    qtds_reprovacoes = list(dados.keys())
    qtds_alunos_com_aquela_qtd_reprovacoes_em_alguma_disciplina = list(dados.values())

    plotar_grafico(
        sigla_curso,
        outputdir,
        qtds_reprovacoes,
        qtds_alunos_com_aquela_qtd_reprovacoes_em_alguma_disciplina,
    )


def plotar_grafico(sigla_curso, outputdir, lista_qtd_reprovacoes, lista_qtd_alunos_com_aquela_qtd_reprovacoes):
    lista_qtd_reprovacoes_str = ['= ' + str(x) for x in lista_qtd_reprovacoes]
    lista_qtd_reprovacoes_str[-1] = lista_qtd_reprovacoes_str[-1].replace('=', '>=')

    trace_quantidades = go.Bar(
        x=lista_qtd_reprovacoes_str,
        y=lista_qtd_alunos_com_aquela_qtd_reprovacoes,
        name='Reprovações',
        marker=dict(color='rgb(55, 83, 109)'),
    )

    layout_quantidades = go.Layout(
        title=f'Curso: {sigla_curso}\n "Quantidades de alunos" versus "Qtd. máxima de reprovações em alguma disciplina"',
        xaxis=dict(tickfont=dict(size=14, color='rgb(107, 107, 107)')),
        yaxis=dict(
            title='qtd de alunos',
            titlefont=dict(size=16, color='rgb(107, 107, 107)'),
            tickfont=dict(size=14, color='rgb(107, 107, 107)'),
        ),
        legend=dict(x=0, y=1.0, bgcolor='rgba(255, 255, 255, 0)', bordercolor='rgba(255, 255, 255, 0)'),
    )

    fig = go.Figure(data=[trace_quantidades], layout=layout_quantidades)
    plot(fig, filename=os.path.join(outputdir, f"{sigla_curso}-reprovacoes.html"))


def main(argv):
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

    print("Input dir: ", inputdir)
    print("Output dir:", outputdir)

    os.chdir(inputdir)
    for planilha in glob.glob("*.csv"):
        print(f"Iniciando processamento do arquivo {planilha}...")
        processar_reprovacoes(planilha, outputdir)


if __name__ == "__main__":
    main(sys.argv[1:])
    
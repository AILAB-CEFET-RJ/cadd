# coding=UTF-8

import csv
import sys
from aluno import Aluno
from parametros import Parametros
from periodo_letivo import PeriodoLetivo


class AnalisadorDesempenhoAcademico(object):

    @staticmethod
    def construir_mapa_alunos(planilha):
        with open(planilha, encoding="utf-8") as f_obj:
            reader = csv.DictReader(f_obj, delimiter=';')

            mapa_alunos = {}
            mapa_alunos_para_ultimo_periodo_letivo_inscricao = {}
            line_number = 0

            for line in reader:
                line_number += 1
                try:
                    matr_aluno = line["MATR_ALUNO"]
                    nome_aluno = line["NOME_PESSOA"]
                    cpf_aluno = line["CPF"]
                    cod_curso = line["COD_CURSO"]

                    ano = int(line["ANO"])
                    periodo_str = line["PERIODO"]
                    periodo = int(periodo_str[0])
                except ValueError:
                    print("Valor não pode ser convertido: %s. Linha: %d." % (line["ANO"], line_number))
                    print(matr_aluno)
                    sys.exit(1)

                if matr_aluno not in mapa_alunos_para_ultimo_periodo_letivo_inscricao:
                    mapa_alunos_para_ultimo_periodo_letivo_inscricao[matr_aluno] = PeriodoLetivo(ano, periodo)
                else:
                    ultimo_matriculado = mapa_alunos_para_ultimo_periodo_letivo_inscricao[matr_aluno]
                    periodo_corrente = PeriodoLetivo(ano, periodo)
                    if (ultimo_matriculado < periodo_corrente) and (periodo_corrente <= PeriodoLetivo(Parametros.ANO_BASE, Parametros.PERIODO_BASE)):
                        mapa_alunos_para_ultimo_periodo_letivo_inscricao[matr_aluno] = periodo_corrente

                if matr_aluno not in mapa_alunos:
                    aluno = Aluno(nome_aluno, cpf_aluno, matr_aluno, cod_curso)
                    mapa_alunos[matr_aluno] = aluno

        for (matr_aluno, aluno) in mapa_alunos.items():
            aluno = mapa_alunos[matr_aluno]
            aluno.ultimo_periodo_letivo_inscricao = mapa_alunos_para_ultimo_periodo_letivo_inscricao[matr_aluno]

        '''
        NB: apenas alunos que fizeram matricula (mesmo que seja de 
        trancamento total) no período letivo tomado como base são 
        considerados. Os demais são ignorados, por estarem 
        em situação de ABANDONO ou CONCLUÍDO.
        '''
        mapa_alunos_temp = {}
        for (matr_aluno, aluno) in mapa_alunos.items():
            if aluno.tem_matricula_ativa():
                mapa_alunos_temp[matr_aluno] = aluno
        mapa_alunos = mapa_alunos_temp

        mapas_periodos = AnalisadorDesempenhoAcademico.construir_mapas_periodos(planilha, mapa_alunos)
        mapa_periodos_cursados_por_aluno = mapas_periodos[0]
        mapa_trancamentos_totais_por_aluno = mapas_periodos[1]

        mapas_reprovacoes = AnalisadorDesempenhoAcademico.construir_mapas_reprovacoes(planilha, mapa_alunos)
        mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina = mapas_reprovacoes[0]

        print(len(mapa_alunos))
        print(len(mapa_periodos_cursados_por_aluno))
        print(len(mapa_trancamentos_totais_por_aluno))
        print(len(mapas_reprovacoes))

        for (matr_aluno, aluno) in mapa_alunos.items():
            if matr_aluno in mapa_trancamentos_totais_por_aluno:
                aluno.trancamentos_totais = mapa_trancamentos_totais_por_aluno[matr_aluno]

            if matr_aluno in mapa_periodos_cursados_por_aluno:
                aluno.qtd_periodos_cursados = mapa_periodos_cursados_por_aluno[matr_aluno]

            if matr_aluno in mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina:
                aluno.qtd_maxima_reprovacoes = int(
                    mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina[matr_aluno]
                )

        return mapa_alunos

    @staticmethod
    def construir_mapas_periodos(planilha, mapa_alunos):
        with open(planilha, encoding="utf-8") as f_obj:
            reader = csv.DictReader(f_obj, delimiter=';')

            mapa_periodos_cursados_por_aluno = {}
            mapa_trancamentos_totais_por_aluno = {}
            conjunto_trancamentos_por_aluno_e_por_periodo = set()

            for line in reader:
                situacao = line["SITUACAO"]
                matr_aluno = line["MATR_ALUNO"]

                if matr_aluno not in mapa_alunos:
                    continue

                # Marca trancamentos
                if situacao == 'Trancamento Total':
                    ano_str = line["ANO"]
                    periodo = line["PERIODO"][0]
                    chave = matr_aluno + ',' + ano_str + ',' + periodo
                    conjunto_trancamentos_por_aluno_e_por_periodo.add(chave)

                # Conta períodos cursados
                if matr_aluno not in mapa_periodos_cursados_por_aluno:
                    if matr_aluno[0] == '0' or matr_aluno[0] == '1':
                        ano_ingresso = 2000 + int(matr_aluno[0:2])
                    elif matr_aluno[0] == '9':
                        ano_ingresso = 1900 + int(matr_aluno[0:2])
                    else:
                        print(matr_aluno)
                        raise ValueError('Ano de ingresso não pode ser determinado.')
                    periodo_ingresso = int(matr_aluno[2])
                    qtd_periodos_cursados = (Parametros.ANO_BASE * 2 + Parametros.PERIODO_BASE) - (
                                ano_ingresso * 2 + periodo_ingresso) + 1
                    if qtd_periodos_cursados <= 0:
                        print("Aluno fora do período base (%s/%s). Aluno: %s" %
                              (ano_ingresso, periodo_ingresso, matr_aluno))
                    else:
                        mapa_periodos_cursados_por_aluno[matr_aluno] = qtd_periodos_cursados

        # Contabiliza trancamentos totais
        for chave in conjunto_trancamentos_por_aluno_e_por_periodo:
            partes = chave.split(',')
            matr_aluno = partes[0]
            if matr_aluno not in mapa_trancamentos_totais_por_aluno:
                mapa_trancamentos_totais_por_aluno[matr_aluno] = 1
            else:
                mapa_trancamentos_totais_por_aluno[matr_aluno] += 1

        conjunto_trancamentos_por_aluno_e_por_periodo.clear()

        # Subtrai os trancamentos dos períodos cursados
        for (matr_aluno, qtd_trancamentos) in mapa_trancamentos_totais_por_aluno.items():
            mapa_periodos_cursados_por_aluno[matr_aluno] -= qtd_trancamentos

        return (mapa_periodos_cursados_por_aluno, mapa_trancamentos_totais_por_aluno)

    @staticmethod
    def construir_mapas_reprovacoes(planilha, mapa_alunos):
        print("construir mapas reprovacoes")

        with open(planilha, encoding="utf-8") as f_obj:
            reader = csv.DictReader(f_obj, delimiter=';')

            mapa_aluno_e_disciplina_para_qtd_reprovacoes = {}
            mapa_aluno_e_disciplina_para_aprovacoes = {}
            mapa_de_cod_disciplina_para_nome_disciplina = {}

            rowNum = 1
            for line in reader:
                if rowNum % 500 == 0:
                    print('# linhas processadas %d' % rowNum)
                rowNum += 1

                ano_str = line["ANO"]
                if ano_str == '':
                    continue

                cod_disciplina = line["COD_DISCIPLINA"]
                nome_disciplina = line["NOME_DISCIPLINA"]
                matr_aluno = line["MATR_ALUNO"]

                if matr_aluno not in mapa_alunos:
                    continue

                chave = cod_disciplina + ';' + matr_aluno
                mapa_de_cod_disciplina_para_nome_disciplina[cod_disciplina] = nome_disciplina

                qtd_reprovacoes = mapa_aluno_e_disciplina_para_qtd_reprovacoes.get(chave, 0)
                situacao = line["SITUACAO"]

                if situacao.startswith('Reprovado'):
                    qtd_reprovacoes += 1

                mapa_aluno_e_disciplina_para_qtd_reprovacoes[chave] = qtd_reprovacoes

                if situacao.startswith('Aprovado'):
                    mapa_aluno_e_disciplina_para_aprovacoes[chave] = 1

        # Remove casos em que o aluno foi aprovado depois de reprovar
        for chave in list(mapa_aluno_e_disciplina_para_aprovacoes.keys()):
            if chave in mapa_aluno_e_disciplina_para_qtd_reprovacoes:
                print("removendo:", chave)
                del mapa_aluno_e_disciplina_para_qtd_reprovacoes[chave]

        mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina = {}

        # Calcula máximo de reprovações por aluno
        for (chave, qtd_reprovacoes) in mapa_aluno_e_disciplina_para_qtd_reprovacoes.items():
            partes = chave.rsplit(';')
            cod_disciplina = partes[0]
            matr_aluno = partes[1]

            if matr_aluno not in mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina:
                mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina[matr_aluno] = qtd_reprovacoes
            else:
                if qtd_reprovacoes > mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina[matr_aluno]:
                    mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina[matr_aluno] = qtd_reprovacoes

        print("1)*********************************************************")
        print(mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina)
        print("2)*********************************************************")
        print(mapa_aluno_e_disciplina_para_qtd_reprovacoes)
        print("3)*********************************************************")
        print(mapa_de_cod_disciplina_para_nome_disciplina)

        return (
            mapa_de_aluno_para_qtd_max_de_reprovacoes_em_uma_disciplina,
            mapa_aluno_e_disciplina_para_qtd_reprovacoes,
            mapa_de_cod_disciplina_para_nome_disciplina
        )

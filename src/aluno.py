from pessoa import Pessoa
from parametros import Parametros, Criticidade
from periodo_letivo import PeriodoLetivo

class Aluno(Pessoa):

    def __init__(self, nome, cpf, matricula, cod_curso):
        self.matricula = matricula
        self.cod_curso = cod_curso
        self.qtd_periodos_cursados = 0;
        self.trancamentos_totais = 0
        self.qtd_maxima_reprovacoes = 0
        self.ultimo_periodo_letivo_inscricao = PeriodoLetivo(0, 0)
        super(Aluno, self).__init__(nome, cpf)

    def qtd_periodos_cursados(self):
        if self.matricula[0] == '0' or self.matricula[0] == '1':
            ano_ingresso = 2000 + int(self.matricula[0:2])
        elif self.matricula[0] == '9':
            ano_ingresso = 1900 + int(self.matricula[0:2])
        periodo_ingresso = int(self.matricula[2])
        qtd_periodos_cursados = (Parametros.ANO_BASE * 2 + Parametros.PERIODO_BASE) - (ano_ingresso * 2 + periodo_ingresso) + 1
        return qtd_periodos_cursados

    def esta_regular(self):
        qtd_maxima_periodos_integralizacao = Parametros.mapa_qtd_maxima_periodos_integralizacao[self.cod_curso]
        regular_por_qtd_periodos = self.qtd_periodos_cursados() <= qtd_maxima_periodos_integralizacao
        regular_por_qtd_reprovacoes = self.qtd_maxima_reprovacoes < Parametros.mapa_qtd_reprovacoes_jubilacao_por_curso[self.cod_curso]
        return regular_por_qtd_periodos and regular_por_qtd_reprovacoes

    def faixa_criticidade_considerando_reprovacoes(self):
        qtd = Parametros.mapa_qtd_reprovacoes_jubilacao_por_curso[self.cod_curso]
        r = self.qtd_maxima_reprovacoes
        if r < qtd - 2:
            return Criticidade.AZUL
        elif r == qtd - 2:
            return Criticidade.LARANJA
        elif r == qtd - 1:
            return Criticidade.VERMELHA
        else:
            return Criticidade.PRETA

    def faixa_criticidade_considerando_periodos(self):
        N = (1 + Parametros.mapa_qtd_maxima_periodos_integralizacao[self.cod_curso] / 2) / 2
        p = self.qtd_periodos_cursados
        if p < 2 * N:
            return Criticidade.AZUL
        elif p <= 4 * N - 4:
            return Criticidade.LARANJA
        elif p <= 4 * N - 2:
            return Criticidade.VERMELHA
        else:
            return Criticidade.PRETA
    
    def faixa_criticidade(self):
        maximo = max(self.faixa_criticidade_considerando_periodos(), self.faixa_criticidade_considerando_reprovacoes())
        # print 'Aluno %s com criticidade %d' % (self.nome, maximo)
        return maximo
    
    def tem_matricula_ativa(self):
        resultado = self.ultimo_periodo_letivo_inscricao == PeriodoLetivo(Parametros.ANO_BASE, Parametros.PERIODO_BASE)
        return resultado 

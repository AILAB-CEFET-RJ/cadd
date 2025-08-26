# coding=UTF-8

class Parametros(object):

    ANO_BASE = 0
    
    PERIODO_BASE = 0
    
    mapa_qtd_maxima_periodos_integralizacao = {"AMB": 10, "BCC": 14, "GADM": 14,
                                           "GAMB": 18, "GAUT": 18, "GCIV": 18, 
                                           "GEL": 18, "GELT": 18, "GLEA": 14, 
                                           "GMEC": 18, "GPROD": 18, "GTEL": 18, 
                                           "WEB": 10}

    mapa_qtd_reprovacoes_jubilacao_por_curso = {"AMB": 3, "BCC": 4, "GADM": 4,
                                                "GAMB": 4, "GAUT": 4, "GCIV": 4, 
                                                "GEL": 4, "GELT": 4, "GLEA": 4, 
                                                "GMEC": 4, "GPROD": 4, "GTEL": 4, 
                                                "WEB": 3}

class Criticidade:
    AZUL = 1
    LARANJA = 2
    VERMELHA = 3
    PRETA = 4
    
    @staticmethod
    def descritor(id_criticidade):
        if id_criticidade == 1:
            return "AZUL"
        elif id_criticidade == 2:
            return "LARANJA"
        elif id_criticidade == 3:
            return "VERMELHA"
        elif id_criticidade == 4:
            return "PRETA"
        else:
            raise ValueError('Valor inválido para id de criticidade')
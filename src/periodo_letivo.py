import sys

class PeriodoLetivo:

    def __init__(self, ano, periodo):
        self.ano = ano
        self.periodo = periodo

    @classmethod
    def fromstring(cls, periodo_letivo_str):
        componentes = periodo_letivo_str.split('.')
        if len(componentes) == 2:
            ano = int(componentes[0])
            periodo = int(componentes[1])
        else:
            sys.exit()
        return cls(ano, periodo)

    def __lt__(self, other):
        return (self.ano < other.ano) or ((self.ano == other.ano) and (self.periodo < other.periodo))

    def ___le__(self, other):
        return self.__lt__(other) or self.__eq__(other)

    def __eq__(self, other):
        return self.ano == other.ano and self.periodo == other.periodo

    def __ne__(self, other):
        return self.ano != other.ano or self.periodo != other.periodo

    def __gt__(self, other):
        return (self.ano > other.ano) or ((self.ano == other.ano) and (self.periodo > other.periodo))

    def __ge__(self, other):
        return self.__gt__(other) or self.__eq__(other)

    def __str__(self):
        return str(self.ano) + '.' + str(self.periodo)

    def __repr__(self):
        return self.__str__()


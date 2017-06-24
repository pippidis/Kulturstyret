'''
Denne filen inneholder klassen for søknader
'''


class application(object):
    '''
    Dette er klassen for søknader til kulturstyret. Hver søknad har et sett med
        egenskaper og opperasjoner.
    '''
    def __init__(self, arg):
        super(application, self).__init__()
        self.arg = arg


    def dummy(self, arg):
        # Lager en dummy søknadene

        self.orgName = 'Improperatørene'
        self.amount_asked = 10000
        self.amount_given = 3000

    def check(self):
        pass


    def __str__(self):
    # Gjør så man kan printe den
        sb = []
        for key in self.__dict__:
            sb.append("{key}:' {value}'".format(key=key, value=self.__dict__[key]))

        return '\n '.join(sb)

    def __repr__(self):
        return self.__str__()

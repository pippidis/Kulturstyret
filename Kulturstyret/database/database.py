'''
Denne filen inneholder alle klassen som har ansvaret for kontakt med databasen.
'''


class database(object):
    """docstring for database."""
    def __init__(self, arg):
        super(database, self).__init__()
        self.arg = arg



    def __str__(self):
    # Gjør så man kan printe den
        sb = []
        for key in self.__dict__:
            sb.append("{key}:' {value}'".format(key=key, value=self.__dict__[key]))

        return '\n '.join(sb)

    def __repr__(self):
        return self.__str__()

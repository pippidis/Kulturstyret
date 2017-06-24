'''
Denne filen inneholder klassen til organiasjonen
'''


class organisation(object):
    '''
    Denne klassen inneholder den statiske informasjonen om hver organiasjon.
        det er i all hovedsak navnet, organiasjonsnummer, statiske epostadresser,
        og  pekere til alle søknadene som kommer fra organiasjonen. Og Kommentarer.
    '''
    def __init__(self, arg):
        super(organisation, self).__init__()
        self.arg = arg

    def check(self, arg):
        pass

    def __str__(self):
    # Gjør så man kan printe den
        sb = []
        for key in self.__dict__:
            sb.append("{key}:' {value}'".format(key=key, value=self.__dict__[key]))

        return '\n '.join(sb)

    def __repr__(self):
        return self.__str__()

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

        self.org.name = 'Improperatørene'
        self.org.number = '1233243435123'
        self.org.webpage = 'pippidis.no'
        self.org.postal_adress = '2314 Langtvekkistan'

        self.bank.name = 'DNB'
        self.bank.number ='123412351232311'

        self.member.total = 100
        self.member.pay_students = 30
        self.member.pay_other = 10
        self.member.other

        self.memberFee.amount = 100
        self.memberFee.total = 100200

        self.amount.asked = 10000
        self.amount.given = 3000

        self.town = 'Trondheim'
        self.granted_before = 'True'

        self.date.generated = '20/01/1990'
        self.date.last_edited = '21/01/1990'

        self.contact.name = 'Johannes'
        self.contact.phone = '47367741'
        self.contact.email = 'pippidis+testing@gmail.com'

        self.leader.name = 'Maja'
        self.leader.phone = '22274956'
        self.leader.email = 'pippidis+testing@gmail.com'




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

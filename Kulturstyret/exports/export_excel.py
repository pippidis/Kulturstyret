'''
Denne klassen tar seg av å eksportere til excel fil for både SIT admin,
    og for eventuell annen oversikt

'''
import xlwt

class export_excel(object):
    '''Håndterer eksporteringer til excel dokumenter'''
    def __init__(self, arg):
        super(export_excel, self).__init__()
        self.arg = arg

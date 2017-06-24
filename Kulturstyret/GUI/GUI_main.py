'''
Dette er hovedfila for GUI delen av programmet. Den samler alle de andre GUI
    delene sammen og lar main kjøre hovedloopen.

'''


# Bygger på denne guiden slik den er nå: https://kivy.org/docs/guide/basic.html
# Ved problemer se her: https://stackoverflow.com/questions/40769386/kivy-windows-unable-to-find-any-valuable-window-provider-at-all
# per nå ser det ikke ut til å ville fungere i det hele tatt

import kivy
kivy.require('1.10.0') # replace with your current kivy version !

from kivy.app import App
from kivy.uix.label import Label


class MyApp(App):

    def build(self):
        return Label(text='Hello world')


if __name__ == '__main__':
    MyApp().run()

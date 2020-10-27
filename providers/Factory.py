# encoding: utf-8

from providers.OpenPyXLProvider import OpenPyXLProvider

class Factory:
    
    def __init__(self):
        self.providers = {
            "XLSXProvider": OpenPyXLProvider
        }
        

    def openProvider(self, providerName):
        return self.providers[providerName]()
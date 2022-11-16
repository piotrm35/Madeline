# this file version: 0.2

from PyQt5 import QtWidgets


#====================================================================================================================

class My_QLabel(QtWidgets.QLabel):
    
    def __init__(self, parent, text):
        super(My_QLabel, self).__init__()
        self.parent = parent
        self.text = text

    def mouseReleaseEvent(self, e):  
        self.parent.set_selected_text_QLabel(self.text)


#====================================================================================================================

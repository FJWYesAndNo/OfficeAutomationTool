# -*- coding:GBK -*-

##  GUIӦ�ó������������

import sys

from PyQt5.QtWidgets import  QApplication

from MyMainWindow import QmyMainWindow
    
app = QApplication(sys.argv)    #����GUIӦ�ó���
##icon = QIcon(":/icons/images/app.ico")
##app.setWindowIcon(icon)

mainform=QmyMainWindow()        #����������
mainform.show()                 #��ʾ������

sys.exit(app.exec_()) 

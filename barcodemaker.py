import os
import qrcode
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtGui
import pandas as pd

class MyGUI(QMainWindow):

    def __init__(self):
        super(MyGUI, self).__init__()
        uic.loadUi("barcodegui.ui", self)
        self.show()

        self.currentdoc = ""
        self.actionLoad.triggered.connect(self.loadXL)
        self.actionLoad_TXT.triggered.connect(self.loadTXT)
        self.pushButton.clicked.connect(self.generatedoc)

    def loadXL(self):
        options= QFileDialog.Options()
        filename, _ =QFileDialog.getOpenFileName(self ,"Open XL file", "", "Files (*.xlsx)")
        print("XL file is "+filename)
        df 


        pass

    def loadTXT(self):
        options= QFileDialog.Options()
        filename, _ =QFileDialog.getOpenFileName(self ,"Open TXT file", "", "Files (*.txt)")
        with open(filename,"r") as f:
            inputtext = f.read()
        self.textEdit.setPlainText(inputtext)
        
    def generatedoc(self):
        text2add = self.textEdit.toPlainText().upper().split("\n")
        document= Document()
        section = document.sections[0]
        footer= section.header
        footertext = footer.paragraphs[0]
        footer.text="Barcode Maker (c)2023 K. Winstanley"
        # Loop through the input box , create QRcode, Add to doc file
        for vin in text2add:
            if vin !='':
                print("Creating QRcode for "+ vin)
                qr = qrcode.QRCode(version=1,
                            error_correction=qrcode.constants.ERROR_CORRECT_L,
                            box_size=20,
                            border=2)
                qr.add_data(vin)
                qr.make(fit=True)
                img= qr.make_image(fill_color="Black", back_color="White")
                img.save("currentqr.png")
                document.add_heading(vin,0)
                document.add_paragraph()
                document.add_picture("currentqr.png", width=Inches(1), height=Inches(1))
                last_paragraph = document.paragraphs[-1]
                last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                document.add_paragraph()
                document.add_picture("currentqr.png", width=Inches(5), height=Inches(5))
                last_paragraph = document.paragraphs[-1]
                last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                document.add_page_break()
            
        pixmap = QtGui.QPixmap("currentqr.png")
        pixmap = pixmap.scaled(300, 300)
        
        options= QFileDialog.Options()
        filename, _ =QFileDialog.getSaveFileName(self ,"Save doc file", "", "Files (*.doc)")
        
        document.save(filename)
        self.label.setScaledContents(True)
        self.label.setPixmap(pixmap)
        
        print("Process Completed")

# for windows only
#        os.startfile(filename, "print")
    
       



def main():
    app=QApplication([])
    window = MyGUI()
    app.exec_()





if __name__ == "__main__":
    main()

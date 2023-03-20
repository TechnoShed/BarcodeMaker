import os
import qrcode
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtGui

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
        
        pass

    def loadTXT(self):
        options= QFileDialog.Options()
        filename, _ =QFileDialog.getOpenFileName(self ,"Open TXT file", "", "Files (*.txt)")
        with open(filename,"r") as f:
            inputtext = f.read()
        print(inputtext)
        self.textEdit.setPlainText(inputtext)
        
        pass

    def generatedoc(self):
        text2add = self.textEdit.toPlainText().upper().split("\n")
        document= Document()

        # Loop through the input box , create QRcode, Add to doc file and display

        for vin in text2add:
            if vin !='':
                print("vin is "+ vin)
                qr = qrcode.QRCode(version=1,
                            error_correction=qrcode.constants.ERROR_CORRECT_L,
                            box_size=20,
                            border=2)
                qr.add_data(vin)
                qr.make(fit=True)
                img= qr.make_image(fill_color="Black", back_color="White")
                img.save("currentqr.png")
                pixmap = QtGui.QPixmap("currentqr.png")
                pixmap = pixmap.scaled(300, 300)
                document.add_heading(vin,0)
                p = document.add_paragraph()
                document.add_picture("currentqr.png", width=Inches(3), height=Inches(3))
                last_paragraph = document.paragraphs[-1]
                last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                document.add_page_break()
            else:
                print("Blank line")

        document.save("barcodes.doc")
        self.label.setScaledContents(True)
        self.label.setPixmap(pixmap)
        print("Process Completed")

# for windows only
#        os.startfile("barcodes.doc", "print")
    
       



def main():
    app=QApplication([])
    window = MyGUI()
    app.exec_()





if __name__ == "__main__":
    main()

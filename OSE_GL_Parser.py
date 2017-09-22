import os
import sys

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter  # process_pdf
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO
from PySide import QtGui, QtCore

__author__ = "William Yang"

def pdf_to_text(pdfname):
    # PDFMiner boilerplate
    rsrcmgr = PDFResourceManager()
    sio = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, sio, codec=codec, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Extract text
    fp = file(pdfname, 'rb')
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)

    fp.close()

    # Get text from StringIO
    text = sio.getvalue()

    # Cleanup
    device.close()
    sio.close()

    return text

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False



class window(QtGui.QMainWindow):

    def __init__(self):
        super(window, self).__init__()

        self.setGeometry(50, 50, 500, 300)
        self.setWindowTitle("OSE GL Parser")
        self.home()

    def home(self):
        btn = QtGui.QPushButton("Select Directory", self)
        self.progress = QtGui.QProgressBar(self)
        self.progress.setGeometry(120, 30, 250, 20)


        btn.clicked.connect(self.ose_gl_parser)
        btn.resize(100,60)
        btn.move(10,10)
        self.show()

    def ose_gl_parser(self):
        dirname = QtGui.QFileDialog.getExistingDirectory()


        if dirname:
            # This is just comment
            # This is PO Parser for OSE Invoice
            # ===================================================

            # These are the Keys for Finished goods and SMC non valuable
            #    Updates this if you want to add more keys
            # ==============================================================
            Finished_Goods = ['AOC', 'CPU', 'CSE', 'DAC', 'GPU', 'HDD', 'HDK', 'HDS',
                              'MBD', 'MBL', 'MEM', 'PWS', 'SBD', 'SFA', 'SFT', 'SRK',
                              'SSD', 'SSE', 'SSG', 'SVC', 'SYS']

            SMC_non_valuable = ['CSV', 'ZMK']

            # Pdf file directory
            pdfdir = dirname # os.path.expanduser(os.getcwd() + "/pdf_files")

            po_file_output = open(pdfdir + "/PO_file.csv", 'w')

            # Write CSV file Header
            po_file_output.write(
                'FileName, Invoice Number, Customer Part Number, Order Number, Amount, Inventory PO, GL Category' + '\n')

            numfiles = len(os.listdir(pdfdir))
            count = 0
            progress_value = 0

            for i in os.listdir(pdfdir):
                progress_value = float(count)/numfiles * 100
                self.progress.setValue(progress_value)
                self.progress.show()
                count += 1

                if i.endswith(".pdf"):
                    po_file = pdf_to_text(pdfdir + '/' + i)

                    po_split = po_file.rstrip().split('\n')

                    temp = 0

                    for index, line in enumerate(po_split):
                        # Tool first finds invoice number.  Store invocice index into temp variable.
                        if po_split[index] == "INVOICENO" and po_split[index + 1] == "SuperMicroComputer,Inc.":
                            temp = index
                        # once you find Customer P/N then We can categorize if the GL
                        if po_split[index].find("CUSTP/N") >= 0:
                            customer_PN = po_split[index].split(":")
                            PN_GL_Category = customer_PN[1].split("-")
                            if PN_GL_Category[0] in Finished_Goods:
                                GL = "Finished Goods"
                            elif PN_GL_Category[0] in SMC_non_valuable:
                                GL = "SMC Non-Valuable Item"
                            else:
                                GL = "RAW Materials"

                            if po_split[index + 1].find("OrderNo") == -1:
                                order_number_list = ['OrderNo', '9999']
                            else:

                                order_number_list = po_split[index + 1].split(":")

                                if is_number(order_number_list[1]):
                                    order_number_list = po_split[index + 1].split(":")
                                else:
                                    order_number_list = ['OrderNo', '9999']  # value not number

                            if int(order_number_list[1]) == 9999:
                                # print ( po_split[temp+3], ',', customer_PN[1], ",", order_number_list[1], ', Inventory PO')
                                po_file_output.write(
                                    i + ',' + po_split[temp + 3] + ',' + customer_PN[1] + "," + order_number_list[
                                        1] + ',' + po_split[index - 2].replace(",",
                                                                               "") + ',' + 'No PO' + ',' + GL + '\n')
                            elif int(order_number_list[1]) < 3000000000:
                                # print ( po_split[temp+3], ',', customer_PN[1], ",", order_number_list[1], ', Inventory PO')
                                po_file_output.write(
                                    i + ',' + po_split[temp + 3] + ',' + customer_PN[1] + "," + order_number_list[
                                        1] + ',' + po_split[index - 2].replace(",",
                                                                               "") + ',' + 'Inventory PO' + ',' + GL + '\n')
                            else:
                                # print (po_split[temp+3],",", customer_PN[1], ',', order_number_list[1], ', Non-Inventory PO')
                                po_file_output.write(
                                    i + ',' + po_split[temp + 3] + "," + customer_PN[1] + ',' + order_number_list[
                                        1] + ',' + po_split[index - 2].replace(",",
                                                                               "") + ',' + ', Non-Inventory PO' + ',' + GL + '\n')

                                # f.close()
                else:
                    # print ("hi")
                    continue


            po_file_output.close()
            self.progress.setValue(100)
            QtGui.QMessageBox.about(self, "Message Box", "Parser is DONE!")

def run():
    app = QtGui.QApplication(sys.argv)
    GUI = window()
    sys.exit(app.exec_())

run()




















import datetime
import sqlite3
import os
import sys
from PyQt6.QtWidgets import *
from PyQt6.QtCore import Qt
from docxtpl import DocxTemplate

class mailbackGenWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Test Mailback Letter Generator")
        self.setFixedSize(722, 479)

        main_layout = QVBoxLayout()

        self.db = sqlite3.connect("test_mailback.db")
        self.cur = self.db.cursor()

        client_label = QLabel("Select Client: ")
        self.client_select = QComboBox()
        self.populateClientSelect()

        def setAndGet():
            self.getDefaultAddress()
            self.setDefaultAddress()

        self.client_select.currentIndexChanged.connect(setAndGet)

        reason_label = QLabel("Select All Reasons for Return ")
        self.reason_select = QFrame()
        self.reason_layout = QGridLayout()
        self.reasonCheckBoxList = []
        self.address1 = QLineEdit()
        self.address1.setFixedWidth(200)
        self.address2 = QLineEdit()
        self.address2.setFixedWidth(200)
        self.address3 = QLineEdit()
        self.address3.setFixedWidth(200)

        self.clear_address_button = QPushButton("Clear Address")
        self.clear_address_button.clicked.connect(self.clearAddress)
        self.default_address_button = QPushButton("Default")
        self.default_address_button.clicked.connect(self.setDefaultAddress)
        self.populateReasonLayout()
        self.reason_select.setLayout(self.reason_layout)

        self.reason_error = QLabel("Please select at least one reason.")
        self.reason_error.setStyleSheet("color: red")
        self.reason_error.hide()

        self.envelope_button = QPushButton("Generate Envelope")
        self.envelope_button.clicked.connect(self.printEnvelope)

        self.large_envelope_button = QPushButton("Large Envelope Sheet")
        self.large_envelope_button.clicked.connect(self.printLargeEnvelope)

        self.submit_button = QPushButton("Generate Letter")
        self.submit_button.clicked.connect(self.generateLetter)

        widgets = [client_label, self.client_select, reason_label, self.reason_select, self.reason_error,
                   self.submit_button, self.envelope_button, self.large_envelope_button]

        for w in widgets:
            main_layout.addWidget(w)
        widget = QWidget()
        widget.setLayout(main_layout)

        # Set the central widget of the Window. Widget will expand
        # to take up all the space in the window by default.
        self.setCentralWidget(widget)
        self.template = DocxTemplate("test_mailback_template.docx")

        self.envelope = DocxTemplate("mailout.docx")

        self.big_envelope = DocxTemplate("large envelope template.docx")

        self.current_date = datetime.date.today().strftime('%m/%d/%Y')

        self.currentClient = ""
        self.currentAddr1 = ""
        self.currentAddr2 = ""
        self.currentPhoneNumber = ""
        self.getDefaultAddress()
        self.setDefaultAddress()

    def populateClientSelect(self):
        tups = self.cur.execute("""SELECT query_name FROM client
                                ORDER BY query_name ASC;""")

        clients = [name for t in tups for name in t]
        self.client_select.addItems(clients)

    def getDefaultAddress(self):
        client_name = self.client_select.currentText()
        client_row = self.cur.execute("""SELECT full_name, address, phone_number 
                                                    FROM client 
                                                    WHERE query_name = ?""", (client_name,))
        self.currentClient, full_addr, self.currentPhoneNumber = [c for t in client_row for c in t]
        self.currentAddr1, self.currentAddr2 = full_addr.split('*')

    def setDefaultAddress(self):
        self.address1.setText(self.currentClient)
        self.address2.setText(self.currentAddr1)
        self.address3.setText(self.currentAddr2)

    def clearAddress(self):
        self.address1.clear()
        self.address2.clear()
        self.address3.clear()

    def populateReasonLayout(self):
        reasonTypes = self.cur.execute("""SELECT DISTINCT type FROM mailback_reason;""")
        reasonTypes = [t for rt in reasonTypes for t in rt]
        print(reasonTypes)

        column = 0
        row = 0
        for t in reasonTypes:
            if column == 2:
                column = 0
                row += 1

            frame = QFrame()
            layout = QVBoxLayout()
            layout.addWidget(QLabel(t + ':'))
            reasons = self.cur.execute("""SELECT reason FROM mailback_reason
                                            WHERE type = ?;""", (t,))
            reasons = [r for rt in reasons for r in rt]
            for r in reasons:
                box = QCheckBox(r)
                self.reasonCheckBoxList.append(box)
                layout.addWidget(box)
            frame.setLayout(layout)
            self.reason_layout.addWidget(frame, column, row, Qt.AlignmentFlag.AlignTop)

            column += 1

        if column == 2:
            column = 0
            row += 1

        frame = QFrame()
        layout = QGridLayout()
        layout.addWidget(QLabel('Name:'), 0, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.address1, 0, 1, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(QLabel('Address:'), 1, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.address2, 1, 1, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(QLabel('City/State/Zip:'), 2, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.address3, 2, 1, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.clear_address_button, 3, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.default_address_button, 3, 1, Qt.AlignmentFlag.AlignLeft)
        frame.setLayout(layout)
        self.reason_layout.addWidget(frame, column, row, Qt.AlignmentFlag.AlignLeft)


    def generateLetter(self):
        #FOR SETTING FIXED WIDTH/HEIGHT
        #print(self.width())
        #print(self.height())

        # avoids Microsoft Word opening dialog box saying that letter.docx caused error
        if os.path.exists("letter.docx"):
            os.remove("letter.docx")

        reasons = []

        for box in self.reasonCheckBoxList:
            if box.isChecked():
                reasons.append(box.text())
                box.setChecked(False)

        reason = ""
        rlength = len(reasons)
        if rlength == 1:
            reason = reasons[0]
        elif rlength == 2:
            reason = reasons[0] + ' and ' + reasons[1]
        elif rlength > 2:
            for i in range(0, rlength):
                if i != rlength - 1:
                    reason += reasons[i] + ', '
                else:
                    reason += 'and ' + reasons[i]
        else:  # reasons is empty
            self.reason_error.show()
            return 1

        self.reason_error.hide()
        self.submit_button.setEnabled(False)

        fill_in = {"date": self.current_date,
                   "client": self.currentClient,
                   "reason": reason,
                   "address_1": self.currentAddr1,
                   "address_2": self.currentAddr2,
                   "phone_number": self.currentPhoneNumber
                   }

        self.template.render(fill_in)
        self.template.save('letter.docx')
        os.startfile("letter.docx", "print")

        self.submit_button.setEnabled(True)

    def printEnvelope(self):
        self.envelope_button.setEnabled(False)

        fill_in = {"client": self.address1.text(),
                   "addr_1": self.address2.text(),
                   "addr_2": self.address3.text()}

        self.envelope.render(fill_in)
        self.envelope.save('envelope.docx')
        os.startfile("envelope.docx", "print")

        self.envelope_button.setEnabled(True)

    def printLargeEnvelope(self):
        self.large_envelope_button.setEnabled(False)


        fill_in = {"client": self.address1.text(),
                   "addr_1": self.address2.text(),
                   "addr_2": self.address3.text()}

        self.big_envelope.render(fill_in)
        self.big_envelope.save('big_envelope.docx')
        os.startfile("big_envelope.docx", "print")

        self.large_envelope_button.setEnabled(True)

def main():
    app = QApplication(sys.argv)
    window = mailbackGenWindow()
    window.show()

    app.exec()


if __name__ == "__main__":
    main()

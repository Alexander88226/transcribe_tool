from PyQt5 import QtWidgets, uic, QtCore, QtGui

from PyQt5.QtWidgets import QFileDialog, QMessageBox

import sys
import openpyxl
import os
import re

def display_error_message(content):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText("Error")
    msg.setInformativeText(content)
    msg.setWindowTitle("Error")
    msg.exec_()

def get_file_type(extension):
    if extension == ".txt":
        extension = "Text"
    if extension == ".xlsx":
        extension = "Excel"
    if extension == ".xliff":
        extension = "Xliff"
    if extension == ".tmx":
        extension = "Tmx"
    return extension

CONVERT_TYPE = ["Text to Excel", "Excel to Text", "Text to Tmx", "Tmx to Text", "Text to Xliff", "Xliff to Text", "Excel to Xliff", "Xliff to Excell", \
    "Excel to Tmx", "Tmx to Excel", "Xliff to Tmx", "Tmx to Xliff"]

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        uic.loadUi("Main.ui", self)
        self.selected_task_id = -1
        self.tbl_task.setHorizontalHeaderLabels(['Source File', 'Destination File', 'Convert Type'])
        header = self.tbl_task.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

    @QtCore.pyqtSlot()
    def create_task(self):
        dialog = Dialog()
        if dialog.exec_():# accept, save item
            rows = self.tbl_task.rowCount()
            self.tbl_task.insertRow(rows)
            source_item = QtWidgets.QTableWidgetItem(dialog.txt_source.text())
            source_item.setToolTip(dialog.txt_source.text())
            self.tbl_task.setItem(rows, 0, source_item)
            dest_item = QtWidgets.QTableWidgetItem(dialog.txt_destination.text())
            dest_item.setToolTip(dialog.txt_destination.text())
            self.tbl_task.setItem(rows, 1, dest_item)
            convert_type = ""
            filename, file_extension = os.path.splitext(dialog.txt_source.text())
            
            convert_type += get_file_type(file_extension) + " to "
            filename, file_extension = os.path.splitext(dialog.txt_destination.text())
            convert_type += get_file_type(file_extension)            
            self.tbl_task.setItem(rows, 2, QtWidgets.QTableWidgetItem(convert_type))
        else:# reject
            pass
        
    @QtCore.pyqtSlot()
    def remove_task(self):
        if self.selected_task_id != -1:
            self.tbl_task.removeRow(self.selected_task_id)
            self.selected_task_id = -1
    @QtCore.pyqtSlot()
    def run_task(self):
        if self.selected_task_id != -1:
            source_file = self.tbl_task.item(self.selected_task_id, 0).text()
            dest_file = self.tbl_task.item(self.selected_task_id, 1).text()
            c_type = self.tbl_task.item(self.selected_task_id, 2).text()
            if c_type == "Text to Excel":
                convert_Text2Excel(source_file, dest_file)
            if c_type == "Excel to Text":
                convert_Excel2Text(source_file, dest_file)
            if c_type == "Text to Tmx":
                convert_Text2Tmx(source_file, dest_file)
            if c_type == "Tmx to Text":
                convert_Tmx2Text(source_file, dest_file)
            if c_type == "Text to Xliff":
                convert_Text2Xliff(source_file, dest_file)
            if c_type == "Xliff to Text":
                convert_Xliff2Text(source_file, dest_file)
            if c_type == "Excel to Tmx":
                convert_Excel2Tmx(source_file, dest_file)
            if c_type == "Tmx to Excel":
                convert_Tmx2Excel(source_file, dest_file)
            if c_type == "Tmx to Xliff":
                convert_Tmx2Xliff(source_file, dest_file)
            if c_type == "Xliff to Tmx":
                convert_Xliff2Tmx(source_file, dest_file)
            if c_type == "Excel to Xliff":
                convert_Excel2Xliff(source_file, dest_file)
            if c_type == "Xliff to Excel":
                convert_Xliff2Excel(source_file, dest_file)


    @QtCore.pyqtSlot()
    def select_item(self):
        self.selected_task_id = self.tbl_task.currentRow()

class Dialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        uic.loadUi("Task.ui", self)
    @QtCore.pyqtSlot()
    def browse_source(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Source File", "","Available File(*.txt *.xlsx *.tmx *.xliff);;Bilingual Text Files (*.txt);;Excel File(*.xlsx);;Tmx File(*.tmx);;Xliff file(*.xliff)", options=options)
        if fileName:
            self.txt_source.setText(fileName)

    @QtCore.pyqtSlot()
    def browse_destination(self):
        extension_list = ['.txt', '.xlsx', '.tmx', '.xliff']
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, extension = QFileDialog.getSaveFileName(self,"Select Destination File","","Bilingual Text Files (*.txt);;Excel File(*.xlsx);;Tmx File(*.tmx);;Xliff file(*.xliff)", options=options)        
        if fileName:
            for ext in extension_list:
                if ext in extension and not fileName.endswith(ext):
                    fileName += ext
                    break
            self.txt_destination.setText(fileName)
    @QtCore.pyqtSlot()
    def save(self):
        if self.txt_source.text() == "":
            display_error_message("Please put the source file")
        elif self.txt_destination.text() == "":
            display_error_message("Please put the destination file")
        else:
            source_path = self.txt_source.text()
            dest_path = self.txt_destination.text()
            s_filename, s_file_extension = os.path.splitext(self.txt_source.text())
            d_filename, d_file_extension = os.path.splitext(self.txt_destination.text())
            if source_path == dest_path:
                QMessageBox.warning(self, "Warning", "File paths are the same")
                return
            if s_file_extension == d_file_extension:
                QMessageBox.warning(self, "Warning", "Extensions are the same")
                return
            self.accept()

if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication([])
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


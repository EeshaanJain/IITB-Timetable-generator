# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
import xlsxwriter
import slot
import mappings
import ctypes
import time


myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

def set_defaults():
    # Getting maps from mappings
    time_map = mappings.get_time_maps()
    header = mappings.get_headers()

    # Setting defaults
    worksheet.set_default_row(16)
    worksheet.set_row(6, 22)
    worksheet.set_row(11, 22)

    # Setting column properties
    worksheet.set_column(0, 5, 20)

    # Setting first row
    worksheet.write_row('A1', header, cell_format_center_bold)


    for i, time_map_ in enumerate(time_map):
        worksheet.write(i+1, 0, time_map_, cell_format_center_bold)
        if time_map_ == '12:30 - 14:00':
            worksheet.merge_range(i+1, 1, i+1, 5, 'LUNCH', cell_format_break)

        if time_map_ == '17:00 - 17:30':
            worksheet.merge_range(i+1, 1, i+1, 5, 'SNACKS', cell_format_break)




def empty_fill():
    formatter = workbook.add_format()
    formatter.set_align('center')
    formatter.set_align('vcenter')
    formatter.set_border()
    for key in slot_maps.keys():
        cell = slot_maps[key]
        if(len(cell)>3):
            worksheet.merge_range(cell, '', formatter)
        else:
            worksheet.write(cell, '', formatter)


class Ui_Timetable(object):
    def setupUi(self, Timetable):
        Timetable.setObjectName("Timetable")
        Timetable.resize(720, 503)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 127, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 170, 170))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.BrightText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Shadow, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 220))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ToolTipText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 127, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 170, 170))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.BrightText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Shadow, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 220))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ToolTipText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 127, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 127, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 170, 170))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 127, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.BrightText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 127, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Shadow, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 220))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ToolTipText, brush)
        Timetable.setPalette(palette)
        font = QtGui.QFont()
        font.setPointSize(12)
        Timetable.setFont(font)
        self.Heading = QtWidgets.QLabel(Timetable)
        self.Heading.setGeometry(QtCore.QRect(60, 10, 331, 51))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Semibold")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.Heading.setFont(font)
        self.Heading.setObjectName("Heading")
        self.CourseCodeL = QtWidgets.QLabel(Timetable)
        self.CourseCodeL.setGeometry(QtCore.QRect(4, 80, 211, 40))
        self.CourseCodeL.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.CourseCodeL.setObjectName("CourseCodeL")
        self.SlotL = QtWidgets.QLabel(Timetable)
        self.SlotL.setGeometry(QtCore.QRect(4, 130, 211, 40))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.SlotL.setFont(font)
        self.SlotL.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.SlotL.setObjectName("SlotL")
        self.TAL = QtWidgets.QLabel(Timetable)
        self.TAL.setGeometry(QtCore.QRect(4, 180, 211, 40))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.TAL.setFont(font)
        self.TAL.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.TAL.setObjectName("TAL")
        self.TutL = QtWidgets.QLabel(Timetable)
        self.TutL.setGeometry(QtCore.QRect(4, 230, 211, 40))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.TutL.setFont(font)
        self.TutL.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.TutL.setObjectName("TutL")
        self.Course_num = QtWidgets.QTextEdit(Timetable)
        self.Course_num.setGeometry(QtCore.QRect(220, 80, 141, 40))
        self.Course_num.setObjectName("Course_num")
        self.TA = QtWidgets.QComboBox(Timetable)
        self.TA.setGeometry(QtCore.QRect(220, 180, 141, 40))
        self.TA.setObjectName("TA")
        self.TA.addItem("")
        self.TA.addItem("")
        self.Tut = QtWidgets.QComboBox(Timetable)
        self.Tut.setGeometry(QtCore.QRect(220, 230, 141, 40))
        self.Tut.setObjectName("Tut")
        self.Tut.addItem("")
        self.Tut.addItem("")
        self.Slot = QtWidgets.QComboBox(Timetable)
        self.Slot.setGeometry(QtCore.QRect(220, 130, 141, 40))
        self.Slot.setObjectName("Slot")
        self.Table = QtWidgets.QTableWidget(Timetable)
        self.Table.setGeometry(QtCore.QRect(410, 30, 290, 451))
        self.Table.setObjectName("Table")
        self.Table.setColumnCount(2)
        self.Table.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.Table.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.Table.setHorizontalHeaderItem(1, item)
        self.AddCourse = QtWidgets.QPushButton(Timetable)
        self.AddCourse.setGeometry(QtCore.QRect(150, 320, 131, 41))
        self.AddCourse.setObjectName("AddCourse")
        self.Generate = QtWidgets.QPushButton(Timetable)
        self.Generate.setGeometry(QtCore.QRect(100, 390, 231, 71))
        self.Generate.setObjectName("Generate")

        self.slots = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15',
                      'X1', 'X2', 'X3', 'XC', 'XD', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'Lx']

        self.selected_slots = []

        self.Slot.addItems(self.slots)

        self.AddCourse.clicked.connect(self.add_entry)
        self.Generate.clicked.connect(self.generate_tt)

        self.retranslateUi(Timetable)
        QtCore.QMetaObject.connectSlotsByName(Timetable)

    def retranslateUi(self, Timetable):
        _translate = QtCore.QCoreApplication.translate
        Timetable.setWindowTitle(_translate("Timetable", "IIT Bombay Timetable Maker"))
        app_icon = QtGui.QIcon()
        app_icon.addFile('gui/icons/16x16.png', QtCore.QSize(16,16))
        app_icon.addFile('gui/icons/24x24.png', QtCore.QSize(24,24))
        app_icon.addFile('gui/icons/32x32.png', QtCore.QSize(32,32))
        app_icon.addFile('gui/icons/48x48.png', QtCore.QSize(48,48))
        #app_icon.addFile('gui/icons/256x256.png', QtCore.QSize(256,256))
        app.setWindowIcon(app_icon)
        self.Heading.setText(_translate("Timetable", "IIT Bombay Timetable"))
        self.CourseCodeL.setText(_translate("Timetable", "Course Code: "))
        self.SlotL.setText(_translate("Timetable", "Slot: "))
        self.TAL.setText(_translate("Timetable", "Are you a TA for this?: "))
        self.TutL.setText(_translate("Timetable", "Is this a tutorial?: "))
        self.TA.setItemText(0, _translate("Timetable", "No"))
        self.TA.setItemText(1, _translate("Timetable", "Yes"))
        self.Tut.setItemText(0, _translate("Timetable", "No"))
        self.Tut.setItemText(1, _translate("Timetable", "Yes"))
        item = self.Table.horizontalHeaderItem(0)
        item.setText(_translate("Timetable", "Course"))
        item = self.Table.horizontalHeaderItem(1)
        item.setText(_translate("Timetable", "Slot"))
        self.AddCourse.setText(_translate("Timetable", "Add Course"))
        self.Generate.setText(_translate("Timetable", "Generate Timetable"))




    def append_table(self, course_number, slot_num):
        rowPosition = self.Table.rowCount()
        self.Table.insertRow(rowPosition)
        numcols = self.Table.columnCount()
        numrows = self.Table.rowCount()
        self.Table.setRowCount(numrows)
        self.Table.setColumnCount(numcols)
        cn = QtWidgets.QTableWidgetItem(course_number)
        cn.setTextAlignment(QtCore.Qt.AlignHCenter)
        self.Table.setItem(numrows-1, 0, cn)
        sn = QtWidgets.QTableWidgetItem(slot_num)
        sn.setTextAlignment(QtCore.Qt.AlignHCenter)
        self.Table.setItem(numrows-1, 1, sn)


    def add_entry(self):

        course_number = self.Course_num.toPlainText()
        slot_num = self.Slot.currentText()
        if self.TA.currentText()=='Yes':
            is_ta = 1
        else:
            is_ta = 0

        if self.Tut.currentText()=='Yes':
            is_tut = 1
        else:
            is_tut = 0

        slot_ = slot.Slot(slot_num, course_number, is_tut, is_ta)
        self.selected_slots.append(slot_)

        if slot_num == '8' or slot_num == '9':
            if 'L1' in self.slots:
                self.slots.remove('L1')
            if 'L3' in self.slots:
                self.slots.remove('L3')
        elif slot_num == 'L1' or slot_num == 'L3':
            if '8' in self.slots:
                self.slots.remove('8')
            if '9' in self.slots:
                self.slots.remove('9')
        elif slot_num == '10' or slot_num == '11':
            if 'L2' in self.slots:
                self.slots.remove('L2')
            if 'L4' in self.slots:
                self.slots.remove('L4')
        elif slot_num == 'L2' or slot_num == 'L4':
            if '10' in self.slots:
                self.slots.remove('10')
            if '11' in self.slots:
                self.slots.remove('11')
        elif slot_num == '5' or slot_num == '6':
            if 'L5' in self.slots:
                self.slots.remove('L5')
            if 'L6' in self.slots:
                self.slots.remove('L6')
        elif slot_num == 'L5' or slot_num == 'L6':
            if '5' in self.slots:
                self.slots.remove('5')
            if '6' in self.slots:
                self.slots.remove('6')
        elif slot_num == 'X1' or slot_num == 'X2' or slot_num == 'X3':
            if 'Lx' in self.slots:
                self.slots.remove('Lx')
        elif slot_num == 'Lx':
            if 'X1' in self.slots:
                self.slots.remove('X1')
            if 'X2' in self.slots:
                self.slots.remove('X2')
            if 'X3' in self.slots:
                self.slots.remove('X3')

        self.slots.remove(slot_num)
        self.Slot.clear()
        self.Course_num.clear()
        self.Slot.addItems(self.slots)


        self.append_table(course_number, slot_num)

    def generate_tt(self):

        for new_slot in self.selected_slots:
            msg = new_slot.course_number
            if new_slot.is_ta:
                msg += " (TA) "
            elif new_slot.is_tut:
                msg += " (Tut) "
            formatter = workbook.add_format()
            formatter.set_align('center')
            formatter.set_align('vcenter')
            formatter.set_border()
            formatter.set_bg_color(new_slot.color)
            formatter.set_font_size(13)

            if new_slot.is_sub_slot:

                cell = slot_maps[new_slot.slot_num]

                if(len(cell)>3):
                    worksheet.merge_range(cell, msg, formatter)
                else:
                    worksheet.write(cell, msg, formatter)
            else:
                sub_slots_list = mappings.get_sub_slots()[new_slot.slot_num]
                if new_slot.slot_num.startswith('L'):
                    start  = slot_maps[sub_slots_list[0]].split(':')[0]
                    end = slot_maps[sub_slots_list[-1]].split(':')[-1]
                    cell = start + ':' + end
                    worksheet.merge_range(cell, new_slot.course_number, formatter)
                else:
                    for sub_slot in sub_slots_list:
                        cell = slot_maps[sub_slot]

                        if(len(cell)>3):
                            worksheet.merge_range(cell, msg, formatter)
                        else:
                            worksheet.write(cell, msg, formatter)
        workbook.close()

        sys.exit(1)



if __name__ == "__main__":
    import sys
    import os
    curr = os.getcwd()
    workbook = xlsxwriter.Workbook(os.path.join(curr,'Timetable.xlsx'))
    worksheet = workbook.add_worksheet()
    # Add center formatting
    cell_format_center = workbook.add_format()
    cell_format_center.set_align('center')
    cell_format_center.set_align('vcenter')
    cell_format_center.set_border()

    # Add center-bold formatting
    cell_format_center_bold = workbook.add_format()
    cell_format_center_bold.set_align('center')
    cell_format_center_bold.set_align('vcenter')
    cell_format_center_bold.set_bold()
    cell_format_center_bold.set_border()
    cell_format_center_bold.set_font_size(13)

    # Add break formatting
    cell_format_break = workbook.add_format()
    cell_format_break.set_align('center')
    cell_format_break.set_align('vcenter')
    cell_format_break.set_border(2)
    cell_format_break.set_font_size(18)

    slot_maps = mappings.get_slot_maps()

    set_defaults()
    empty_fill()

    app = QtWidgets.QApplication(sys.argv)
    Timetable = QtWidgets.QDialog()
    ui = Ui_Timetable()
    ui.setupUi(Timetable)
    Timetable.show()
    sys.exit(app.exec_())

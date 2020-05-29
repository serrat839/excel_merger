import sys
from PySide2.QtCore import Slot
from PySide2.QtGui import QKeySequence, QStandardItemModel, QStandardItem
from PySide2.QtWidgets import QMainWindow, QAction, QApplication, QVBoxLayout, QHBoxLayout,\
  QPushButton, QWidget, QGroupBox, QFileDialog, QLineEdit, QListView, QLabel, QAbstractItemView, QMessageBox
import pandas as pd

class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.GROUP1 = "Group 1"
        self.GROUP2 = "Group 2"
        self.worksheets = {}
        self.selected = {}
        self.columns = {}
        self.FILE_SELECT = "File Select"
        self.setWindowTitle("Excel Column Merger")
        self.__init_ui()

        self.setCentralWidget(self.main)

        # Window dimensions
        geometry = qApp.desktop().availableGeometry(self)
        self.resize(900, 600)

    # This function sets up our applications ui
    def __init_ui(self):
        # Make main widget
        self.main = QWidget()

        # Make our main widget's layout
        vert_layout = QVBoxLayout()
        self.main.setLayout(vert_layout)

        horiz_layout = QHBoxLayout()
        vert_layout.addLayout(horiz_layout)

        # Add file control groups to our layout
        horiz_layout.addWidget(self.__create_vertical_section(self.GROUP1))
        horiz_layout.addWidget(self.__create_vertical_section(self.GROUP2))

        button = QPushButton("Merge!")
        button.clicked.connect(self.merge)
        vert_layout.addWidget(button)

    # This function takes in a group_name and then populates a QGroupBox with
    # the proper widgets to navigate and select parts of an excel file.
    def __create_vertical_section(self, group_name):
        # Setup layout
        groupbox = QGroupBox(group_name)
        groupbox.setObjectName(group_name)
        vertical_layout = QVBoxLayout()
        groupbox.setLayout(vertical_layout)

        '''
        Add widgets
        '''
        # Add file selection/loading widgets
        file_select_organizer = QHBoxLayout()
        text_input = QLineEdit() # Does this need a name yet?
        file_select_button = QPushButton("...")
        file_select_organizer.addWidget(text_input)
        file_select_organizer.addWidget(file_select_button)
        # Add the section we just made to layout
        vertical_layout.addLayout(file_select_organizer)

        # add listview for excel pages
        excel_label = QLabel("Excel Workbook Pages")
        excel_sheets = QListView()
        excel_sheets.setEditTriggers(QAbstractItemView.NoEditTriggers)
        excel_label_model = QStandardItemModel()
        excel_sheets.setModel(excel_label_model)

        vertical_layout.addWidget(excel_label)
        vertical_layout.addWidget(excel_sheets)


        # add listview for column headers
        variable_label = QLabel("Merge on column:")
        variables = QListView()
        variables.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.columns[group_name] = variables
        variables_model = QStandardItemModel()
        variables.setModel(variables_model)
        vertical_layout.addWidget(variable_label)
        vertical_layout.addWidget(variables)

        '''
        Connect Functions
        '''
        # Connect File dialog to file selection
        file_select_button.clicked.connect(lambda: self.openFileNameDialog(text_input, excel_label_model, group_name))
        # Connect listview to populate listview for column headers
        excel_sheets.clicked.connect(lambda x: self.populateColumns(x, excel_label_model, variables_model, group_name))

        return groupbox

    def populateColumns(self, index, model, insert_model, group):
        page = model.itemFromIndex(index).text()
        self.selected[group] = self.worksheets[group][page]
        insert_model.removeRows(0, insert_model.rowCount())
        for column in self.selected[group].columns:
            insert_model.appendRow(QStandardItem(column))

    def openFileNameDialog(self, line_edit, list_model, group_name):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Select an Excel file", "", "Excel Workbooks (*.xlsx)", options=options)
        if fileName:
            line_edit.setText(fileName)
            self.worksheets[group_name] = pd.read_excel(fileName, None)
            list_model.removeRows(0, list_model.rowCount())
            for sheet in self.worksheets[group_name]:
                list_model.appendRow(QStandardItem(sheet))

    def merge(self):
        vars = []
        try:
            for x in self.columns:
                index = self.columns[x].selectedIndexes()
                vars.append(index[0].data())

            # WRAp up this section then u r done and free from thsi mortal coil
            if len(vars) == 2:

                saved = self.saveFileDialog()

                if saved:
                    # ADD THING TO TELL USER THAT THEY CAN SAVE THE EXCEL WORKSHEET!!
                    result = self.selected[self.GROUP1].merge(right=self.selected[self.GROUP2], how="outer", left_on=vars[0], right_on=vars[1])
                    result.to_excel(saved)

        except:
            msg = QMessageBox()
            msg.setWindowTitle("Not enough information!")
            msg.setText("Please make sure you have a sheet and column selected for each Excel workbook")
            msg.setIcon(QMessageBox.Critical)
            x = msg.exec_()

    def saveFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, "Save file location", "",
                                                  "Excel Workbooks (*.xlsx)", options=options)
        if fileName:
            return (fileName)
        return (None)




if __name__ == "__main__":
  app = QApplication([])
  window = MainWindow()
  window.show()
  app.exec_()
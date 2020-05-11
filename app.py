import sys
from PySide2.QtCore import Slot
from PySide2.QtGui import QKeySequence
from PySide2.QtWidgets import QMainWindow, QAction, QApplication, QVBoxLayout, QHBoxLayout,\
  QPushButton, QWidget, QGroupBox, QFileDialog, QLineEdit

class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.GROUP1 = "Group 1"
        self.GROUP2 = "Group 2"
        self.FILE_SELECT = "File Select"
        self.setWindowTitle("Excel Column Merger")
        self.__init_ui()

        self.setCentralWidget(self.main)

        # Window dimensions
        geometry = qApp.desktop().availableGeometry(self)
        self.setFixedSize(geometry.width() * 0.8, geometry.height() * 0.7)

    # This function sets up our applications ui
    def __init_ui(self):
        # Make main widget
        self.main = QWidget()

        # Make our main widget's layout
        horiz_layout = QHBoxLayout()
        self.main.setLayout(horiz_layout)

        # Add file control groups to our layout
        horiz_layout.addWidget(self.__create_vertical_section(self.GROUP1))
        horiz_layout.addWidget(self.__create_vertical_section(self.GROUP2))

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


        # add listview for column headers

        '''
        Connect Functions
        '''

        file_select_button.clicked.connect(lambda: self.openFileNameDialog(text_input))
        # add function to populate listview for excel pages
        # add funciton to populate listview for column headers
        return groupbox

    def openFileNameDialog(self, line_edit, listview):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"Title", "","Excel Workbooks (*.xlsx)", options=options)
        if fileName:
          line_edit.setText(fileName)
          # populate page listview


if __name__ == "__main__":
  app = QApplication([])
  window = MainWindow()
  window.show()
  app.exec_()
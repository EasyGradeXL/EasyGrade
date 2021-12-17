#************************************************************************************************************************************************************************************
# Copyright (c) 2021 Tony L. Jones
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”),
# to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
# OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#
# The sound files were downloaded from https://www.fesliyanstudios.com
#************************************************************************************************************************************************************************

import os
import win32com.client
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox, QLabel, QDesktopWidget
from PyQt5.QtGui import *
from PyQt5.QtCore import QEventLoop, QTimer
import sys
from pathlib import Path
import shutil

def resourcePath(relativePath):
    '''
    Gets the absolute path to resource. Required for PyInstaller. Its input is the relative path.
    It gets the absolute path from the operating system and appends the relative path.

    Parameters
    ----------
    relativePath :
        The relative path of the resource.
    Returns
    -------
        A path-like object representing a file system path.
    '''
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        basePath = sys._MEIPASS
    except Exception:
        basePath = os.path.abspath('.')
    return os.path.join(basePath, relativePath)

class MainW(QWidget):
    """
        Sends the audio chunks to the Google speech-to-text API and sends the recognized text to Excel

        Attributes
        ----------
        excelFileName : str
            The full name and path of the Excel file that is a copy of the template
        jsonFileName : str
            The full name and path of the Google Cloud credentials file
        numTimesIncompleteExitCalled : int
            Number of times the IncompleteExit method was called

        Methods
        _______
            runGUI()
                Runs the graphical user interface
            isExcelOpen
                Checks if an instance of MS Excel is running
            askActivateSpeechRecognition()
                Asks the user if speech recognition must be enabled
            openJsonFileNameDialog()
                Asks the user for the name of the Google credentials file
            openXlsmFileNameDialog()
                Asks the user for the name and folder under which the xlsm file must be save
            saveGradeBookCopy()
                Saves a copy of the GradeBook.xlsm template
            setSettings()
                Writes the settings to the Settings sheet of the workbook
            incompleteExit()
                Displays a message to inform the user that Setup is incomplete and exits the application
    """

    def __init__(self):
        super(MainW, self).__init__()
        self.excelFileName = None
        self.jsonFileName = None
        self.numTimesIncompleteExitCalled = 0
        # Display image
        label = QLabel(self)
        pixmap = QPixmap(resourcePath('pineAppleBackGround.jpg'))
        sizeObject = QDesktopWidget().screenGeometry(-1)
        pixmapScaled = pixmap.scaledToWidth(int(sizeObject.width()/2))
        label.setPixmap(pixmapScaled)
        # Resize window to image size
        self.resize(pixmapScaled.width(), pixmapScaled.height())
        self.setWindowIcon(QIcon(resourcePath('pineAppleBackGround.jpg')))
        self.setWindowTitle("EasyGradeXL Setup")
        self.show()
        # Display the image for 3 s
        loop = QEventLoop()
        QTimer.singleShot(3000, loop.quit)  # Wait 3 s
        loop.exec_()
        self.runGUI()

    def runGUI(self):
        """
        Runs the main GUI
        """

        if self.isExcelOpen():
            button = QMessageBox.warning(self, "Excel is running",
                                         "Microsoft Excel is currently running on this computer.\
                                          Please close Excel before proceeding.",
                                         QMessageBox.Ok | QMessageBox.Cancel)
            while button == QMessageBox.Ok and self.isExcelOpen():
                button = QMessageBox.warning(self, "Excel is running",
                                            "Microsoft Excel is currently running on this computer.\
                                             Please close Excel before proceeding.",
                                            QMessageBox.Ok | QMessageBox.Cancel)
            if button == QMessageBox.Cancel:
                self.incompleteExit()

        self.saveGradeBookCopy()

        if self.askActivateSpeechRecognition():
            self.openJsonFileNameDialog()
            if self.jsonFileName:
                self.setSettings(True)
            else:
                QMessageBox.warning(self, "No Google credentials",
                                    'No Google credentials file has been selected. Manually enter the path and file \
                                     name in cell H5 of the "Settings" sheet, after this installation is complete.')
        else:
            self.setSettings(False)
        if self.excelFileName:
            reply = QMessageBox.question(self, 'Setup Complete', 'Do you want to open the Excel file?',
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            try:
                if reply == QMessageBox.Yes:
                    os.system('start EXCEL.EXE '+str(self.excelFileName))
            except:
                QMessageBox.warning(self, "Unable to open file", "Unable to open file: "+self.execelFileName)
        sys.exit()

    def isExcelOpen(self):
        '''
        Checks if an instance of Microsoft Excel is running on this computer

        Returns
        -------
        bool
            True if an active Excel instance was found
        '''

        isOpen = False
        try:
            win32com.client.GetActiveObject("Excel.Application")
            isOpen = True
        except:
            isOpen = False
        finally:
            return isOpen

    def askActivateSpeechRecognition(self):
        '''
        Asks the user if speech recognition must be enabled
        '''

        try:
            reply = QMessageBox.question(self, 'Speech recognition', 'Do you want to enable speech recognition',
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                self.incompleteExit()
            else:
                return reply == QMessageBox.Yes
        except:
            self.incompleteExit()

    def openJsonFileNameDialog(self):
        '''
        Asks the user for the name of the Google Cloud credentials file
        '''

        try:
            reply = QMessageBox.question(self, 'Google cloud API key',
                                        'You will now be asked to select your Google JSON credential key file. Follow the instructions at the following link if you do not yet have a key file.\n\r\n\r https://cloud.google.com/speech-to-text/docs/before-you-begin \n\r\n\r Do you want to proceed with this installation?',
                                        QMessageBox.Yes | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                self.incompleteExit()
            homeDir = Path.home()
            fileName, _ = QFileDialog.getOpenFileName(self, "Select Google speech API credentials file",
                                                      str(homeDir), "(*.json)")
            if fileName:
                fileName = fileName.replace('/', '\\')
            self.jsonFileName = fileName
        except:
            self.incompleteExit()

    def openXlsmFileNameDialog(self):
        """
        Asks the user for the name and folder under which the xlsm file must be save
        """

        try:
            homeDir = Path.home()
            fileName, _ = QFileDialog.getSaveFileName(self, 'Where would you like to save a copy of GradeBook.xlsm',
                                                      str(homeDir)+'\\Documents\\GradeBook.xlsm', "(*.xlsm)")
            if fileName:
                fileName = fileName.replace('/', '\\')
            else:
                self.incompleteExit()
            self.excelFileName = fileName
        except:
            self.incompleteExit()

    def saveGradeBookCopy(self):
        """
        Saves a copy of the GradeBook.xlsm template
        """

        try:
            self.openXlsmFileNameDialog()
            if self.excelFileName:
                shutil.copyfile(resourcePath('GradeBook.xlsm'), self.excelFileName)
            else:
                self.incompleteExit()
        except:
            self.incompleteExit()

    def setSettings(self, enableSpeechRecognition):
        """
        Writes the settings to the Settings sheet of the workbook
        """

        try:
            pathToExe = resourcePath('')
            xl = win32com.client.Dispatch("Excel.application")
            xl.Visible = False
            wb = xl.Workbooks.Open(self.excelFileName)
            ws = wb.Sheets('Settings')
            ws.Range('H5').Value = pathToExe
            if self.jsonFileName:
                ws.Range('H4').Value = self.jsonFileName
            else:
                ws.Range('H4').Value = ""
            if enableSpeechRecognition:
                ws.Range('B2').Value = 'Yes'
            else:
                ws.Range('B2').Value = 'No'
            wb.Close(SaveChanges=1)
            xl.Quit()
        except:
            QMessageBox.critical(self, 'Error setting settings',
                                'Error while writing to "Settings" worksheet.')
            self.incompleteExit()

    def incompleteExit(self):
        """
        Displays a message to inform the user that Setup is incomplete and exits the application
        """

        if self.numTimesIncompleteExitCalled == 0:
            QMessageBox.warning(self, "Incomplete Installation",
                            'This installation is incomplete. Exiting this application.')
            self.numTimesIncompleteExitCalled = 1
        sys.exit()

def main():
    app = QApplication(sys.argv)
    MainW()
    sys.exit(app.exec_())

if __name__ == '__main__':
    a = main()
    a.show()

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
# The password for the VBA project is: Answer42
#********************************************************************************************************************************************************************************

from __future__ import absolute_import
from __future__ import division
import os
import sys
import threading
import time
from typing import List
import pyaudio  # Look at https://stackoverflow.com/questions/52283840/i-cant-install-pyaudio-on-windows-how-to-solve-error-microsoft-visual-c-14
import pywintypes
import win32com.client
import win32event
import win32api
from six.moves import queue
import platform
import getsettings as settings
import subprocess
import logging
import traceback
import threading
import time
import psutil
import pywintypes
import win32file
from google.cloud import speech
from google.cloud import texttospeech  # pip install google-cloud-texttospeech
import google.api_core
import winsound
from typing import List
from google.cloud import speech
import google.api_core
from winerror import ERROR_ALREADY_EXISTS
from pathlib import Path
import getsettings
import audioengine
import speechtotext
import pipeclient
import texttospeech

# Important for PyInstaller
# Create hook-grpc.py in the hooks folder (\Lib\site - packages\PyInstaller\hooks) and put the following code in it:
# from PyInstaller.utils.hooks import collect_data_files
# datas = collect_data_files ( 'grpc' )

# Also: pip install https://github.com/rokm/pyinstaller/archive/refs/heads/python-3.10.zip

# IP address to ping in order to check internet connection
googleIp = '8.8.8.8'

# Audio sampling parameters
sampleRate = 16000  # 16 kHz sampling rate
chunkSize = 1600    # 100ms chunks

# Pipe to receive messages from Excel
pipeName = r'\\.\pipe\TuinXSlang172'
# Time out connecting to pipe in seconds
pipeTimeOut = 10
# Time to wait between calling VBA macros in seconds
excelMacroWaitTime = 0.5

# Global variables
xl = None                                       # Handle to Excel app
wb = None                                       # Handle to Excel workbook
excelFileName = None                            # Excel file name. Passed as argument when Excel call this app
textToSpeech = None                             # Instantiation of the TextToSpeech class
pipeClient = None                               # Instantiation of the PipeClient class
audioEngine = None                              # Instantiation of the AudioEngine class
speechToText = None                             # Instantiation of the SpeechToText class
excelMacroBuffer = queue.Queue()                # Queue for messages send to Excel by calling Excel macros
prevExcelCallMacroTime = time.time()

appRunning = True
timeSpeechRecognitionResumed = time.time()      # Time at which speech recognition was resumed

# Turn logging on
try:
    homeDir = Path.home()
    logger = logging.getLogger(__name__)
    logging.basicConfig(filename=str(homeDir)+'\\EasyGradexl.log', filemode='w', level=logging.DEBUG)
except:
    pass


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

def isExcelOpen()->bool:
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

def checkIfOnlyInstance()->bool:
    '''
        Checks if there are more instances of GradeBook.exe running on this computer

        Returns
        -------
    bool
        True if another instance was found
        '''
    count = 0
    try:
        for process in psutil.process_iter():
            if "GradeBook.exe" in str(process.name):
                count += 1
    except:
        callExcelMacro('pythonError', 'noInternetConnection')
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)
    finally:
        return count > 1

def ping(host: str):
    """
    Returns True if host responds to a ping request.

    Parameters
    ----------
     host : str
        Host IP address
    """

    try:
        result = False
        param = '-n' if platform.system().lower() == 'windows' else '-c'
        command = ['ping', param, '1', host]
        CREATE_NO_WINDOW = 0x08000000
        subprocess.call(command, creationflags=CREATE_NO_WINDOW)     # Call once. Otherwise we may get an incorrect value
        count = 1
        if (not result) and count <= 50:          # Try for 5 s
                time.sleep(0.1)
                result = subprocess.call(command, creationflags=CREATE_NO_WINDOW) == 0
                count += 1
        if not result:
                callExcelMacro("pythonError", 'noInternetConnection')
    except:
        callExcelMacro('pythonError', 'noInternetConnection')
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)

def setGoogleCredentials(pathToGoogleCredentials: str):
    """
    Sets the environmental variable for Google speech API credentials

    Parameters
    ----------
    pathToGoogleCredentials : str
        The complete file path to the .jsn file that contains the Google API credentials
    """
    try:
        if os.path.isfile(pathToGoogleCredentials):     # Check if the file exists
            os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = settings.pathToGoogleCredentials
        else:
            callExcelMacro('pythonError', 'NoGoogleCredentialsFile')
    except:
        callExcelMacro('pythonError', 'NoGoogleCredentialsFile')
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)

def callExcelMacro(macroName: str, argument1:str =None, argument2:str =None):
    """
    Places a macro call to Excel in the Excel message buffer queue. Excel can be in blocking mode and may not be
    able to run macro immediately. That is why the call is placed in a queue.

    Parameters
    ----------
    macroName : str
        Name of the Excel macro to call
    argument1 : str, optional
        First argument
    argument2 : str, optional
        Second argument
    """
    # Place the parameters in a dictionary
    try:
        dictItem = {
            "macroName": str(macroName),
            "argument1": (argument1),
            "argument2": (argument2),
            "timeStamp": time.time()
        }
        excelMacroBuffer.put(dictItem)      # Add the dictionary to the queue
        processExcelMacros()                # Call the method to send the macro calls to Execl
    except:
        logException(sys.exc_info())

def processExcelMacros():
    """
        Checks for Excel macro calls in the excelMacroBuffer queue and send them to Excel. It will only
        remove the call from the queue if no exception occurred when trying to execute the macro.
    """
    global xl, excelFileName, prevExcelCallMacroTime
    # Check for incoming calls
    if not excelMacroBuffer.empty():
        # Do not remove the message from the queue unless Excel runs the Macro
        lst = list(excelMacroBuffer.queue)  # Convert queue to a list
        length = len(lst)
        while length >= 1:
            timeNow = time.time()
            if time.time()-prevExcelCallMacroTime < excelMacroWaitTime:
                break
            prevExcelCallMacroTime = timeNow
            dictItem = lst[0]    # Get the next item
            macroName = dictItem["macroName"]
            argument1 = dictItem["argument1"]
            argument2 = dictItem["argument2"]
            returnValue = None
            try:
                if argument1 == None:
                    returnValue = xl.Application.Run("'" + excelFileName + "'" + '!' + macroName)
                elif argument2 == None:
                    returnValue = xl.Application.Run("'" + excelFileName + "'" + '!' + macroName, str(argument1))
                else:
                    returnValue = xl.Application.Run("'" + excelFileName + "'" + '!' + macroName, str(argument1), str(argument2))
            except BaseException as ex:
                exceptionCode = ex.args[0]
                # Experience has shown that the following list of exceptions can be ignored
                ignoreNumbers = [-2147418111, -2147221008, -2147352565, -2147417842]
                if not exceptionCode in ignoreNumbers: raise
            except:
                exType, exValue, exTraceback = sys.exc_info()
                exceptionHook(exType, exValue, exTraceback)
            else:
                if returnValue != None:
                    if returnValue == 1:
                        excelMacroBuffer.get(block=True, timeout=0.1)
                        # Remove the call from the queue if no error has occurred
                        lst = list(excelMacroBuffer.queue)
                        length = len(lst)

def processIncomingMessage(message: str):
    global audioEngine, speechToText, textToSpeech
    """
    Process incoming messages received from Excel trough the pipe. Calls the appropriate methods to execute commands
    from Excel
    
    Parameters
    ----------
    message : str 
        String containing the message
    """

    try:
        if message != None:
            if message == 'resumeSpeechRecognition':
                audioEngine.resume()            # Unpause the audio engine
                speechToText.resume()           # Unpause speech-to-text conversion
                timeSpeechRecognitionResumed    # Log the time
            if message == 'pauseSpeechRecognition':
                audioEngine.pause()         # Pause the audio engine
                speechToText.pause()        # Pause the speech-to-text conversion
            if 'textMessage#' in message:   # Text-to-speech messages will contain 'textMessage#'
                x = message.split('#', 1)   # Split at '#'
                if len(x) == 2:
                    textToSpeech.speak(x[1])    # Run text-to-speech conversion
            if message == 'killGradebookExe':
                quitApp()                   # Quit App
    except Exception as ex:
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)

def exceptionHookThreading(arguments, /):
    """
    Catches exceptions occurring in threads
    """
    stringMicrophoneDisconnected = 'grpc._channel._MultiThreadedRendezvous: <_MultiThreadedRendezvous of RPC that\
     terminated with:\n\tstatus = StatusCode.UNKNOWN\n\tdetails\
      = "Exception iterating requests!"\n\tdebug_error_string = "None"\n>\n'
    stringInternetDisconnectedTextToSpeech = 'grpc._channel._InactiveRpcError: <_InactiveRpcError of RPC that\
     terminated with:\n\tstatus = StatusCode.UNKNOWN\n\tdetails = "Stream removed"'
    stringInternetDisconnectedSpeechToText = 'google.api_core.exceptions.Unknown: None Stream removed'
    try:
        tb = arguments.exceptionTraceback
        etype = arguments.exc_type
        value = arguments.exc_value
        traceBack = traceback.TracebackException(etype, value, tb, limit=None, lookup_lines=True,
                                     capture_locals=False, compact=False)
        gen = (traceBack.format(chain=True))
        for x in gen:
            if stringMicrophoneDisconnected in x:
                # This is a hack. This exception occurs spuriously just after speech recognition is resumed.
                elapsedTime = timeSpeechRecognitionResumed-time.time()
                if elapsedTime > 1:      # Allow one second
                    audioEngine.pause()
                    speechToText.pause()
                    callExcelMacro("pythonError", 'Microphone disconnected')
                break
            if 'The pipe has been ended.' in x:
                quitApp()
            if stringInternetDisconnectedTextToSpeech in x:
                callExcelMacro("pythonError", 'InterNetConnectionLost')
            if stringInternetDisconnectedSpeechToText in x:
                callExcelMacro("pythonError", 'InterNetConnectionLost')
    except:
        logException(sys.exc_info())

# Redirect exceptions in threads to exceptionHookThreading
threading.excepthook = exceptionHookThreading   # Catch exceptions is threads

def logException(ExceptionInfo):
    """
        Writes exceptions to log file
    """

    try:
        ExceptionType, ExceptionValue, exceptionTraceback = ExceptionInfo
        logger.critical("Unhandled exception", exc_info=(ExceptionType, ExceptionValue, exceptionTraceback))
    except:
        pass

def exceptionHook(exceptionType, exceptionValue, exceptionTraceback):
    """
        Catches all unhandled exceptions

    Parameters
    ----------
    exceptionType : Exception type
    exceptionValue : Exception's value
    exceptionTraceback : Exception's traceback
    """
    global appRunning
    try:
        if appRunning:
            errorMessage = f"Type: {exceptionType}\n" \
                           f"Value: {exceptionValue}\n"
            logger.critical("Unhandled exception", exc_info=(exceptionType, exceptionValue, exceptionTraceback))
            arguments = str(exceptionValue.args)
            if '-9996' in arguments:
                callExcelMacro("pythonError", 'noMicrophone')
                quitApp()
            if '-9988' in arguments:
                callExcelMacro("pythonError", 'Pipe Closed')
                quitApp()
            if 'Timeout error opening pipe' in arguments:
                callExcelMacro("pythonError", 'unableToConnectToPipe')
                quitApp()
            if '_InactiveRpcError' in arguments:
                callExcelMacro("pythonError", 'Google API error.')
                quitApp()
            if "<class 'google.auth.exceptions.DefaultCredentialsError'>" == exceptionType:
                callExcelMacro("pythonError", 'Google API error.')
                quitApp()
            if 'Unable to find setting' in arguments:
                callExcelMacro("pythonError", 'settingsSheetNotFound')
                quitApp()
            callExcelMacro("pythonException", errorMessage)
            quitApp()
    except:
        pass

# Redirect all unhandled exceptions to exceptionHook
BackupExceptionHook = sys.excepthook
sys.excepthook = exceptionHook

def quitApp():
    """
    Quits application
    """
    global appRunning
    global BackupException
    appRunning = False
    # Restore exception hook
    sys.excepthook = BackupExceptionHook
    stopTime = time.time() + 1  # Process Excel macros for 1 seconds
    try:
        while time.time() < stopTime:
            processExcelMacros()
            time.sleep(0.25)    # Wait for 0.25 seconds
    except:
        pass
    sys.exit(0)

def main():
    """
    Initializes the App and runs the main loop.
    """
    global wb, xl, excelFileName, settings, pipeClient, audioEngine, speechToText, textToSpeech, appRunning
    try:
        # Check if this computer is connected to the internet
        ping(googleIp)
        # Check if excel is running
        if not isExcelOpen():
            quitApp()
        # Check if this is the only running instance of this App
        if checkIfOnlyInstance():
            callExcelMacro('pythonError', 'InstanceAlreadyRunning')
            quitApp()
    except:
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)
    try:
        arguments = sys.argv
        excelFileName = arguments[1]
        # Excel app and workbook
        xl = win32com.client.Dispatch("Excel.application")
        wb = xl.Workbooks.Open(excelFileName)
        # Get settings from Settings sheet in workbook
        settings = getsettings.GetSettings(wb, callExcelMacro, exceptionHook)
        # Open pipe
        pipeClient = pipeclient.PipeClient(pipeName, processIncomingMessage, pipeTimeOut, exceptionHook)
        pipeClient.openPipe()
        # Set Google Cloud API credentials
        setGoogleCredentials(settings.pathToGoogleCredentials)
        processExcelMacros()
        # Instantiate objects
        audioEngine = audioengine.AudioEngine(sampleRate, chunkSize)
        speechToText = speechtotext.SpeechToText(audioEngine, settings.speechToTextLanguageCode, settings.commonPhrases,
                                                 settings.replacementDictionary, callExcelMacro, sampleRate, exceptionHook)
        textToSpeech = texttospeech.TextToSpeech(settings.textToSpeechLanguageCode, settings.textToSpeechVoice,
                                                 settings.textToSpeechGender, settings.SsmlSayAs, callExcelMacro,
                                                 exceptionHook)
        # Pause speech recognition
        audioEngine.pause()
        speechToText.pause()
    except:
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)
    # Main loop
    try:
        checkEveryNumCycles = 999999                  # Check every 1000000 cycles if Excel is still open
        count = checkEveryNumCycles
        while appRunning:
            speechToText.runSpeechRecognition()
            pipeClient.readFromExcel()
            processExcelMacros()
            textToSpeech.runTextToSpeech()
            if count == checkEveryNumCycles:
                if not isExcelOpen():   # Check if Excel is running
                    break
                count = 0
            else:
                count += 1
    except:
        exType, exValue, exTraceback = sys.exc_info()
        exceptionHook(exType, exValue, exTraceback)
    quitApp()

if __name__ == '__main__':
    main()


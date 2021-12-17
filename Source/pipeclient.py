import threading
import time
import sys
import pywintypes
import win32file

class PipeClient(object):
    """ Connects as a client to a pipe to receive messages from Excel

        Methods
        _______
            openPipe
                Connects to the pipe a get a pipe handle
            readFromExcel
                Must be run periodically to receive messages from Excel

    """

    def __init__(self, pipeName: str, processIncomingMessages: callable, timeOut: int, exceptionHook: callable):
        """
        Parameters
        ----------
        pipeName : str
            Name of the pipe that was opened by Excel
        processIncomingMessages : callable
            Handle to method that must be called to process the incoming messages
        timeOut : int
            The number of seconds to retry opening the pipe if open fails
        exceptionHook : callable
            Handle to the method that handles exceptions
        """
        self.__pipeName = pipeName
        self.__processIncomingMessages = processIncomingMessages
        self.__timeOut = timeOut
        self.__exceptionHook = exceptionHook
        self.__pipeHandle = None        # Handle to pipe
        self.__incomingText = None      # Text that was received through pipe
        self.__readThread = None        # Thread in which the pipe read process runs

    def openPipe(self):
        """
        Opens the receiving end of the pipe

        Raises
        ------
        RuntimeError
        If the connecting to the pipe timed out
        """
        success = False
        stopTime = time.time()+self.__timeOut
        try:
            while time.time() < stopTime and not success:
                try:
                        # See http://timgolden.me.uk/pywin32-docs/win32file__CreateFile_meth.html
                        self.__pipeHandle = win32file.CreateFile(
                            self.__pipeName,
                            win32file.GENERIC_READ,
                            0,
                            None,
                            win32file.OPEN_EXISTING,
                            0,
                            None
                        )
                        success = True
                except pywintypes.error as ex:
                    if ex.args == (2, 'CreateFile', 'The system cannot find the file specified.'):
                        time.sleep(0.1)     # Wait 0.1s before trying another time
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)
        finally:
            if not success:
                self.__pipeHandle = None
                raise RuntimeError('Timeout error opening pipe')

    def __readFromPipe(self):
        """
        Reads text from incoming pipe and puts it in self.__incomingText. This method runs
        in its own thread.
        """
        if self.__pipeHandle != None:
            result, str = win32file.ReadFile(self.__pipeHandle, 10240, None)  # Read 10 KB
            if result == 0:     # No error
                self.__incomingText = str.decode("utf-16")  # Convert to UTF-16 format
            else:   # Error
                self.__incomingText = None
        else:
            self.__incomingText = None

    def readFromExcel(self):
        """
        Runs the thread that reads data from the pipe. Must be called regularly.
        """
        try:
            if self.__readThread == None:   # Thread has not yet been created
                self.__readThread = threading.Thread(target=self.__readFromPipe, daemon=True)
                self.__readThread.start()
            else:
                if not self.__readThread.is_alive():    # Check if the thread is alive
                    messages = self.__incomingText
                    if messages != None:
                        messages = messages.split('\r\n')   # Messages are separated by '\r\n'
                        for message in messages:
                            message = message.replace('\r\n', '')       # Remove '\r\n'
                            self.__processIncomingMessages(message)     # Process incoming text
                    self.__readThread = threading.Thread(target=self.__readFromPipe, daemon=True)    # Restart thread
                    self.__readThread.start()
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)

    def __del__(self):
        """
        Closes the pipe handle
        """
        try:
            self.__pipeHandle.close()
        except:
            pass


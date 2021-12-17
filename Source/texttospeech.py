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
#********************************************************************************************************************************************************************************

from __future__ import absolute_import
from __future__ import division
from google.cloud import speech
from google.cloud import texttospeech  # pip install google-cloud-texttospeech
import google.api_core
import threading
import winsound
import sys
from six.moves import queue

# Settings
textToSpeechTimeout = 60.0      # Timeout in seconds
maxQueueSize = 10               # Maximum number of text messages in the queue

class TextToSpeech(object):
    """
    Synthesizes speech from the input string of text using the Google text-to-speech API

    Methods
    _______
    speak(text)
        Converts the string of text to speech and plays it through the default audio output device
    """
    def __init__(self, textToSpeechLanguageCode: str, textToSpeechVoice: str, textToSpeechGender: str, ssmlSayAs: str, callExcelMacro: callable, exceptionHook: callable):
        """
        Parameters
        ----------
        textToSpeechLanguageCode : str
            The Google API language code as selected on the "Settings" sheet
        textToSpeechVoice :
            The Google voice as selected on the "Settings" sheet
        textToSpeechGender : str
            The Google API gender as selected on the "Settings" sheet
        ssmlSayAs : str
            The Google API SSML "say as" as selected in the "Settings" sheet
        callExcelMacro : callable
            Method to call in order to send message to Excel
        exceptionHook : callable
            Method to call when an exception has occurred
        """
        self.__textToSpeechLanguageCode = textToSpeechLanguageCode
        self.__textToSpeechVoice = textToSpeechVoice
        self.__textToSpeechGender = textToSpeechGender
        self.__ssmlSayAs = ssmlSayAs
        self.__callExcelMacro = callExcelMacro
        self.__exceptionHook = exceptionHook
        self.__text = ''
        self.__thread = None            # Thread in which text to speech runs
        self.__queue = queue.Queue()    # Queue messages to read back
        # Instantiate text-to-speech a client
        try:
            self.__client = google.cloud.texttospeech.TextToSpeechClient()
        except:
            self.__callExcelMacro("pythonError", 'GoogleCredentialsError.')
        try:
            # Set the voice gender
            if self.__textToSpeechGender == 'neutral':
                gender = google.cloud.texttospeech.SsmlVoiceGender.NEUTRAL
            else:
                gender = self.__textToSpeechGender
            # Set the voice configuration
            self.__voice = google.cloud.texttospeech.VoiceSelectionParams(
                language_code=self.__textToSpeechLanguageCode,
                name=self.__textToSpeechVoice,
                ssml_gender=gender
            )
            # Select the type of audio file to return
            self.__audio_config = google.cloud.texttospeech.AudioConfig(audio_encoding=google.cloud.texttospeech.AudioEncoding.LINEAR16)
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)

    def __speakGoogle(self):
        """
        Converts the self.__text to speech and plays it through the default audio output device
        """
        try:
            if self.__text != '':
                # Replace - with minus character
                text1 = self.__text.replace('-', '\u2212')
                # Add SSML markups
                ssmlText = '<speak> <say-as interpret-as=\"'+self.__ssmlSayAs+'\">'+text1+'</say-as> </speak>'
                charactersBilled = str(len(ssmlText))
                # Set the text input to be synthesized
                synthesis_input = google.cloud.texttospeech.SynthesisInput(ssml=ssmlText)
                # Build the voice request, select the language code and the ssml voice gender.
                # Perform the text-to-speech request on the text input with the selected
                # voice parameters and audio file type.
                response = self.__client.synthesize_speech(
                    input=synthesis_input, voice=self.__voice, audio_config=self.__audio_config, timeout=textToSpeechTimeout
                )
                waveData = response.audio_content
                # This is in the form of a Wave file. We need to strip header before playing it
                self.__callExcelMacro("charactersBilled", str(charactersBilled))
                winsound.PlaySound(waveData, winsound.SND_MEMORY | winsound.SND_NOSTOP)
        except:
            ExceptionType, ExceptionValue, excpetionTraceback = sys.exc_info()
            arguments = str(ExceptionValue.args)
            if 'Failed to play sound' in arguments:
                self.__callExcelMacro('pythonError', 'FailedToPlaySound')
            else:
                exType, exValue, exTraceback = sys.exc_info()
                self.__exceptionHook(exType, exValue, exTraceback)

    def speak(self, text: str):
        """
        Places the text to be converted to speech in the queue and calls runTextToSpeech

        Parameters
        ----------
        text : str
            Text to be converted to speec
        """
        try:
            if self.__queue.qsize() <= maxQueueSize:
                self.__queue.put(text)
                self.runTextToSpeech
            else:
                self.__callExcelMacro('pythonError', 'textToSpeechQueueOverflow')
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)

    def runTextToSpeech(self):
        """
        Creates the thread in which to run text-to-speech conversion
        """
        try:
            if not self.__queue.empty():
                if self.__thread == None:
                    self.__text = self.__queue.get()
                    self.__thread = threading.Thread(target=self.__speakGoogle, daemon=True)
                    self.__thread.start()
                else:
                    if not self.__thread.is_alive():  # Restart the thread if it has completed
                        self.__text = self.__queue.get()
                        self.__thread = threading.Thread(target=self.__speakGoogle)
                        self.__thread.start()
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)

    def __del__(self):
        pass

#********************************************************************************************************************************************************************************
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
from typing import Dict, Any
import sys
""" Reads the settings from the "Settings" sheet of the Excel workbook

    Attributes
    ----------
    speechToTextLanguageCode : str
        Google Cloud speech-to-text API language code
    textToSpeechLanguageCode : str
        Google Cloud text-to-speech API language code
    textToSpeechSsml : str = ''
        The SSML string to be passed to the Google Cloud text-to-speech API 
    pathToGoogleCredentials : str
        Full path to the Google API json credentials file
    replacementDictionary : Dict[Any, Any]
        Dictionary of words that will be replaced in the output of the speech recognition
    commonPhrases : list
        Common phrase that will be sent to the Google Cloud speech-to-text API
    speechToTextLanguageCode : str
        Google Cloud API speech-to-text language code
    textToSpeechVoice : str
        Google Cloud API speech-to-text voice
    textToSpeechGender : str
        Google Cloud API speech-to-text gender
    SsmlSayAs : str = ''
        Google Cloud API say as string
"""
class GetSettings(object):
    def __init__(self, wb: object, callExcelMacro: callable, exceptionHook: callable):
        """
            Parameters
            ----------
            pipeName : wb
                The workbook object
            callExcelMacro : callable
                Handle to method that must be called to call Excel VBA macros
            exceptionHook : callable
                Handle to the method that handles exceptions
        """
        self.speechToTextLanguageCode: str = ''
        self.textToSpeechLanguageCode: str = ''
        self.textToSpeechSsml: str = ''
        self.pathToGoogleCredentials: str = ''
        self.replacementDictionary: Dict[Any, Any] = {}
        self.commonPhrases = []
        self.speechToTextLanguageCode: str = ''
        self.textToSpeechVoice: str = ''
        self.textToSpeechGender: str = ''
        self.SsmlSayAs: str = ''
        self.__wb: object
        self.__callExcelMacro: callable
        self.__exceptionHook: callable
        self.__wb = wb
        self.__callExcelMacro = callExcelMacro
        self.__exceptionHook = exceptionHook
        self.__getSettingsFromExcel()

    def __getLanguageCode(self, wb: object, selectedLanguage: str)->str:
        """
            Reads the language code from the "Settings" sheet of the workbook

            Parameters
            ----------
            wb : object
               The workbook object
            selectedLanguage : string
               The name of the selected language
            Returns
            -------
    l       string
                The Google Cloud API language code
       """

        try:
            wsLanguages = wb.Sheets('Languages')
            language = wsLanguages.Range('A1').Value
            languageCode = wsLanguages.Range('B1').Value
            found = str(selectedLanguage) == str(language)
            row = 2
            while not found and language != None and languageCode != None:
                language = wsLanguages.Range('A' + str(row)).Value
                languageCode = wsLanguages.Range('B' + str(row)).Value
                found = str(selectedLanguage) == str(language)
                row += 1
            return languageCode
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)

    def __getSettingsFromExcel(self):
        """
            Reads the settings from the "Settings" sheet in the Execl workbook
        """

        settingsFound = True
        try:
            ws = self.__wb.Sheets('Settings')
        except:
            settingsFound = False
            self.__callExcelMacro("pythonError", 'settingsSheetNotFound')
        try:
            if settingsFound:
                # Construct replacements dictionary
                key = ws.Range('D2').Value
                if key != None:
                    keyString = str(key)
                    keyString = keyString.replace('"', '')
                value = ws.Range('E2').Value
                if value != None:
                    valueString = str(value)
                    valueString = valueString.replace('"', '')
                if keyString != None:
                    row = 3
                    while key != None and value != None:
                        self.replacementDictionary[keyString] = valueString
                        key = ws.Range('D' + str(row)).Value
                        if key != None:
                            keyString = str(key)
                            keyString = keyString.replace('"', '')
                        value = ws.Range('E' + str(row)).Value
                        if value != None:
                            valueString = str(value)
                            valueString = valueString.replace('"', '')
                        row += 1
                # Get commonPhrases
                phrase = ws.Range('F2').Value
                if phrase != None:
                    phraseString = str(phrase)
                row = 3
                while phrase != None:
                    phraseString = phraseString.replace('"', '')
                    self.commonPhrases.append(phraseString)
                    phrase = ws.Range('F' + str(row)).Value
                    if phrase != None:
                        phraseString = str(phrase)
                    row += 1
                # Get speech-to-text language code
                selectedLanguage = ws.Range('B3').Value
                self.speechToTextLanguageCode = self.__getLanguageCode(self.__wb, selectedLanguage)
                # Get text-to-speech language code
                selectedLanguage = ws.Range('B4').Value
                self.textToSpeechLanguageCode = self.__getLanguageCode(self.__wb, selectedLanguage)
                # Get text to speech voice
                voiceAndGender = str(ws.Range('B5').Value)
                if '(' in voiceAndGender:
                    voiceAndGenderSplit = voiceAndGender.split('(')
                    voice = voiceAndGenderSplit[0].strip()
                    self.textToSpeechVoice = voice
                    gender = voiceAndGenderSplit[1].replace(')', '').strip()
                    self.textToSpeechGender = gender
                else:
                    self.textToSpeechVoice = voiceAndGender
                    self.textToSpeechGender = 'neutral'
                # Get SSML <say as>
                self.SsmlSayAs = str(ws.Range('B6').Value)
                # Get path to Google API key
                self.pathToGoogleCredentials = str(ws.Range('H4').Value)
        except:
            exType, exValue, exTraceback = sys.exc_info()
            self.__exceptionHook(exType, exValue, exTraceback)
        finally:
            if not settingsFound:
                self.__callExcelMacro("pythonError", "settingsSheetNotFound")

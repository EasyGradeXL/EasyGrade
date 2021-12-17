@RD /S /Q dist
@RD /S /Q build
pyinstaller -y --noconsole --icon=PineAppleOriginal.ico Complete.spec
copy GradeBook.xlsm dist\GradeBook\
copy *.wav dist\GradeBook\
copy *.jpg dist\GradeBook\
copy *.rtf dist\GradeBook\
copy "..\User Manual\EasyGradeXLUserManual.pdf" dist\GradeBook\
cd dist\SetupEasyGrade\
copy *.* ..\GradeBook /s
cd..
cd..
@RD /S /Q dist\SetupEasyGrade
@RD /S /Q build

'========================================================�ļ�·����������================================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'=============================================================�������=======================================================================

Sub OnClick()
    
    FilePath = SSProcess.SelectFileName(1,"ѡ���ļ�",0,"All Files (*.*)|*.*||")

    LineOne = "cd /d %~dp0"
    LineTwo = FilePath & " fgdbtomdb2.py"

    fGDBtoMdbFilePath = SSProcess.GetSysPathName(7) & "fGDBtoMdb.bat"
    
    Set fGDBtoMdbFile = FileSysObj.CreateTextFile(fGDBtoMdbFilePath,True)

    fGDBtoMdbFile.WriteLine(LineOne)
    fGDBtoMdbFile.WriteLine(LineTwo)
End Sub ' OnClick

'========================================================文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'=============================================================功能入口=======================================================================

Sub OnClick()
    
    FilePath = SSProcess.SelectFileName(1,"选择文件",0,"All Files (*.*)|*.*||")

    LineOne = "cd /d %~dp0"
    LineTwo = FilePath & " fgdbtomdb2.py"

    fGDBtoMdbFilePath = SSProcess.GetSysPathName(7) & "fGDBtoMdb.bat"
    
    Set fGDBtoMdbFile = FileSysObj.CreateTextFile(fGDBtoMdbFilePath,True)

    fGDBtoMdbFile.WriteLine(LineOne)
    fGDBtoMdbFile.WriteLine(LineTwo)
End Sub ' OnClick
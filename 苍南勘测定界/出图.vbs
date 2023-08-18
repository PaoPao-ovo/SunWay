
'========================================================文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'复制后路径字符串
Dim CopyPathStr
CopyPathStr = ""

'=============================================================功能入口=======================================================================

Function OpenProject()
    '复制征地范围面（504）到粘贴板
    CloneArea
    
    '选择文件的路径(多选之间以","进行分隔)
    FilePathStr = SSProcess.SelectFileName(1,"选择文件",1,"EDB Files (*.edb)|*.edb|All Files (*.*)|*.*||")
    
    FilePathArr = Split(FilePathStr,",", - 1,1)
    
    '复制工程
    For i = 0 To UBound(FilePathArr)
        Set EdbFile = FileSysObj.GetFile(FilePathArr(i))
        EdbFile.Copy SSProcess.GetSysPathName(5) & "报告成果\" & EdbFile.Name
        If CopyPathStr = "" Then
            CopyPathStr = SSProcess.GetSysPathName(5) & "报告成果\" & EdbFile.Name
        Else
            CopyPathStr = CopyPathStr & "," & SSProcess.GetSysPathName(5) & "报告成果\" & EdbFile.Name
        End If
    Next 'i
    
    '打开选择的工程并粘贴
    CopyPathArr = Split(CopyPathStr,",", - 1,1)
    For i = 0 To UBound(CopyPathArr)
        SSProcess.OpenDatabase CopyPathArr(i)
        SSProcess.AddClipBoardObjToMap 0,0
    Next 'i
    
End Function' OpenProject()

'选择要素复制到粘贴板
Function CloneArea()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard
End Function' CloneArea
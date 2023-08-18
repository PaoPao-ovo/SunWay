
'========================================================文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'=============================================================功能入口=======================================================================

Sub OnClick()
    
    SelFolderPath = SSProcess.SelectPathName()
    
    Set SelFolder = FileSysObj.GetFolder(SelFolderPath)
    
    SavePath = SSProcess.GetSysPathName(5) & SelFolder.Name
    
    SelFolder.Copy SavePath,True
    
    Set Folder = FileSysObj.GetFolder(SavePath)
    
    Set AllFiles = Folder.Files
    
    ImgType = "JPG 文件,JPEG 文件,PNG 文件"
    
    For Each CurrentFile In AllFiles
        If InStr(ImgType,CurrentFile.Type) = 0 Then
            CurrentFile.Delete
        End If
    Next ' CurrentFile
    
    For Each CurrentFile In AllFiles
        If CurrentFile.Type = "JPEG 文件" Then
            SavePath = Replace(CurrentFile.Path,CurrentFile.Name,"")
            CurrentName = UCase(CurrentFile.Name)
            SaveName = Replace(CurrentName,".JPEG",".JPG")
            CurrentFile.Copy SavePath & SaveName
            CurrentFile.Delete
        End If
    Next ' CurrentFile
    
    For Each CurrentFile In AllFiles
        If CurrentFile.Type = "PNG 文件" Then
            SavePath = Replace(CurrentFile.Path,CurrentFile.Name,"")
            CurrentName = UCase(CurrentFile.Name)
            SaveName = Replace(CurrentName,".PNG",".JPG")
            CurrentFile.Copy SavePath & SaveName
            CurrentFile.Delete
        End If
    Next ' CurrentFile
End Sub' OnClick
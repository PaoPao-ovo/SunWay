
'=======================================================功能入口====================================================================

Sub OnClick()
    
    '返回ZIP文件名称
    SaveName SaveZipName,Result
    If Result = 1 Then
        YsFolder SaveZipName
    Else
        Exit Sub
    End If
End Sub' OnClick

'获取保存文件名称
Function SaveName(ByRef SaveZipName,ByRef Result)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "保存文件名" , "XXX竣工测量" , 0 , "XXX土地勘测定界,XXX建设工程规划放线,XXX正负零检测,XXX不动产与测绘,XXX竣工多测合一,XXX不动产实测绘,XXX水印及电子章" , ""
    Result = SSProcess.ShowInputParameterDlg ("输入文件名称")
    SaveZipName = SSProcess.GetInputParameter("保存文件名")
    SSProcess.RefreshView
End Function' SaveName


Function YsFolder(ByVal ZipName)
    SelFolderPath = SSProcess.SelectPathName() '选择的文件夹路径
    SavePath = SSProcess.GetSysPathName(5) & ZipName & ".zip"
    ArchiveFolder SavePath,SelFolderPath
End Function' YsFolder

Sub ArchiveFolder (zipFile, sFolder)
    With CreateObject("Scripting.FileSystemObject")
        With .CreateTextFile(zipFile, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))
        End With
    End With
    With CreateObject("Shell.Application")
        .NameSpace(zipFile).CopyHere .NameSpace(sFolder).Items
        Do Until .NameSpace(zipFile).Items.Count = _
            .NameSpace(sFolder).Items.Count
            SSProcess.Sleep 1000
        Loop
    End With
End Sub
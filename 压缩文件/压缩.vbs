
'=======================================================�������====================================================================

Sub OnClick()
    
    '����ZIP�ļ�����
    SaveName SaveZipName,Result
    If Result = 1 Then
        YsFolder SaveZipName
    Else
        Exit Sub
    End If
End Sub' OnClick

'��ȡ�����ļ�����
Function SaveName(ByRef SaveZipName,ByRef Result)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "�����ļ���" , "XXX��������" , 0 , "XXX���ؿ��ⶨ��,XXX���蹤�̹滮����,XXX��������,XXX����������,XXX��������һ,XXX������ʵ���,XXXˮӡ��������" , ""
    Result = SSProcess.ShowInputParameterDlg ("�����ļ�����")
    SaveZipName = SSProcess.GetInputParameter("�����ļ���")
    SSProcess.RefreshView
End Function' SaveName

Function YsFolder(ByVal ZipName)
    SelFolderPath = SSProcess.SelectPathName() 'ѡ����ļ���·��
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
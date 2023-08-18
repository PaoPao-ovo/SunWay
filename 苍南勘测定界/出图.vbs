
'========================================================�ļ�·����������================================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'���ƺ�·���ַ���
Dim CopyPathStr
CopyPathStr = ""

'=============================================================�������=======================================================================

Function OpenProject()
    '�������ط�Χ�棨504����ճ����
    CloneArea
    
    'ѡ���ļ���·��(��ѡ֮����","���зָ�)
    FilePathStr = SSProcess.SelectFileName(1,"ѡ���ļ�",1,"EDB Files (*.edb)|*.edb|All Files (*.*)|*.*||")
    
    FilePathArr = Split(FilePathStr,",", - 1,1)
    
    '���ƹ���
    For i = 0 To UBound(FilePathArr)
        Set EdbFile = FileSysObj.GetFile(FilePathArr(i))
        EdbFile.Copy SSProcess.GetSysPathName(5) & "����ɹ�\" & EdbFile.Name
        If CopyPathStr = "" Then
            CopyPathStr = SSProcess.GetSysPathName(5) & "����ɹ�\" & EdbFile.Name
        Else
            CopyPathStr = CopyPathStr & "," & SSProcess.GetSysPathName(5) & "����ɹ�\" & EdbFile.Name
        End If
    Next 'i
    
    '��ѡ��Ĺ��̲�ճ��
    CopyPathArr = Split(CopyPathStr,",", - 1,1)
    For i = 0 To UBound(CopyPathArr)
        SSProcess.OpenDatabase CopyPathArr(i)
        SSProcess.AddClipBoardObjToMap 0,0
    Next 'i
    
End Function' OpenProject()

'ѡ��Ҫ�ظ��Ƶ�ճ����
Function CloneArea()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard
End Function' CloneArea
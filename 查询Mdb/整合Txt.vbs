
'========================================================Doc����������ļ�·����������================================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'============================================================����SQL=========================================================================

'Txt�ļ���·��
TxtFolderPath = "C:\Users\ASUS\Desktop\�ı��ļ�"

'�½��ļ�����
TotalTxtName = "C:\Users\ASUS\Desktop\�������ļ�.txt"

'YSDM�ַ���
YSDMStr = ""

'===========================================�������==========================================================================================

'�����
Sub OnClick()
    
    GetAllTxt YSDMArr
    
    DelRepeat YSDMArr,ToTalVal,LxCount
    
    CreatTxt TotalTxtName,TotalTxtFile
    ToTalValArr = Split(ToTalVal,";", - 1,1)
    For i = 0 To UBound(ToTalValArr)
        TotalTxtFile.WriteLine ToTalValArr(i)
    Next 'i
    TotalTxtFile.Close
End Sub' OnClick

Function GetAllTxt(ByRef YSDMArr)
    Row = 1
    Set AllFolder = FileSysObj.GetFolder("C:\Users\ASUS\Desktop\�ı��ļ�")
    Set AllFiles = AllFolder.Files
    For Each SingleTxtFile In AllFiles
        Set TextFile = FileSysObj.OpenTextFile(SingleTxtFile.Path)
        Do While Not TextFile.AtEndOfStream
            LineContent = Trim(TextFile.ReadLine)
            If LineContent <> "" Then
                If YSDMStr = "" Then
                    YSDMStr = LineContent
                Else
                    YSDMStr = YSDMStr & ";" & LineContent
                End If
            End If
        Loop
        TextFile.Close
    Next ' SingleTxtFile
    YSDMArr = Split(YSDMStr,";", - 1,1)
End Function' GetAllTxe

'����ı��ļ�
Function CreatTxt(ByVal TextFileName,ByRef TxtFile)
    Set TxtFile = FileSysObj.CreateTextFile(TextFileName,True)
End Function' CreatTxt

'ȥ���ַ������ظ�ֵ
Function DelRepeat(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr)
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & ";" & "'" & StrArr(i) & "'"
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,";", - 1,1)) + 1
End Function' DelRepeat
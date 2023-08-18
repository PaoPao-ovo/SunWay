
'========================================================Doc操作对象和文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'============================================================配置SQL=========================================================================

'Txt文件夹路径
TxtFolderPath = "C:\Users\ASUS\Desktop\文本文件"

'新建文件名称
TotalTxtName = "C:\Users\ASUS\Desktop\总配置文件.txt"

'YSDM字符串
YSDMStr = ""

'===========================================功能入口==========================================================================================

'总入口
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
    Set AllFolder = FileSysObj.GetFolder("C:\Users\ASUS\Desktop\文本文件")
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

'输出文本文件
Function CreatTxt(ByVal TextFileName,ByRef TxtFile)
    Set TxtFile = FileSysObj.CreateTextFile(TextFileName,True)
End Function' CreatTxt

'去除字符串中重复值
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
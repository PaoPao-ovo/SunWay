'======================================================��鼯����=====================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "���۵�Ψһ�Լ��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->���۵�Ψһ�Լ��"

'�������
Dim strDescription
strDescription = "����ͬ�������۵�"

'==================================================ʵ����������=========================================================

'ʵ���Ԥ���Ӧ��ϵ��
' ʵ�����            ����                ���۱���
' 9130512           GPS����            1103021
' 9130412           ˮ׼��                1102021
' 9130311           ���Ƶ㣨��ʯ��         9130211
' 9130312           ���Ƶ㣨����ʯ��         9130212
' 9130217           ��վ��                9130216
' 9130511           ������               9130411


LLCodes = "1103021,1102021,9130211,9130212,9130216,9130411"

'===================================================��������==========================================================


'��ں���
Sub OnClick()
    ClearCheckRecord()
    ExportRecords LLCodes
End Sub' OnClick

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'�����鼯
Function ExportRecords(codes)
    SelLlPoi codes
    SelCount = SSProcess.GetSelGeoCount()
    If SelCount > 0 Then
        StrName = ""
        For i = 0 To SelCount - 1
            poiname = SSProcess.GetSelGeoValue(i,"SSObj_PointName")
            x = SSProcess.GetSelGeoValue(i,"SSObj_X")
            y = SSProcess.GetSelGeoValue(i,"SSObj_Y")
            id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            If StrName = "" Then
                StrName = "'" & poiname & "'"
            ElseIf Replace(StrName,"'" & poiname & "'","") = StrName Then
                StrName = StrName & "," & "'" & poiname & "'"
            Else
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,0,id, ""
            End If
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'ѡ�����ۿ��Ƶ�
Function SelLlPoi(Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
End Function' CheckRealPoi
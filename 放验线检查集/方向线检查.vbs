'======================================================��鼯����=====================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "�����߼��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->�����߼��"

'�������
Dim strDescription
strDescription = "�����ߵĲ�վ��Ų��ڼ������"

'==================================================ʵ����������=========================================================

' ������ 9130251 CeZDH
' ����� 9130241 CeZDH

FxLine = "9130251"
JcLine = "9130241"

'===================================================��������==========================================================


'��ں���
Sub OnClick()
    ClearCheckRecord()
    ExportRecords()
End Sub' OnClick

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'�����鼯
Function ExportRecords()
    SelJcLine()
    SelCount = SSProcess.GetSelGeoCount()
    ReDim JcArr(SelCount)
    If SelCount > 0 Then
        For i = 0 To SelCount - 1
            JcArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        Next 'i
        For i = 0 To SelCount - 1
            CeName = SSProcess.GetObjectAttr(JcArr(i),"[CeZDH]")
            x = SSProcess.GetObjectAttr(JcArr(i),"SSObj_X")
            y = SSProcess.GetObjectAttr(JcArr(i),"SSObj_Y")
            count = GetFxCount(CeName)  
            'MsgBox CeName
            If count = 0 Then
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y,0,1,JcArr(i), ""
            End If
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'ѡ�����м����
Function SelJcLine()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130241"
    SSProcess.SelectFilter
End Function' SelJcLine

'���Ϸ�������Ŀ
Function GetFxCount(CeName)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130251"
    SSProcess.SetSelectCondition "[CeZDH]", "==", CeName
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount()
    GetFxCount = Count
End Function' GetFxCount
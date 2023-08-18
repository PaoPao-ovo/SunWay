'======================================================��鼯����=====================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߼��"

'��鼯������
Dim strCheckName
strCheckName = "ʵ����ظ����"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->ʵ����ظ����"

'�������
Dim strDescription
strDescription = "ͬ�������۵㲻����"

'==================================================ʵ����������=========================================================

'ʵ���Ԥ���Ӧ��ϵ��
' ʵ�����            ����                ���۱���
' 9130512           GPS����            1103021
' 9130412           ˮ׼��                1102021
' 9130311           ���Ƶ㣨��ʯ��         9130211
' 9130312           ���Ƶ㣨����ʯ��         9130212
' 9130217           ��վ��                9130216
' 9130511           ������               9130411


ScdCodes = "9130512,9130412,9130311,9130312,9130217,9130511"

'===================================================��������==========================================================


'��ں���
Sub OnClick()
    ClearCheckRecord()
    ExportRecords ScdCodes
End Sub' OnClick

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'�����鼯
Function ExportRecords(codes)
    ScArr = Split(ScdCodes,",", - 1,1)
    For i = 0 To UBound(ScArr)
        SelRealPoi ScArr(i)
        SelCount = SSProcess.GetSelGeoCount()
        ReDim Selids(SelCount,2)
        If SelCount > 0  Then
            For j = 0 To SelCount - 1
                Selids(j,0) = SSProcess.GetSelGeoValue(j,"SSObj_ID")
                Selids(j,1) = SSProcess.GetSelGeoValue(j,"SSObj_PointName")
            Next 'j
            For k = 0 To SelCount - 1
                Select Case ScArr(i)
                    Case "9130512"
                    SelLlPoi "1103021",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        'MsgBox geoType
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName,strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130412"
                    SelLlPoi "1102021",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130311"
                    SelLlPoi "9130211",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130312"
                    SelLlPoi "9130212",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130217"
                    SelLlPoi "9130216",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130511"
                    SelLlPoi "9130411",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                End Select
            Next 'k
        End If
    Next 'i
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'ѡ��ʵ���
Function SelRealPoi(Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
End Function' CheckRealPoi

Function SelLlPoi(Code,poiname)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SetSelectCondition "SSObj_PointName", "==", poiname
    SSProcess.SelectFilter
End Function' SelLlPoi
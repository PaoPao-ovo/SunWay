
'��ʼ��
Sub OnInitScript()
    
    
    ClearSelection '���ԭ��ϵͳѡ��
    
    Mode = 0 '=0 �޲����Ի��� =1 �в����Ի���

    Title = "������"
    
    SSProcess.ShowScriptDlg Mode,Title
    
End Sub

'����رպ�ִ��
Sub OnExitScript()
    '��Ӵ���
End Sub

'�����ɺ�ִ��
Sub OnOK()
    
    UpdateSelection SelCount,LineIdArr '����ѡ�񼯣�����ѡ���������ID����
    
    For i = 0 To UBound(LineIdArr)
        
        GXQDDH = SSProcess.GetObjectAttr(LineIdArr(i),"[GXQDDH]")
        GXZDDH = SSProcess.GetObjectAttr(LineIdArr(i),"[GXZDDH]")
        
        GetPointXY GXQDDH,GXZDDH,StartX,StartY,EndX,EndY '��ȡͬ���ĵ��X��Yֵ
        
        SetLineXY LineIdArr(i),StartX,StartY,EndX,EndY '�޸���λ��
        
    Next 'i
    
    SSProcess.RefreshView()
    
End Sub


'ȡ����ִ��
Sub OnCancel()
    '��Ӵ���
End Sub

'���ϵͳѡ��
Function ClearSelection()
    SSProcess.ClearSysSelection
End Function' ClearSelection

'��ϵͳѡ�񼯸��µ��ű���
Function UpdateSelection(ByRef SelCount,ByRef LineIdArr)
    
    SSProcess.UpdateSysSelection 0 'ѡ�񼯸���
    SelCount = SSProcess.GetSelGeoCount()
    ReDim LineIdArr(SelCount - 1)
    For i = 0 To SelCount - 1
        LineIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
    
End Function' UpdateSelection

'��ȡ�������
Function GetPointXY(ByVal GXQDDH,ByVal GXZDDH,ByRef StartX,ByRef StartY,ByRef EndX,ByRef EndY)
    
    SqlStr = "Select ���¹��ߵ����Ա�.ID From ���¹��ߵ����Ա� inner join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And WTDH = " & "'" & GXQDDH & "'"
    
    GetSQLRecordAll SqlStr,StartPointArr,Count
    
    If Count > 0 Then
        StartX = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_X")
        StartY = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_Y")
    End If
    
    SqlStr = "Select ���¹��ߵ����Ա�.ID From ���¹��ߵ����Ա� inner join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And WTDH = " & "'" & GXZDDH & "'"
    
    GetSQLRecordAll SqlStr,StartPointArr,Count
    
    If Count > 0 Then
        EndX = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_X")
        EndY = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_Y")
    End If
    
End Function' GetPointXY

'�޸���λ��
Function SetLineXY(ByVal LineID,ByVal StartX,ByVal StartY,ByVal EndX,ByVal EndY)

    Pointcount = Transform(SSProcess.GetObjectAttr(LineID,"SSObj_PointCount"))
    If Pointcount = 2 Then
        SSProcess.GetObjectPoint LineID,0,StartLineX,StartLineY,StartLineZ,StartPointType,StartName
        SSProcess.GetObjectPoint LineID,1,EndLineX,EndLineY,EndLineZ,EndPointType,EndName
    End If
    
    If StartX <> "" And StartY <> "" Then
        SSProcess.SetObjectAttr LineID,"SSObj_X(0)",StartX
        SSProcess.SetObjectAttr LineID,"SSObj_Y(0)",StartY
    End If
    
    If EndX <> "" And EndY <> "" Then
        SSProcess.SetObjectAttr LineID,"SSObj_X(1)",EndX
        SSProcess.SetObjectAttr LineID,"SSObj_Y(1)",EndY
    End If

End Function' SetLineXY

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

'��������ת��
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

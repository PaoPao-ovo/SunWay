'�������
Sub OnClick()
    AllVisible
    ZrzInner
    HxInner
    MsgBox "�������"
End Sub' OnClick

'=========================================================��Ȼ�����Ի���========================================================================

'����������������Ȼ������ZRZGUID��
'���¡�SCDXJZMJ��=����顾SJC��<0 �ġ�JZMJ��+����顾SJC��>0 �ҡ�MJKMC��������������¡��ġ�KZMJ��
'���ϡ�SCQTJZMJ��=����顾SJC��>0�ҡ�MJKMC����������������¡��ġ�JZMJ��
'��SCJZMJ�� = ���µ������֮��
'סլ�����ZhuZMJ��=����SYGN��������סլ���ġ�JZMJ��
'��סլ�����FZhuZMJ��=���ġ�SYGN����������סլ���ġ�JZMJ��+�����ġ�FTLX��Ϊ����̯�ġ�KZMJ��
'��ZTS��=��ͬ��ZRZGUID���Ļ��ĸ���

'��Ȼ����ں���
Function ZrzInner()
    GetZrzGUID "9210123",ZrzCount,ZrzArr
    GetSCDXMJ ZrzCount,ZrzArr
    GetSCQTJZMJ ZrzCount,ZrzArr
    GetZongJzMj ZrzCount,ZrzArr
    SetZhuZhaiMj ZrzCount,ZrzArr
    SetFZhuZhaiMj ZrzCount,ZrzArr
    SetHCount ZrzCount,ZrzArr
End Function' ZrzInner

'����ÿ����Ȼ����ʵ����½��������ˢ��
Function GetSCDXMJ(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210413",MjkCount
    GetSCDXMJ = 0
    For i = 0 To ZrzCount - 1
        For j = 0 To MjkCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") Then
                If Transform(SSProcess.GetSelGeoValue(j,"[SJC]")) < 0 Then
                    If GetSCDXMJ = 0 Then
                        GetSCDXMJ = Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                    Else
                        GetSCDXMJ = GetSCDXMJ + Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                    End If
                ElseIf  IsConTain(SSProcess.GetSelGeoValue(j,"[MJKMC]"),"�������") = 1 And Transform(SSProcess.GetSelGeoValue(j,"[SJC]")) > 0 Then
                    If GetSCDXMJ = 0 Then
                        GetSCDXMJ = Round(Transform(SSProcess.GetSelGeoValue(j,"[KZMJ]")),2)
                    Else
                        GetSCDXMJ = GetSCDXMJ + Round(Transform(SSProcess.GetSelGeoValue(j,"[KZMJ]")),2)
                    End If
                End If
            End If
        Next 'j
        SSProcess.SetObjectAttr ZrzArr(i,1),"[SCDXJZMJ]",Round(GetSCDXMJ,2)
        GetSCDXMJ = 0
    Next 'i
End Function' GetSCDXMJ

'����ÿ����Ȼ����ʵ����Ͻ��������ˢ��
Function GetSCQTJZMJ(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210413",MjkCount
    GetSCQTJZMJ = 0
    For i = 0 To ZrzCount - 1
        For j = 0 To MjkCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") Then
                If Transform(SSProcess.GetSelGeoValue(j,"[SJC]")) > 0 And IsConTain(SSProcess.GetSelGeoValue(j,"[MJKMC]"),"�������") = 0 Then
                    If GetSCQTJZMJ = 0 Then
                        GetSCQTJZMJ = Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]"))
                    Else
                        GetSCQTJZMJ = GetSCQTJZMJ + Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]"))
                    End If
                End If
            End If
        Next 'j
        SSProcess.SetObjectAttr ZrzArr(i,1),"[SCQTJZMJ]",Round(GetSCQTJZMJ,2)
        GetSCQTJZMJ = 0
    Next 'i
End Function' GetSCQTJZMJ

'�����ܽ��������ˢ��
Function GetZongJzMj(ByVal ZrzCount,ByVal ZrzArr())
    For i = 0 To ZrzCount - 1
        SSProcess.SetObjectAttr ZrzArr(i,1),"[SCJZMJ]",Round(Transform(SSProcess.GetObjectAttr(ZrzArr(i,1),"[SCQTJZMJ]")) + Transform(SSProcess.GetObjectAttr(ZrzArr(i,1),"[SCDXJZMJ]")),2)
    Next 'i
End Function' GetZongJzMj

'������Ȼ����סլ���
Function SetZhuZhaiMj(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210513",HCount
    SetZhuZhaiMj = 0
    For i = 0 To  ZrzCount - 1
        For j = 0 To HCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") And IsConTain(SSProcess.GetSelGeoValue(j,"[SYGN]"),"סլ") = 1 Then
                If SetZhuZhaiMj = 0 Then
                    SetZhuZhaiMj = Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                Else
                    SetZhuZhaiMj = SetZhuZhaiMj + Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                End If
            End If
        Next 'j
        SSProcess.SetObjectAttr ZrzArr(i,1),"[ZhuZMJ]",Round(SetZhuZhaiMj,2)
        SetZhuZhaiMj = 0
    Next 'i
End Function' SetZhuZhaiMj

'������Ȼ���ķ�סլ���
Function SetFZhuZhaiMj(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210513",HCount
    SetFZhuZhaiMj = 0
    For i = 0 To  ZrzCount - 1
        For j = 0 To HCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") And IsConTain(SSProcess.GetSelGeoValue(j,"[SYGN]"),"סլ") = 0 Then
                If SetFZhuZhaiMj = 0 Then
                    SetFZhuZhaiMj = Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                Else
                    SetFZhuZhaiMj = SetFZhuZhaiMj + Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                End If
            End If
        Next 'j
        SqlString = "Select Sum(FC_�������Ϣ���Ա�.KZMJ) From FC_�������Ϣ���Ա� inner join GeoAreaTB on FC_�������Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_�������Ϣ���Ա�.ZRZGUID = " & ZrzArr(i,0) & " AND FC_�������Ϣ���Ա�.FTLX = '����̯'"
        GetSQLRecordAll SqlString,BftMjArr,BftMjCount
        SetFZhuZhaiMj = SetFZhuZhaiMj + Round(BftMjArr(0),2)
        SSProcess.SetObjectAttr ZrzArr(i,1),"[FZhuZMJ]",Round(SetFZhuZhaiMj,2)
        SetFZhuZhaiMj = 0
    Next 'i
End Function ' SetFZhuZhaiMj

'��ȡ��������ˢ������
Function SetHCount(ByVal ZrzCount,ByVal ZrzArr())
    SetHCount = 0
    SelFeatures "9210513",HCount
    For i = 0 To ZrzCount - 1
        For j = 0 To HCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") Then SetHCount = SetHCount + 1
        Next 'j
        SSProcess.SetObjectAttr ZrzArr(i,1),"[ZTS]",SetHCount
        SetHCount = 0
    Next 'i
End Function' SetHCount

'��ȡ��Ȼ������ZRZGUID
Function GetZrzGUID(ByVal Code,ByRef ZrzCount,ByRef ZrzArr())
    SelFeatures Code,ZrzCount
    ReDim ZrzArr(ZrzCount,2)
    For i = 0 To ZrzCount - 1
        ZrzArr(i,0) = SSProcess.GetSelGeoValue(i,"[ZRZGUID]")
        ZrzArr(i,1) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' GetZrzGUID

'================================================================�������Ի���=================================================================

'���ܵ����ߡ�ZDGUID��
'סլ��ZZRFYJMJ��=���䷶Χ�ڵ���Ȼ����ZhuZMJ��֮�͡���ZhuZXS��
'��סլ��QTRFYJMJ��=���䷶ΧΪ����Ȼ����FZhuZMJ��֮�͡���FZhuZXS��
'���Ͻ��������JunGCLDSJZMJ��=���䷶ΧΪ����Ȼ����SCQTJZMJ��֮��
'���½��������JunGCLDXJZMJ��=���䷶ΧΪ����Ȼ����SCDXJZMJ��֮��
'�ܽ��������JunGCLZJZMJ��= ���ϼӵ���
'���ϲ�����DSCS��=���䷶ΧΪ����Ȼ��[DSCS]�����ֵ
'������JunGCLZTS��=���䷶ΧΪ����Ȼ���ġ�ZTS��

'������ں���
Function HxInner()
    GetHxGUID "9130223",HxCount,HxArr
    SetRFZhuZhai HxArr
End Function' HxInner

'��ȡ�������Ժ�ZDGUID
Function GetHxGUID(ByVal Code,ByRef HxCount,ByRef HxArr())
    SelFeatures Code,HxCount
    If HxCount = 1 Then
        ReDim HxArr(1,2)
        HxArr(0,0) = SSProcess.GetSelGeoValue(i,"[ZDGUID]")
        HxArr(0,1) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Else
        MsgBox "ͼ���ж������"
        Exit Function
    End If
End Function' GetHxGUID

'��������ˢ��
Function SetRFZhuZhai(ByVal HxArr())
    SelFeatures "9210123",ZrzCount
    ZhuZMJ = 0
    FZhuZMJ = 0
    SCQTJZMJ = 0
    SCDXJZMJ = 0
    ZTS = 0
    MsxCs = 0
    For i = 0 To ZrzCount - 1
        If  HxArr(0,0) = SSProcess.GetSelGeoValue(i,"[ZDGUID]") Then
            If ZhuZMJ = 0 Then
                ZhuZMJ = Round(Transform(SSProcess.GetSelGeoValue(i,"[ZhuZMJ]")) * ToDecimal(SSProcess.GetObjectAttr(HxArr(0,1),"[ZhuZXS]")),2)
            Else
                ZhuZMJ = ZhuZMJ + Round(Transform(SSProcess.GetSelGeoValue(i,"[ZhuZMJ]")) * ToDecimal(SSProcess.GetObjectAttr(HxArr(0,1),"[ZhuZXS]")),2)
            End If
            
            If FZhuZMJ = 0 Then
                FZhuZMJ = Round(Transform(SSProcess.GetSelGeoValue(i,"[FZhuZMJ]")) * ToDecimal(SSProcess.GetObjectAttr(HxArr(0,1),"[FZhuZXS]")),2)
            Else
                FZhuZMJ = FZhuZMJ + Round(Transform(SSProcess.GetSelGeoValue(i,"[FZhuZMJ]")) * ToDecimal(SSProcess.GetObjectAttr(HxArr(0,1),"[FZhuZXS]")),2)
            End If
            
            If SCQTJZMJ = 0 Then
                SCQTJZMJ = Round(Transform(SSProcess.GetSelGeoValue(i,"[SCQTJZMJ]")),2)
            Else
                SCQTJZMJ = SCQTJZMJ + Round(Transform(SSProcess.GetSelGeoValue(i,"[SCQTJZMJ]")),2)
            End If
            
            If SCDXJZMJ = 0 Then
                SCDXJZMJ = Round(Transform(SSProcess.GetSelGeoValue(i,"[SCDXJZMJ]")),2)
            Else
                SCDXJZMJ = SCDXJZMJ + Round(Transform(SSProcess.GetSelGeoValue(i,"[SCDXJZMJ]")),2)
            End If
            
            If ZTS = 0 Then
                ZTS = Transform(SSProcess.GetSelGeoValue(i,"[ZTS]"))
            Else
                ZTS = ZTS + Transform(SSProcess.GetSelGeoValue(i,"[ZTS]"))
            End If
            
            If MsxCs < Transform(SSProcess.GetSelGeoValue(i,"[DSCS]")) Then MsxCs = Transform(SSProcess.GetSelGeoValue(i,"[DSCS]"))
        End If
    Next 'i
    SSProcess.SetObjectAttr HxArr(0,1),"[ZZRFYJMJ]",ZhuZMJ
    SSProcess.SetObjectAttr HxArr(0,1),"[QTRFYJMJ]",FZhuZMJ
    SSProcess.SetObjectAttr HxArr(0,1),"[JunGCLDSJZMJ]",SCQTJZMJ
    SSProcess.SetObjectAttr HxArr(0,1),"[JunGCLDXJZMJ]",SCDXJZMJ
    SSProcess.SetObjectAttr HxArr(0,1),"[JunGCLZJZMJ]",SCDXJZMJ + SCQTJZMJ
    SSProcess.SetObjectAttr HxArr(0,1),"[JunGCLZTS]",ZTS
    SSProcess.SetObjectAttr HxArr(0,1),"[DSCS]",MsxCs
End Function' SetRFZhuZhai

'=====================================================���ߺ���================================================================================

'�ٷֺ�תС��
Function ToDecimal(ByVal Percentage)
    ToDecimal = Transform(Left(Percentage,Len(Percentage) - 1)) * 0.01
End Function' ToDecimal

'��ͼ��
Function AllVisible()
    LayerCount = SSProcess.GetLayerCount
    For i = 0 To LayerCount - 1
        LayerName = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus LayerName, 1, 1
    Next
    SSProcess.RefreshView
End Function'AllVisible

'ѡ��ָ�����ﲢ���ظ���
Function SelFeatures(ByVal Code,ByRef Count)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
End Function' SelFeatures

'�ж��Ƿ����ָ���ַ���
Function IsConTain(ByVal TempStr,ByVal ReplaceValue)
    If Replace(TempStr,ReplaceValue,"") = TempStr Then
        IsConTain = 0
    Else
        IsConTain = 1
    End If
End Function' IsConTain

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

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (SSProcess.GetProjectFileName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst SSProcess.GetProjectFileName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (SSProcess.GetProjectFileName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord SSProcess.GetProjectFileName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext SSProcess.GetProjectFileName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
End Function
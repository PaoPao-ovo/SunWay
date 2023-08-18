'功能入口
Sub OnClick()
    AllVisible
    ZrzInner
    HxInner
    MsgBox "汇总完成"
End Sub' OnClick

'=========================================================自然幢属性汇总========================================================================

'汇总面积块面积到自然幢，【ZRZGUID】
'地下【SCDXJZMJ】=面积块【SJC】<0 的【JZMJ】+面积块【SJC】>0 且【MJKMC】包含“计入地下”的【KZMJ】
'地上【SCQTJZMJ】=面积块【SJC】>0且【MJKMC】不包含“计入地下”的【JZMJ】
'【SCJZMJ】 = 地下地上面积之和
'住宅面积【ZhuZMJ】=户【SYGN】包含“住宅”的【JZMJ】
'非住宅面积【FZhuZMJ】=户的【SYGN】不包含“住宅”的【JZMJ】+面积块的【FTLX】为不分摊的【KZMJ】
'【ZTS】=相同【ZRZGUID】的户的个数

'自然幢入口函数
Function ZrzInner()
    GetZrzGUID "9210123",ZrzCount,ZrzArr
    GetSCDXMJ ZrzCount,ZrzArr
    GetSCQTJZMJ ZrzCount,ZrzArr
    GetZongJzMj ZrzCount,ZrzArr
    SetZhuZhaiMj ZrzCount,ZrzArr
    SetFZhuZhaiMj ZrzCount,ZrzArr
    SetHCount ZrzCount,ZrzArr
End Function' ZrzInner

'计算每个自然幢的实测地下建筑面积并刷新
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
                ElseIf  IsConTain(SSProcess.GetSelGeoValue(j,"[MJKMC]"),"计入地下") = 1 And Transform(SSProcess.GetSelGeoValue(j,"[SJC]")) > 0 Then
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

'计算每个自然幢的实测地上建筑面积并刷新
Function GetSCQTJZMJ(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210413",MjkCount
    GetSCQTJZMJ = 0
    For i = 0 To ZrzCount - 1
        For j = 0 To MjkCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") Then
                If Transform(SSProcess.GetSelGeoValue(j,"[SJC]")) > 0 And IsConTain(SSProcess.GetSelGeoValue(j,"[MJKMC]"),"计入地下") = 0 Then
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

'计算总建筑面积并刷新
Function GetZongJzMj(ByVal ZrzCount,ByVal ZrzArr())
    For i = 0 To ZrzCount - 1
        SSProcess.SetObjectAttr ZrzArr(i,1),"[SCJZMJ]",Round(Transform(SSProcess.GetObjectAttr(ZrzArr(i,1),"[SCQTJZMJ]")) + Transform(SSProcess.GetObjectAttr(ZrzArr(i,1),"[SCDXJZMJ]")),2)
    Next 'i
End Function' GetZongJzMj

'设置自然幢的住宅面积
Function SetZhuZhaiMj(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210513",HCount
    SetZhuZhaiMj = 0
    For i = 0 To  ZrzCount - 1
        For j = 0 To HCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") And IsConTain(SSProcess.GetSelGeoValue(j,"[SYGN]"),"住宅") = 1 Then
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

'设置自然幢的非住宅面积
Function SetFZhuZhaiMj(ByVal ZrzCount,ByVal ZrzArr())
    SelFeatures "9210513",HCount
    SetFZhuZhaiMj = 0
    For i = 0 To  ZrzCount - 1
        For j = 0 To HCount - 1
            If ZrzArr(i,0) = SSProcess.GetSelGeoValue(j,"[ZRZGUID]") And IsConTain(SSProcess.GetSelGeoValue(j,"[SYGN]"),"住宅") = 0 Then
                If SetFZhuZhaiMj = 0 Then
                    SetFZhuZhaiMj = Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                Else
                    SetFZhuZhaiMj = SetFZhuZhaiMj + Round(Transform(SSProcess.GetSelGeoValue(j,"[JZMJ]")),2)
                End If
            End If
        Next 'j
        SqlString = "Select Sum(FC_面积块信息属性表.KZMJ) From FC_面积块信息属性表 inner join GeoAreaTB on FC_面积块信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 AND FC_面积块信息属性表.ZRZGUID = " & ZrzArr(i,0) & " AND FC_面积块信息属性表.FTLX = '不分摊'"
        GetSQLRecordAll SqlString,BftMjArr,BftMjCount
        SetFZhuZhaiMj = SetFZhuZhaiMj + Round(BftMjArr(0),2)
        SSProcess.SetObjectAttr ZrzArr(i,1),"[FZhuZMJ]",Round(SetFZhuZhaiMj,2)
        SetFZhuZhaiMj = 0
    Next 'i
End Function ' SetFZhuZhaiMj

'获取户个数并刷新属性
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

'获取自然幢数和ZRZGUID
Function GetZrzGUID(ByVal Code,ByRef ZrzCount,ByRef ZrzArr())
    SelFeatures Code,ZrzCount
    ReDim ZrzArr(ZrzCount,2)
    For i = 0 To ZrzCount - 1
        ZrzArr(i,0) = SSProcess.GetSelGeoValue(i,"[ZRZGUID]")
        ZrzArr(i,1) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' GetZrzGUID

'================================================================红线属性汇总=================================================================

'汇总到红线【ZDGUID】
'住宅【ZZRFYJMJ】=在其范围内的自然幢【ZhuZMJ】之和×【ZhuZXS】
'非住宅【QTRFYJMJ】=在其范围为的自然幢【FZhuZMJ】之和×【FZhuZXS】
'地上建筑面积【JunGCLDSJZMJ】=在其范围为的自然幢【SCQTJZMJ】之和
'地下建筑面积【JunGCLDXJZMJ】=在其范围为的自然幢【SCDXJZMJ】之和
'总建筑面积【JunGCLZJZMJ】= 地上加地下
'地上层数【DSCS】=在其范围为的自然幢[DSCS]的最大值
'户数【JunGCLZTS】=在其范围为的自然幢的【ZTS】

'红线入口函数
Function HxInner()
    GetHxGUID "9130223",HxCount,HxArr
    SetRFZhuZhai HxArr
End Function' HxInner

'获取红线属性和ZDGUID
Function GetHxGUID(ByVal Code,ByRef HxCount,ByRef HxArr())
    SelFeatures Code,HxCount
    If HxCount = 1 Then
        ReDim HxArr(1,2)
        HxArr(0,0) = SSProcess.GetSelGeoValue(i,"[ZDGUID]")
        HxArr(0,1) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Else
        MsgBox "图上有多个红线"
        Exit Function
    End If
End Function' GetHxGUID

'红线属性刷新
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

'=====================================================工具函数================================================================================

'百分号转小数
Function ToDecimal(ByVal Percentage)
    ToDecimal = Transform(Left(Percentage,Len(Percentage) - 1)) * 0.01
End Function' ToDecimal

'打开图层
Function AllVisible()
    LayerCount = SSProcess.GetLayerCount
    For i = 0 To LayerCount - 1
        LayerName = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus LayerName, 1, 1
    Next
    SSProcess.RefreshView
End Function'AllVisible

'选择指定地物并返回个数
Function SelFeatures(ByVal Code,ByRef Count)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
End Function' SelFeatures

'判断是否包含指定字符串
Function IsConTain(ByVal TempStr,ByVal ReplaceValue)
    If Replace(TempStr,ReplaceValue,"") = TempStr Then
        IsConTain = 0
    Else
        IsConTain = 1
    End If
End Function' IsConTain

'数据类型转换
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
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
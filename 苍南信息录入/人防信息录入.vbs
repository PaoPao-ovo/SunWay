
' strfieldsName = "项目名称,项目地址,项目编号,设计单位,建设单位,委托单位,测绘单位,测绘时间,战时功能,平时功能,审批人防应建面积,住宅系数,非住宅系数,建筑结构"
' strfields = "XiangMMC,XiangMDZ,XiangMBH,SheJDW,JianSDW,WeiTDW,CeLDW,CeLRQ,ZSGN,PSGN,SPRFYJMJ,ZhuZXS,FZhuZXS,JZJG"
strfieldsName = "测绘单位,测绘时间,战时功能,平时功能,审批人防应建面积,住宅系数,非住宅系数,建筑结构,项目名称,项目地址,项目编号,设计单位,建设单位,委托单位"
strfields = "CeLDW,CeLRQ,ZSGN,PSGN,SPRFYJMJ,ZhuZXS,FZhuZXS,JZJG,XiangMMC,XiangMDZ,XiangMBH,SheJDW,JianSDW,WeiTDW"
Sub OnClick()
    
    geocount = GetFeatureCount( "9130223", geocount)
    strfieldsNameList = Split(strfieldsName,",")
    strfieldsList = Split(strfields,",")
    SSProcess.ClearInputParameter
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ObjID = id
        GetFieldValues id, strfields, strValues, fieldCount
        For i1 = 0 To fieldCount - 1
            If fieldCount - 1 = UBound(strfieldsNameList) Then
                If strfieldsNameList(i1) = "住宅系数" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "7%,8%", "县城住宅8%，重点镇住宅7%。"
                ElseIf strfieldsNameList(i1) = "非住宅系数" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "5%,4%", "县城住非住宅5%，重点镇非住宅4%。"
                ElseIf strfieldsNameList(i1) = "建筑结构" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "钢混", ""
                ElseIf strfieldsNameList(i1) = "测绘时间" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), GetNowTime,0, "", ""
                Else
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "", ""
                End If
            End If
        Next
    Next
    res = SSProcess.ShowInputParameterDlg ("人防项目信息录入" )
    If res = 1 Then
        For i = 0 To UBound(strfieldsNameList)
            value = SSProcess.GetInputParameter (strfieldsNameList(i))
            SSProcess.SetObjectAttr ObjID, "[" & strfieldsList(i) & "]", value
        Next
    End If
    SSProcess.ObjectDeal objID, "FreeDisplayList", parameters, result
    SSProcess.RefreshView
    
End Sub


'获取要素个数
Function GetFeatureCount(ByVal Code,ByRef geocount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code","==",Code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    GetFeatureCount = geocount
End Function

'获取字段属性数组
Function GetFieldValues(ByVal id,ByVal fields,ByRef strValues(),ByRef fieldCount)
    ReDim strValues(fieldCount)
    fieldCount = 0
    fieldsList = Split(fields,",")
    For i = 0 To UBound(fieldsList)
        values = SSProcess.GetObjectAttr (id, "[" & fieldsList(i) & "]")
        ReDim Preserve strValues(fieldCount)
        strValues(fieldCount) = values
        fieldCount = fieldCount + 1
    Next
End Function

'获取当前系统时间
Function GetNowTime()
    GetNowTime = FormatDateTime(Now(),1)
End Function' GetNowTime
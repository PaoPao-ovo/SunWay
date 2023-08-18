'======================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "验线检查"

'检查集组名称
Dim strCheckName
strCheckName = "验线高程检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->验线高程检查"

'检查描述
Dim strDescription
strDescription = "验线高程为空"

'==================================================实测点编码配置=========================================================

'实测和预测对应关系：
' 实测编码            点名                理论编码
' 9130512           GPS检测点            1103021
' 9130412           水准点                1102021
' 9130311           控制点（埋石）         9130211
' 9130312           控制点（不埋石）         9130212
' 9130217           测站点                9130216
' 9130511           放样点               9130411


ScdCodes = "9130224"

'===================================================函数主体==========================================================


'入口函数
Sub OnClick()
    ClearCheckRecord()
    ExportRecords ScdCodes
End Sub' OnClick

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'输出检查集
Function ExportRecords(code)
    SelRealPoi code
    SelCount = SSProcess.GetSelGeoCount()
    If SelCount > 0  Then
        polygonID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
        idstr = SSProcess.SearchInnerObjIDs(polygonID, 0, "9130611",0)
        idarr = Split(idstr,",", - 1,1)
        For i = 0 To UBound(idarr)
            yxgc = SSProcess.GetObjectAttr(idarr(i),"[YanXGC]")
            x = SSProcess.GetObjectAttr(idarr(i),"SSObj_X")
            y = SSProcess.GetObjectAttr(idarr(i),"SSObj_Y")
            z = SSProcess.GetObjectAttr(idarr(i),"SSObj_Z")
            yxgc = transform(yxgc)
            'MsgBox yxgc
            If yxgc = 0 Then SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0,0,idarr(i), ""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'选择实测点
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

'数据类型转换
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        content = 0
    End If
    transform = content
End Function
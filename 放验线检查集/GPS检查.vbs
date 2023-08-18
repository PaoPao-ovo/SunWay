'======================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "验线检查"

'检查集组名称
Dim strCheckName
strCheckName = "GPS检查点检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->GPS检查点检查"

'检查描述
Dim strDescription1
strDescription1 = "GPS检查点附近不存在理论控制点"
Dim strDescription2
strDescription2 = "GPS检查点附近存在多个理论控制点"
'==================================================实测点编码配置=========================================================

'实测和预测对应关系：
' 实测编码            点名                理论编码
' 9130512           GPS检测点            1103021
' 9130412           水准点                1102021
' 9130311           控制点（埋石）         9130211
' 9130312           控制点（不埋石）         9130212
' 9130217           测站点                9130216
' 9130511           放样点               9130411


ScdCodes = "9130215"

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
    ' Dim idarr(SelCount)
    If SelCount > 0 Then
        For i = 0 To SelCount - 1
            id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            x = SSProcess.GetSelGeoValue(i,"SSObj_X")
            y = SSProcess.GetSelGeoValue(i,"SSObj_Y")
            z = SSProcess.GetObjectAttr(i,"SSObj_Z")
            idstr = SSProcess.SearchNearObjIDs(x, y, 0.1, 0, "9130211,9130212", 0)
            idarr = Split(idstr,",",-1,1)
            nearcount = UBound(idarr) + 1
            If idstr = "" Then
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription1, x, y, 0, 0,id, ""
            End If
            If nearcount > 1 Then
                SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription2, x, y, 0, 0,id, ""
            End If
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
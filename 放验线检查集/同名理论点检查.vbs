'======================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "验线检查"

'检查集组名称
Dim strCheckName
strCheckName = "理论点唯一性检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->理论点唯一性检查"

'检查描述
Dim strDescription
strDescription = "存在同名的理论点"

'==================================================实测点编码配置=========================================================

'实测和预测对应关系：
' 实测编码            点名                理论编码
' 9130512           GPS检测点            1103021
' 9130412           水准点                1102021
' 9130311           控制点（埋石）         9130211
' 9130312           控制点（不埋石）         9130212
' 9130217           测站点                9130216
' 9130511           放样点               9130411


LLCodes = "1103021,1102021,9130211,9130212,9130216,9130411"

'===================================================函数主体==========================================================


'入口函数
Sub OnClick()
    ClearCheckRecord()
    ExportRecords LLCodes
End Sub' OnClick

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'输出检查集
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

'选择理论控制点
Function SelLlPoi(Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
End Function' CheckRealPoi
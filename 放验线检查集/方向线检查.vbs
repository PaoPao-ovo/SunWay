'======================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "验线检查"

'检查集组名称
Dim strCheckName
strCheckName = "方向线检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->方向线检查"

'检查描述
Dim strDescription
strDescription = "方向线的测站点号不在检查线中"

'==================================================实测点编码配置=========================================================

' 方向线 9130251 CeZDH
' 检查线 9130241 CeZDH

FxLine = "9130251"
JcLine = "9130241"

'===================================================函数主体==========================================================


'入口函数
Sub OnClick()
    ClearCheckRecord()
    ExportRecords()
End Sub' OnClick

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'输出检查集
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

'选择所有检查线
Function SelJcLine()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130241"
    SSProcess.SelectFilter
End Function' SelJcLine

'符合方向线数目
Function GetFxCount(CeName)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130251"
    SSProcess.SetSelectCondition "[CeZDH]", "==", CeName
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount()
    GetFxCount = Count
End Function' GetFxCount
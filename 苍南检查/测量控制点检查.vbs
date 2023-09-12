
'=============================================检查集配置==============================================

'检查组项目名称
Dim strGroupName
strGroupName = "测量控制点属性值必填检查"

'检查集项名称
Dim strCheckName
strCheckName = "测量控制点属性值必填检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->测量控制点属性值必填检查"

'检查描述
Dim strDescription
strDescription = "测量控制点FCODE,NAME,GRADE,XYCOOR,X,Y,ELEVATION,ZCOOR属性值有为空"

'========================================检查字段配置=================================================

Dim FildsName
FildsName = "FCODE,NAME,GRADE,XYCOOR,X,Y,ELEVATION,ZCOOR"

'==========================================功能主体===================================================

'功能入口
Sub OnClick()
    
    AllVisible

    ClearCheckRecord
    
    SelFeatures "测量控制点",IdCount,IdArr
    RecordsInner IdCount,IdArr
    
    Ending

End Sub' OnClick

'=========================================检查集输出=============================================

'检查集输出入口
Function RecordsInner(ByVal IdCount,ByVal IdArr())
    ExportRecords IdCount,IdArr
End Function' RecordsInner

'输出检查集
Function ExportRecords(ByVal IdCount,ByVal IdArr())
    For i = 0 To IdCount - 1
        If IsEmpty(IdArr(i),FildsName) Then
            AddCheckRecord IdArr(i)
        End If
    Next 'i
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ExportRecords

'=========================================工具类函数===========================================

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'选择要素
Function SelFeatures(ByVal LayerName,ByRef TotalCount,ByRef AllIdArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount
    ReDim AllIdArr(TotalCount)
    For i = 0 To TotalCount - 1
        AllIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelAllFeatures

'判断宽度字段是否为空
Function IsEmpty(ByVal Id,ByVal KeyString)
    SplitKeyString KeyString,KeyArr,KeyCount
    IsEmpty = False
    For i = 0 To KeyCount - 1
        If SSProcess.GetObjectAttr(Id,KeyArr(i)) = "" Then IsEmpty = True
    Next 'i
End Function' IsEmpty

'分解键字符串
Function SplitKeyString(ByVal StringName,ByRef StrArr(),ByRef StrCount)
    StrArr = Split(StringName,",", - 1,1)
    StrCount = UBound(StrArr) + 1
    For i = 0 To StrCount - 1
        StrArr(i) = "[" & StrArr(i) & "]"
    Next 'i
End Function' SplitKeyString

'添加单条记录
Function AddCheckRecord(ByVal Id)
    SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(Id,"SSObj_X"),SSProcess.GetObjectAttr(Id,"SSObj_Y"),0,1,Id,""
End Function' AddCheckRecord

'结束函数
Function Ending()
    MsgBox "检查完成"
End Function' Ending
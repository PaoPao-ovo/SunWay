
' 1、刷新通用属性中的数据源和更新时间，更新时间选择当前系统下的时间
' 2、若通用属性有值则按原值，若没有值则按照缺省值刷新

'=====================================================键值对配置==============================================================

'通用属性字段
Dim CommonString
CommonString = "UPDATEDATE,DATASOURCE,FEATURESTATUS"

'通用属性缺省值
Dim CommonValString
CommonValString = GetNowTime & "," & "1:500竣工地形图,2"

'==========================================================功能入口==================================================================

'功能入口
Sub OnClick()
    
    AllVisible
    
    SelAllFeatures IdCount,IdArr
    CommonAttrInner IdCount,IdArr
    
    Ending
    
End Sub' OnClick

'================================================================通用属性刷新=================================================================

'通用属性刷新入口
Function CommonAttrInner(ByVal IdCount,ByVal IdArr())
    UpDateAttribute IdCount,IdArr
End Function' CommonAttrInner

'通用属性刷新函数
Function UpDateAttribute(ByVal IdCount,ByVal IdArr())
    SplitKeyString CommonString,CommonArr,CommonCount
    SplitString CommonValString,ValArr,ValCount
    For i = 0 To IdCount - 1
        For j = 0 To ValCount - 1
            If SSProcess.GetObjectAttr(IdArr(i),CommonArr(j)) = "" Then
                SSProcess.SetObjectAttr IdArr(i),CommonArr(j),ValArr(j)
            End If
        Next 'j
    Next 'i
End Function' UpDateAttribute

'=================================================================工具类函数==================================================================

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'选择所有要素
Function SelAllFeatures(ByRef TotalCount,ByRef AllIdArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT,LINE,AREA,NOTE"
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount
    ReDim AllIdArr(TotalCount)
    For i = 0 To TotalCount - 1
        AllIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelAllFeatures

'分解字符串
Function SplitString(ByVal StringName,ByRef StrArr(),ByRef StrCount)
    StrArr = Split(StringName,",", - 1,1)
    StrCount = UBound(StrArr) + 1
    MsgBox StrCount
End Function' SplitString

'分解键字符串
Function SplitKeyString(ByVal StringName,ByRef StrArr(),ByRef StrCount)
    StrArr = Split(StringName,",", - 1,1)
    StrCount = UBound(StrArr) + 1
    For i = 0 To StrCount - 1
        StrArr(i) = "[" & StrArr(i) & "]"
    Next 'i
End Function' SplitKeyString

'获取当前系统时间
Function GetNowTime()
    GetNowTime = FormatDateTime(Now(),2)
End Function' GetNowTime

Function Ending()
    MsgBox "刷新完成"
End Function' Ending
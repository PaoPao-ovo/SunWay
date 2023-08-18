'======================================================检查集配置=====================================================

'检查集项目名称
Dim strGroupName
strGroupName = "验线检查"

'检查集组名称
Dim strCheckName
strCheckName = "实测点重复检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->实测点重复检查"

'检查描述
Dim strDescription
strDescription = "同名的理论点不存在"

'==================================================实测点编码配置=========================================================

'实测和预测对应关系：
' 实测编码            点名                理论编码
' 9130512           GPS检测点            1103021
' 9130412           水准点                1102021
' 9130311           控制点（埋石）         9130211
' 9130312           控制点（不埋石）         9130212
' 9130217           测站点                9130216
' 9130511           放样点               9130411


ScdCodes = "9130512,9130412,9130311,9130312,9130217,9130511"

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
Function ExportRecords(codes)
    ScArr = Split(ScdCodes,",", - 1,1)
    For i = 0 To UBound(ScArr)
        SelRealPoi ScArr(i)
        SelCount = SSProcess.GetSelGeoCount()
        ReDim Selids(SelCount,2)
        If SelCount > 0  Then
            For j = 0 To SelCount - 1
                Selids(j,0) = SSProcess.GetSelGeoValue(j,"SSObj_ID")
                Selids(j,1) = SSProcess.GetSelGeoValue(j,"SSObj_PointName")
            Next 'j
            For k = 0 To SelCount - 1
                Select Case ScArr(i)
                    Case "9130512"
                    SelLlPoi "1103021",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        'MsgBox geoType
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName,strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130412"
                    SelLlPoi "1102021",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130311"
                    SelLlPoi "9130211",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130312"
                    SelLlPoi "9130212",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130217"
                    SelLlPoi "9130216",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                    Case "9130511"
                    SelLlPoi "9130411",Selids(k,1)
                    Count = SSProcess.GetSelGeoCount()
                    If Count = 0 Then
                        geoType = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Type")
                        x = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_X")
                        y = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Y")
                        z = SSProcess.GetObjectAttr(Selids(k,0),"SSObj_Z")
                        SSProcess.AddCheckRecord strGroupName, strCheckName,CheckmodelName, strDescription, x, y, 0, 0, Selids(k,0), ""
                    End If
                End Select
            Next 'k
        End If
    Next 'i
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

'初始化
Sub OnInitScript()
    
    
    ClearSelection '清空原有系统选择集
    
    Mode = 0 '=0 无参数对话框 =1 有参数对话框

    Title = "生成线"
    
    SSProcess.ShowScriptDlg Mode,Title
    
End Sub

'点击关闭后执行
Sub OnExitScript()
    '添加代码
End Sub

'点击完成后执行
Sub OnOK()
    
    UpdateSelection SelCount,LineIdArr '更新选择集，返回选择个数和线ID数组
    
    For i = 0 To UBound(LineIdArr)
        
        GXQDDH = SSProcess.GetObjectAttr(LineIdArr(i),"[GXQDDH]")
        GXZDDH = SSProcess.GetObjectAttr(LineIdArr(i),"[GXZDDH]")
        
        GetPointXY GXQDDH,GXZDDH,StartX,StartY,EndX,EndY '获取同名的点的X和Y值
        
        SetLineXY LineIdArr(i),StartX,StartY,EndX,EndY '修改线位置
        
    Next 'i
    
    SSProcess.RefreshView()
    
End Sub


'取消后执行
Sub OnCancel()
    '添加代码
End Sub

'清空系统选择集
Function ClearSelection()
    SSProcess.ClearSysSelection
End Function' ClearSelection

'将系统选择集更新到脚本中
Function UpdateSelection(ByRef SelCount,ByRef LineIdArr)
    
    SSProcess.UpdateSysSelection 0 '选择集更新
    SelCount = SSProcess.GetSelGeoCount()
    ReDim LineIdArr(SelCount - 1)
    For i = 0 To SelCount - 1
        LineIdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
    
End Function' UpdateSelection

'获取点的坐标
Function GetPointXY(ByVal GXQDDH,ByVal GXZDDH,ByRef StartX,ByRef StartY,ByRef EndX,ByRef EndY)
    
    SqlStr = "Select 地下管线点属性表.ID From 地下管线点属性表 inner join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And WTDH = " & "'" & GXQDDH & "'"
    
    GetSQLRecordAll SqlStr,StartPointArr,Count
    
    If Count > 0 Then
        StartX = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_X")
        StartY = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_Y")
    End If
    
    SqlStr = "Select 地下管线点属性表.ID From 地下管线点属性表 inner join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And WTDH = " & "'" & GXZDDH & "'"
    
    GetSQLRecordAll SqlStr,StartPointArr,Count
    
    If Count > 0 Then
        EndX = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_X")
        EndY = SSProcess.GetObjectAttr(StartPointArr(0),"SSObj_Y")
    End If
    
End Function' GetPointXY

'修改线位置
Function SetLineXY(ByVal LineID,ByVal StartX,ByVal StartY,ByVal EndX,ByVal EndY)

    Pointcount = Transform(SSProcess.GetObjectAttr(LineID,"SSObj_PointCount"))
    If Pointcount = 2 Then
        SSProcess.GetObjectPoint LineID,0,StartLineX,StartLineY,StartLineZ,StartPointType,StartName
        SSProcess.GetObjectPoint LineID,1,EndLineX,EndLineY,EndLineZ,EndPointType,EndName
    End If
    
    If StartX <> "" And StartY <> "" Then
        SSProcess.SetObjectAttr LineID,"SSObj_X(0)",StartX
        SSProcess.SetObjectAttr LineID,"SSObj_Y(0)",StartY
    End If
    
    If EndX <> "" And EndY <> "" Then
        SSProcess.SetObjectAttr LineID,"SSObj_X(1)",EndX
        SSProcess.SetObjectAttr LineID,"SSObj_Y(1)",EndY
    End If

End Function' SetLineXY

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

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

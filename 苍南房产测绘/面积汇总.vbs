
'入口
Sub OnClick()
    
    SetArea
    
End Sub' OnClick

'=============================================业务函数============================================

'三类面积系数汇总
Function SetArea()
    
    '阳台面积
    Dim YT_Area
    
    '设备平台面积
    Dim SB_Area
    
    '飘窗面积
    Dim PC_Area
    
    '获取关联的【BZGUID】
    SqlStr = "Select FC_LPB_户信息表.BZGUID From FC_LPB_户信息表 Where FC_LPB_户信息表.ID > 0"
    GetSQLRecordAll SqlStr,BZGUIDArr,HCount
    
    For i = 0 To HCount - 1
        
        '阳台
        SqlStr = "Select Sum(FC_面积块信息属性表.KZMJ) From FC_面积块信息属性表 Inner Join GeoAreaTB On FC_面积块信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark Mod 2) <> 0 And FC_面积块信息属性表.MJXS = 0 And FC_面积块信息属性表.MJKMC Like '*阳台*' And FC_面积块信息属性表.BZGUID = '" & BZGUIDArr(i) & "'"
        GetSQLRecordAll SqlStr,YTArr,YTCount
        YT_Area = YTArr(0)
        If YT_Area <> "" Then
            SqlStr = "Update FC_LPB_户信息表 Set BFMYTTXMJ = " & YT_Area & " Where FC_LPB_户信息表.BZGUID = '" & BZGUIDArr(i) & "'"
            ProJectName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb ProJectName
            SSProcess.ExecuteAccessSql ProJectName,SqlStr
            SSProcess.CloseAccessMdb ProJectName
            
        End If
        
        
        '设备平台
        SqlStr = "Select Sum(FC_面积块信息属性表.KZMJ) From FC_面积块信息属性表 Inner Join GeoAreaTB On FC_面积块信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark Mod 2) <> 0 And FC_面积块信息属性表.MJXS = 0 And FC_面积块信息属性表.MJKMC Like '*设备平台*' And FC_面积块信息属性表.BZGUID = '" & BZGUIDArr(i) & "'"
        GetSQLRecordAll SqlStr,SBArr,SBCount
        SB_Area = SBArr(0)
        
        If SB_Area <> "" Then
            
            SqlStr = "Update FC_LPB_户信息表 Set SBPTTXMJ = " & SB_Area & " Where FC_LPB_户信息表.BZGUID = '" & BZGUIDArr(i) & "'"
            ProJectName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb ProJectName
            SSProcess.ExecuteAccessSql ProJectName,SqlStr
            SSProcess.CloseAccessMdb ProJectName
            
        End If
        
        
        '飘窗
        SqlStr = "Select Sum(FC_面积块信息属性表.KZMJ) From FC_面积块信息属性表 Inner Join GeoAreaTB On FC_面积块信息属性表.ID = GeoAreaTB.ID Where (GeoAreaTB.Mark Mod 2) <> 0 And FC_面积块信息属性表.MJXS = 0 And FC_面积块信息属性表.MJKMC Like '*飘窗*' And FC_面积块信息属性表.BZGUID = '" & BZGUIDArr(i) & "'"
        GetSQLRecordAll SqlStr,PCArr,PCCount
        PC_Area = PCArr(0)
        
        If PC_Area <> "" Then

            SqlStr = "Update FC_LPB_户信息表 Set PCTXMJ = " & PC_Area & " Where FC_LPB_户信息表.BZGUID = '" & BZGUIDArr(i) & "'"
            ProJectName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb ProJectName
            SSProcess.ExecuteAccessSql ProJectName,SqlStr
            SSProcess.CloseAccessMdb ProJectName

        End If
        
        
    Next 'i
End Function' SetArea





'==================================================工具类函数===================================================

'数值转换
Function Transform(ByVal Values)
    If Values <> "" Then
        If IsNumeric(Values) = True Then
            Values = CDbl(Values)
        End If
    Else
        Values = 0
    End If
    Transform = Values
End Function'Transform

'SQL查询，获取所有的记录
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
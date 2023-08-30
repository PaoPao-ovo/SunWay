

Sub OnClick()

    GetZT Ceng,H,ZRZ

End Sub

'获取测绘状态
Function GetZT(ByRef Ceng,ByRef H,ByRef ZRZ)
    
    SqlStr = "Select DISTINCT CHZT From FC_楼层信息属性表 Inner Join GeoAreaTB on FC_楼层信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,CCHZTArr,LXCount
    Ceng = CCHZTArr(0)

    SqlStr = "Select DISTINCT CHZT From FC_户信息属性表 Inner Join GeoAreaTB on FC_户信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,HCHZTArr,LXCount
    H = HCHZTArr(0)

    SqlStr = "Select DISTINCT CHZT From FC_自然幢信息属性表 Inner Join GeoAreaTB on FC_自然幢信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,ZRZCHZTArr,LXCount
    ZRZ = ZRZCHZTArr(0)

End Function ' GetZT

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
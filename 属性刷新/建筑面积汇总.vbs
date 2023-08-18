'人防防护范围数量
Dim RenfCount
'
Dim RenfValues

'入口函数
Sub Onclick()
    GetRenfCount()
    SetArea 9450033
    EmptyRFFHDYXXInfo()
    SQLRefresh()
End Sub ' Onclick

'获取人防范围面的数量
Function GetRenfCount()
    SSProcess.PushUndoMark
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_code", "==", 9450013
	SSProcess.SelectFilter
    RenfCount = SSProcess.GetSelGeoCount()
    RenfCount = transform(RenfCount)
End Function ' GetInnerIds

'属性表刷新掩蔽区面积
Function SetArea(Code)
    If RenfCount > 0 Then
        For i = 0 To RenfCount -1
            Dim temp:temp = 0.0
            ID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            'MsgBox ID
			InnerIds = SSProcess.SearchInnerObjIDs(ID, 2,Code, 1)
            Arr = Split(InnerIds,",",-1,1)
            For j = 0 To UBound(Arr)
                Area = SSProcess.GetObjectAttr(Arr(j), "SSObj_Area")
                Area = transform(Area)
                If temp = 0.0 Then
                    temp = Area
                Else
                    temp = temp + Area
                End If
            Next
            temp = formatnumber(temp,2)
            SetArea = temp
            SetArea = transform(SetArea)
            'MsgBox SetArea
			Featureid = GenNewGUID()
            SSProcess.SetObjectAttr ID,"[YBMJ]",SetArea
			SSProcess.SetObjectAttr ID,"[ID_FHDY]",Featureid
        Next
    End If
End Function ' SetArea

'二维表刷新
Function SQLRefresh()
    Fields = "MC,BH,ZSGN,PROTECTIONL,FHDJ,KBDYS,KBSL,SZCS,PSGN,TCWSL,FJDCSL,BZ,JZMJ,ID_ZRZ,ID_LJZ,ID_LC,ID_FHDY,YBMJ,GYMJ"
    sql = "select  " & Fields & "  from RF_人防防护单元范围属性表 inner join GeoAreaTB on RF_人防防护单元范围属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetInfo "RF_人防防护单元范围属性表",Fields,arSQLRecord,iRecordCount
    File = "MC,BH,ZSGN,PROTECTIONLEVEL,FHDJ,KBDYS,KBSL,SZCS,PSGN,TCWSL,FJDCSL,BZ,JZMJ,ID_ZRZ,ID_LJZ,ID_LC,ID_FHDY,YBMJ,GYMJ"
    SqlString = "select  " & File & "  from RFFHDYXX "
    'MsgBox RenfValues
    For i = 0 To iRecordCount -1
        InsertInfo SqlString,File,arSQLRecord(i)
    Next
End Function ' SQLRefresh

'==========================================工具类函数==================================================

'数据类型转换
Function transform(content)
	If content <> "" Then
		content = CDbl(content)
	Else 
		MsgBox "存在空值"
	End If
	transform = content
End Function

'获取最新的FeatureGUID
Function GenNewGUID()
	set TypeLib = CreateObject("Scriptlet.TypeLib")
	GenNewGUID = Left(TypeLib.Guid,38)
	set TypeLib=nothing
End Function

'清空面积表数据
Function EmptyRFFHDYXXInfo()
	sql = "SELECT * FROM RFFHDYXX " & ";"
	mdbName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql 
	
	while  SSProcess.AccessIsEOF (mdbName, sql)=false
			SSProcess.DelAccessRecord mdbName, sql 
	wend 
		SSProcess.CloseAccessRecordset mdbName, sql '关库
		SSProcess.CloseAccessMdb mdbName
End Function

'获取表信息
Function GetInfo(Tablename,fields, ByRef arSQLRecord(), ByRef iRecordCount)
	MdbName=SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb MdbName 
	SqlString ="SELECT  " & fields & "  From  " & Tablename & "  inner join GeoAreaTB on RF_人防防护单元范围属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
	GetSQLRecordAll MdbName,SqlString,arSQLRecord,iRecordCount
    SSProcess.CloseAccessMdb MdbName
End Function 

'插入信息
Function InsertInfo(sql,Infile,Values)
	mdbName = SSProcess.GetProjectFileName 
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql
	recordc=SSProcess.GetAccessRecordCount(mdbName, sql)
	SSProcess.AddAccessRecord mdbName,sql,Infile,Values
	SSProcess.CloseAccessRecordset mdbName, sql
	SSProcess.CloseAccessMdb mdbName
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
	if StrSqlStatement ="" then
		msgbox "查询语句为空，操作停止！",48
	end if
	iRecordCount = -1
	'SQL语句
	sql =StrSqlStatement
	'打开记录集
	SSProcess.OpenAccessRecordset mdbName, sql
	'获取记录总数
	RecordCount =SSProcess.GetAccessRecordCount (mdbName, sql)
	if RecordCount >0 then
		iRecordCount =0
		ReDim arSQLRecord(RecordCount)
		'将记录游标移到第一行
		SSProcess.AccessMoveFirst mdbName, sql
		iRecordCount = 0
		'浏览记录
		While SSProcess.AccessIsEOF (mdbName, sql) = 0
			fields = ""
			values = ""
			'获取当前记录内容
			SSProcess.GetAccessRecord mdbName, sql, fields, values
			arSQLRecord(iRecordCount) =values
			iRecordCount =iRecordCount +1		
			'移动记录游标
			SSProcess.AccessMoveNext mdbName, sql
		Wend
	end if
	'关闭记录集
	SSProcess.CloseAccessRecordset mdbName, sql
End FUnction

'Excel变量
	Dim xlApp,xlFile,xlsheet,XKZValues,XKZZD,JZCount
	
	'用地红线GUID 和 宗地代码
	Dim YDHXGUID
    ZDCode="9410001"
	
	'幢的ID
	Dim DTid(10000,2)
	
	Dim arrkey(10000000)
Sub OnClick()

	'选取宗地
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	'SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=", ZDCode
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount
	If geoCount=0 Then  
		GetZDID=0 : Exit Sub
	ElseIf geoCount=1 Then
					ZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")	
					YDHXGUID=SSProcess.GetSelGeoValue (0, "[YDHXGUID]")
					If YDHXGUID="{00000000-0000-0000-0000-000000000000}"  Then  
					YDHXGUID =  GenNewGUID
					SSProcess.SetObjectAttr ZDID, "[YDHXGUID]", YDHXGUID 
					End If 
	Else
		Msgbox "图上有多个地!" : Exit Sub
	End If
	aa=MsgBox("将覆盖已有数据，是否导入信息？",4+64)'是6 否7
		if aa=7 then  exit sub
	'打开Excel表格
	ExcelFile=SSProcess.SelectFileName(1,"选择excel文件",0,"EXCEL Files(*.xlsx)|*.xlsx|EXCEL Files(*.xls)|*.xls|All Files (*.*)|*.*||")
	If ExcelFile="" Then Exit Sub
	Set xlApp=CreateObject("Excel.Application")
	Set xlFile=xlApp.Workbooks.Open(ExcelFile)
	
	'用地红线Excel信息调入
	SSProcess.MapMethod "clearattrbuffer", "JG_用地红线信息属性表"
	
	DT()
	JMZB()
	xlApp.quit
End Sub

	'属性表名称
	Table_JZDT = "JG_建设工程建筑单体信息属性表" 
	Table_DTMJ = "JG_建筑物单体建筑面积指标核实信息属性表" 
	Table_YDHX = "JG_用地红线信息属性表"
	'单体基本信息
	Function DT()
		EmptyDTInfo()
		Set xlsheet = xlFile.Worksheets("单体基本信息")
		xlsheet.Activate
		dtxx = ""
		excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
		'MsgBox excelhs
		For i =2 To excelhs 
			If xlApp.Cells(i,1) <> "" Then 
				For k = 1 To 8
					str = xlApp.Cells(i,k)
					If dtxx = ""  Then 
					dtxx = str
					ElseIf k=1 Then
					dtxx = dtxx & str
					ElseIf k = 8 Then
					dtxx = dtxx &","& str &";"
					Else
					dtxx = dtxx &","& str 	
					End If
				Next
			End If
		Next 
		Infile = "FeatureGUID,YDHXGUID,JZWMCGUID,JianZWMC,SNDPBG,SWDPBG,DCNDPBG,JZZGDBG,GuiHSPDSCS,GuiHSPDXCS,GuiHSPJDMJ"
		Sql= "select "&Infile&" from "&Table_JZDT
		Dim arr(100000),Count,Info
		SSFunc.ScanString dtxx, ";", arr, Count
		JZCount = Count-2
		For y = 0 To Count-2
		FeatureGUID = GenNewGUID
		JZWMCGUID = GenNewGUID
		DTid(y,1) = JZWMCGUID
		DTid(y,0) = xlApp.Cells(y+2,1)
		Values = FeatureGUID & "," & YDHXGUID & "," & JZWMCGUID & "," & arr(y)
		InsertInfo Sql,Infile,Values
		Next
			GetInfo(Table_YDHX)
			'Dim arrval(1000),vCount
			'SSFunc.ScanString XKZValues, "," , arrval, vCount 
			'For i =0 To vCount-1
				SqlString ="SELECT " & "GuiHYDXKZBH,GuiHXKZBH" & " From " & Table_JZDT
				'MsgBox SqlString
				'Val = arrval(i)
				'InsertInfo SqlString,"GuiHYDXKZBH,GuiHXKZBH",arrkey(i) & "," & Val
				'MsgBox Val
				'msgbox XKZValues
				inAttr SqlString,"GuiHYDXKZBH,GuiHXKZBH",XKZValues
				'InsertInfo SqlString,"GuiHYDXKZBH,GuiHXKZBH",XKZValues
			'Next 
	End Function
	
'面信息添加
Function JMZB()
	EmptyMJInfo()
	Set xlsheet = xlFile.Worksheets("单体建面指标")
	xlsheet.Activate
	jmxx = ""
	excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
	For i =2 To excelhs 
		If xlApp.Cells(i,1) <> "" Then 
			For k = 1 To 4
				str = xlApp.Cells(i,k)
				If jmxx = ""  Then 
				jmxx = "'" & str & "'"
				ElseIf k = 1 Then
				jmxx = jmxx & "'" & str & "'"
				ElseIf k = 2 Or k = 3 Then
				jmxx = jmxx & "," & "'" & str & "'"
				ElseIf k = 4 Then
				jmxx = jmxx &","& str &";"
				Else
				jmxx = jmxx &","& str 	
				End If
			Next
		End If
	Next 
		
	'sInfile = "FeatureGUID,JianZWMC,GongNLX,GuiHSPJZMJ,GuiHSPDSJZMJ"
	'xInfile = "FeatureGUID,JianZWMC,GongNLX,GuiHSPJZMJ,GuiHSPDXJZMJ"
	'bhfile = "YDHXGUID,JSGHXKZGUID,JZWMCGUID"
		
	sInfile = "FeatureGUID,YDHXGUID,JZWMCGUID,JianZWMC,GongNLX,GongNMC,GuiHSPJZMJ,WZ"
	xInfile = "FeatureGUID,YDHXGUID,JZWMCGUID,JianZWMC,GongNLX,GongNMC,GuiHSPJZMJ,WZ"
		
	'Ssql= "select "&sInfile&" from "&Table_DTMJ&" where ID>0"
	'Xsql ="select "&xInfile&" from "&Table_DTMJ&" where ID>0"
	Dim arr(100000),Count,Info
		SSFunc.ScanString jmxx, ";", arr, Count
	For y = 0 To Count-2
		'MsgBox Count-2
			'JZWMCGUID = DTid(y,1)
			'MsgBox JZWMCGUID
			For j = 0 To JZCount
				If xlApp.Cells(y+2,1) = DTid(j,0) Then
					JZWMCGUID = DTid(j,1)
				End If 
			Next
			If xlApp.Cells(y+2,5) ="地上" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID  & "," & JZWMCGUID & "," &arr(y)& "," & "'"&xlApp.Cells(y+2,5)&"'"
				'MsgBox Values
				InsertRecord  Table_DTMJ,sInfile,Values
			ElseIf xlApp.Cells(y+2,5) ="地下" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID & "," & JZWMCGUID & "," &arr(y)& "," & "'"&xlApp.Cells(y+2,5)&"'"
				InsertRecord  Table_DTMJ,xInfile,Values
			End If
	Next
	GetInfo(Table_YDHX)
			'Dim arrval(1000),vCount
			'SSFunc.ScanString XKZValues, "," , arrval, vCount 
			'For i =0 To vCount-1
				'SqlString ="SELECT " & "GuiHYDXKZBH,GuiHXKZBH" & " From " & Table_DTMJ
				'Val = arrval(i)
				'MsgBox Val
				'InsertInfo SqlString,"GuiHYDXKZBH,GuiHXKZBH",arrkey(i) & "," & Val
			'Next
	SqlString ="SELECT " & "GuiHYDXKZBH,GuiHXKZBH" & " From " & Table_DTMJ
	inAttr SqlString,"GuiHYDXKZBH,GuiHXKZBH",XKZValues
	'InsertInfo SqlString,"GuiHYDXKZBH,GuiHXKZBH",XKZValues
End Function
	
'修改表信息
Function inAttr(sql,infile,invalues)
	ProjectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb ProjectName
	SSProcess.OpenAccessRecordset ProjectName, sql
	rscount = SSProcess.GetAccessRecordCount (ProjectName, sql)
	If rscount > 0 Then
		SSProcess.AccessMoveFirst ProjectName, sql
		While (SSProcess.AccessIsEOF (ProjectName, sql ) = False)
			SSProcess.ModifyAccessRecord  ProjectName, sql, infile , invalues'输出到mdb表中
			SSProcess.AccessMoveNext ProjectName, sql 
		Wend
	End If
	SSProcess.CloseAccessRecordset ProjectName, sql 
	SSProcess.CloseAccessMdb ProjectName
End Function
	
'获取最新的FeatureGUID
Function GenNewGUID()
	set TypeLib = CreateObject("Scriptlet.TypeLib")
	GenNewGUID = Left(TypeLib.Guid,38)
	set TypeLib=nothing
End Function
	
'********插入新纪录
Function InsertRecord( tableName, fields, values)
		sqlString = "insert into " & tableName & " (" & fields &  ") values (" & values & ")"
		'MsgBox sqlString
		InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function

Function InsertInfo(sql,Infile,Values)
	mdbName = SSProcess.GetProjectFileName 
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql
	recordc=SSProcess.GetAccessRecordCount(mdbName, sql)
	SSProcess.AddAccessRecord mdbName,sql,Infile,Values
	SSProcess.CloseAccessRecordset mdbName, sql
	SSProcess.CloseAccessMdb mdbName
End Function

'清空面积表数据
Function EmptyMJInfo()
	sql = "SELECT * FROM JG_建筑物单体建筑面积指标核实信息属性表 where JG_建筑物单体建筑面积指标核实信息属性表.ID > " &"0" &";"
	mdbName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql 
	
	while  SSProcess.AccessIsEOF (mdbName, sql)=false
			SSProcess.DelAccessRecord mdbName, sql 
	wend 
		SSProcess.CloseAccessRecordset mdbName, sql '关库
		SSProcess.CloseAccessMdb mdbName
End Function

'清空单体表信息
Function EmptyDTInfo()
	sql ="SELECT * FROM JG_建设工程建筑单体信息属性表 where JG_建设工程建筑单体信息属性表.ID > " &"0" &";"
	mdbName = SSProcess.GetProjectFileName  

	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql  '打开数据库

	while  SSProcess.AccessIsEOF (mdbName, sql)=false
		SSProcess.DelAccessRecord mdbName, sql 
	wend 
	
	SSProcess.CloseAccessRecordset mdbName, sql '关库
	SSProcess.CloseAccessMdb mdbName
End Function

Function GetInfo(Tablename)
	MdbName=SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb MdbName 
	SqlString ="SELECT GuiHYDXKZBH,GuiHXKZBH FROM " & Tablename & " WHERE " & Tablename & "." & "YDHXGUID = " & YDHXGUID
	SSProcess.OpenAccessRecordset MdbName, SqlString
	SSProcess.GetAccessRecord MdbName,SqlString,Fields,IdValues
	SSProcess.CloseAccessRecordset MdbName, SqlString
	SSProcess.CloseAccessMdb MdbName
	
	XKZValues = IdValues
	'MsgBox XKZValues
	XKZZD = Fields
End Function 
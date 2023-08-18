
'Excel变量
	Dim xlApp,xlFile,xlsheet
	
	'用地红线GUID 和 宗地代码
	Dim YDHXGUID
    ZDCode="9410001"
	'建设工程规划许可证GUID
	Dim JSGHXKZGUID 
	
	'幢的ID
	Dim DTid(10000000)
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
	'建设工程规划许可证GUID
	JSGHXKZGUID = GenNewGUID
	'用地红线Excel信息调入
	SSProcess.MapMethod "clearattrbuffer", "JG_用地红线信息属性表"
	YDHX()
	XKZ(JSGHXKZGUID)
	DT(JSGHXKZGUID)
	JMZB()
	xlApp.quit
End Sub

	'属性表名称
	Table_YDHX = "JG_用地红线信息属性表" 
	Table_GHXKZ = "JG_建设工程规划许可证信息属性表" 
	Table_JZDT = "JG_建设工程建筑单体信息属性表" 
	Table_DTMJ = "JG_建筑物单体建筑面积指标核实信息属性表" 
	
	'用地红线信息导入
	Function YDHX()
		'选择Sheet
		Set xlsheet = xlFile.Worksheets("项目信息")
		xlsheet.Activate
		'获取Excel中的数据
		xmxx=""
		For i=1 To 18
			Redim arr(1000)
			ikey = xlApp.Cells(i,1)
			str =xlApp.Cells(i,2)
			keys = "装配式建筑面积,实测住宅户数,规划住宅户数"
			SSFunc.ScanString keys, ",", arr,IdCount
			If ikey <> arr(0) And ikey <> arr(1) And ikey <> arr(2) Then 
				If xmxx="" And i=1 Then 
				xmxx=str
				Else
				xmxx=xmxx&","&str
				End If
			End If
		Next
		
		Infile = "XiangMBH,GuiHYDXKZBH,XiangMMC,XiangMDD,CeHDWDZ,WeiTDW,JianSDW,YongDMJ,GuiHSPZJZMJ,GuiHSPDSJZMJ,GuiHSPDXJZMJ,GuiHSPJDMJ,GuiHSPRJL,GuiHSPJZMD,GuiHSPLHL"
		
		Sql = "Select "&infile&" From "&Table_YDHX&" Where "&Table_YDHX&".YDHXGUID ="&YDHXGUID
		inAttr Sql,Infile,xmxx
		
	End Function
	
	'工程信息录入
	Function XKZ(xkzid)
	EmptyHXZInfo
	'选择Sheet
		Set xlsheet = xlFile.Worksheets("项目信息")
		'获取Excel中的数据
		xkzxx=""
		For i=1 To 18
			Redim arr(1000)
			ikey = xlApp.Cells(i,1)
			str =xlApp.Cells(i,2)
			keys = "总用地面积,测绘单位地址,委托单位"
			SSFunc.ScanString keys, ",", arr,IdCount
			If ikey <> arr(0) And ikey <> arr(1) And ikey <> arr(2) Then 
				If xkzxx="" And i=1 Then 
				xkzxx=str
				Else
				xkzxx=xkzxx&","&str
				End If
			End If
		Next
		
		FeatureGUID = GenNewGUID
		
		Infile = "FeatureGUID,YDHXGUID,JSGHXKZGUID,XiangMBH,GuiHYDXKZBH,XiangMMC,XiangMDD,JianSDW,GuiHSPZJZMJ,GuiHSPDSJZMJ,GuiHSPDXJZMJ,ZpsJZMJ,GuiHSPJDMJ,GuiHSPRJL,GuiHSPJZMD,GuiHSPLHL,ScZZHS,GhZZHS"
		Values = FeatureGUID & "," & YDHXGUID & "," & xkzid & "," & xkzxx
		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb mdbName
		sql= "select "&Infile&" from "&Table_GHXKZ&" where ID>0"
		SSProcess.OpenAccessRecordset mdbName, sql
		recordc=SSProcess.GetAccessRecordCount(mdbName, sql)
		SSProcess.AddAccessRecord mdbName,sql,Infile,Values
		SSProcess.CloseAccessRecordset mdbName, sql
		SSProcess.CloseAccessMdb mdbName
		
	End Function
	'单体基本信息
	Function DT(xkzid)
		EmptyDTInfo()
		Set xlsheet = xlFile.Worksheets("单体基本信息")
		xlsheet.Activate
		dtxx = ""
		j = 3
		For i =2 To j 
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
			j=j+1
		Next 
		Infile = "FeatureGUID,YDHXGUID,JSGHXKZGUID,JZWMCGUID,JianZWMC,SNDPBG,SWDPBG,DCNDPBG,JZZGDBG,GuiHSPDSCS,GuiHSPDXCS,GuiHSPJDMJ"
		Sql= "select "&Infile&" from "&Table_JZDT&" where ID>0"
		JSGHXKZGUID = xkzid
		Dim arr(100000),Count,Info
		SSFunc.ScanString dtxx, ";", arr, Count
		For y = 0 To Count-2
		FeatureGUID = GenNewGUID
		JZWMCGUID = GenNewGUID
		DTid(y) = JZWMCGUID
		Values = FeatureGUID & "," & YDHXGUID & "," & JSGHXKZGUID & "," & JZWMCGUID & "," & arr(y)
		InsertInfo Sql,Infile,Values
		Next
		
	End Function
	
'面信息添加
Function JMZB()
	EmptyMJInfo()
	Set xlsheet = xlFile.Worksheets("单体建面指标")
	xlsheet.Activate
	jmxx = ""
	j = 3
	For i =2 To j 
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
		j=j+1
	Next 
		
	'sInfile = "FeatureGUID,JianZWMC,GongNLX,GuiHSPJZMJ,GuiHSPDSJZMJ"
	'xInfile = "FeatureGUID,JianZWMC,GongNLX,GuiHSPJZMJ,GuiHSPDXJZMJ"
	'bhfile = "YDHXGUID,JSGHXKZGUID,JZWMCGUID"
		
	sInfile = "FeatureGUID,YDHXGUID,JSGHXKZGUID,JZWMCGUID,JianZWMC,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDSJZMJ"
	xInfile = "FeatureGUID,YDHXGUID,JSGHXKZGUID,JZWMCGUID,JianZWMC,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDXJZMJ"
		
	'Ssql= "select "&sInfile&" from "&Table_DTMJ&" where ID>0"
	'Xsql ="select "&xInfile&" from "&Table_DTMJ&" where ID>0"
	Dim arr(100000),Count,Info
		SSFunc.ScanString jmxx, ";", arr, Count
	For y = 0 To Count-2
		JZWMCGUID = DTid(y)
			If xlApp.Cells(y+2,5) ="地上" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID & "," & JSGHXKZGUID & "," & JZWMCGUID & "," &arr(y)& "," & xlApp.Cells(y+2,4)
				InsertRecord  Table_DTMJ,sInfile,Values
			ElseIf xlApp.Cells(y+2,5) ="地下" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID & "," & JSGHXKZGUID & "," & JZWMCGUID & "," &arr(y)& "," & xlApp.Cells(y+2,4)
				InsertRecord  Table_DTMJ,xInfile,Values
			End If
	Next
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
		InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function


		
'清空许可证表信息
Function EmptyHXZInfo()
	sql ="SELECT * FROM JG_建设工程规划许可证信息属性表 where JG_建设工程规划许可证信息属性表.ID > " &"0" &";"
	mdbName = SSProcess.GetProjectFileName  

	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql  

	while  SSProcess.AccessIsEOF (mdbName, sql)=false
		SSProcess.DelAccessRecord mdbName, sql 
	wend 
	SSProcess.CloseAccessRecordset mdbName, sql 
	SSProcess.CloseAccessMdb mdbName
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

'Excel����
	Dim xlApp,xlFile,xlsheet,XKZValues,XKZZD,JZCount
	
	'�õغ���GUID �� �ڵش���
	Dim YDHXGUID
    ZDCode="9410001"
	
	'����ID
	Dim DTid(10000,2)
	
	Dim arrkey(10000000)
Sub OnClick()

	'ѡȡ�ڵ�
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
		Msgbox "ͼ���ж����!" : Exit Sub
	End If
	aa=MsgBox("�������������ݣ��Ƿ�����Ϣ��",4+64)'��6 ��7
		if aa=7 then  exit sub
	'��Excel���
	ExcelFile=SSProcess.SelectFileName(1,"ѡ��excel�ļ�",0,"EXCEL Files(*.xlsx)|*.xlsx|EXCEL Files(*.xls)|*.xls|All Files (*.*)|*.*||")
	If ExcelFile="" Then Exit Sub
	Set xlApp=CreateObject("Excel.Application")
	Set xlFile=xlApp.Workbooks.Open(ExcelFile)
	
	'�õغ���Excel��Ϣ����
	SSProcess.MapMethod "clearattrbuffer", "JG_�õغ�����Ϣ���Ա�"
	
	DT()
	JMZB()
	xlApp.quit
End Sub

	'���Ա�����
	Table_JZDT = "JG_���蹤�̽���������Ϣ���Ա�" 
	Table_DTMJ = "JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�" 
	Table_YDHX = "JG_�õغ�����Ϣ���Ա�"
	'���������Ϣ
	Function DT()
		EmptyDTInfo()
		Set xlsheet = xlFile.Worksheets("���������Ϣ")
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
	
'����Ϣ���
Function JMZB()
	EmptyMJInfo()
	Set xlsheet = xlFile.Worksheets("���彨��ָ��")
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
			If xlApp.Cells(y+2,5) ="����" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID  & "," & JZWMCGUID & "," &arr(y)& "," & "'"&xlApp.Cells(y+2,5)&"'"
				'MsgBox Values
				InsertRecord  Table_DTMJ,sInfile,Values
			ElseIf xlApp.Cells(y+2,5) ="����" Then
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
	
'�޸ı���Ϣ
Function inAttr(sql,infile,invalues)
	ProjectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb ProjectName
	SSProcess.OpenAccessRecordset ProjectName, sql
	rscount = SSProcess.GetAccessRecordCount (ProjectName, sql)
	If rscount > 0 Then
		SSProcess.AccessMoveFirst ProjectName, sql
		While (SSProcess.AccessIsEOF (ProjectName, sql ) = False)
			SSProcess.ModifyAccessRecord  ProjectName, sql, infile , invalues'�����mdb����
			SSProcess.AccessMoveNext ProjectName, sql 
		Wend
	End If
	SSProcess.CloseAccessRecordset ProjectName, sql 
	SSProcess.CloseAccessMdb ProjectName
End Function
	
'��ȡ���µ�FeatureGUID
Function GenNewGUID()
	set TypeLib = CreateObject("Scriptlet.TypeLib")
	GenNewGUID = Left(TypeLib.Guid,38)
	set TypeLib=nothing
End Function
	
'********�����¼�¼
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

'������������
Function EmptyMJInfo()
	sql = "SELECT * FROM JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա� where JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.ID > " &"0" &";"
	mdbName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql 
	
	while  SSProcess.AccessIsEOF (mdbName, sql)=false
			SSProcess.DelAccessRecord mdbName, sql 
	wend 
		SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
		SSProcess.CloseAccessMdb mdbName
End Function

'��յ������Ϣ
Function EmptyDTInfo()
	sql ="SELECT * FROM JG_���蹤�̽���������Ϣ���Ա� where JG_���蹤�̽���������Ϣ���Ա�.ID > " &"0" &";"
	mdbName = SSProcess.GetProjectFileName  

	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql  '�����ݿ�

	while  SSProcess.AccessIsEOF (mdbName, sql)=false
		SSProcess.DelAccessRecord mdbName, sql 
	wend 
	
	SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
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

'Excel����
	Dim xlApp,xlFile,xlsheet
	
	'�õغ���GUID �� �ڵش���
	Dim YDHXGUID
    ZDCode="9410001"
	'���蹤�̹滮���֤GUID
	Dim JSGHXKZGUID 
	
	'����ID
	Dim DTid(10000000)
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
	'���蹤�̹滮���֤GUID
	JSGHXKZGUID = GenNewGUID
	'�õغ���Excel��Ϣ����
	SSProcess.MapMethod "clearattrbuffer", "JG_�õغ�����Ϣ���Ա�"
	YDHX()
	XKZ(JSGHXKZGUID)
	DT(JSGHXKZGUID)
	JMZB()
	xlApp.quit
End Sub

	'���Ա�����
	Table_YDHX = "JG_�õغ�����Ϣ���Ա�" 
	Table_GHXKZ = "JG_���蹤�̹滮���֤��Ϣ���Ա�" 
	Table_JZDT = "JG_���蹤�̽���������Ϣ���Ա�" 
	Table_DTMJ = "JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�" 
	
	'�õغ�����Ϣ����
	Function YDHX()
		'ѡ��Sheet
		Set xlsheet = xlFile.Worksheets("��Ŀ��Ϣ")
		xlsheet.Activate
		'��ȡExcel�е�����
		xmxx=""
		For i=1 To 18
			Redim arr(1000)
			ikey = xlApp.Cells(i,1)
			str =xlApp.Cells(i,2)
			keys = "װ��ʽ�������,ʵ��סլ����,�滮סլ����"
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
	
	'������Ϣ¼��
	Function XKZ(xkzid)
	EmptyHXZInfo
	'ѡ��Sheet
		Set xlsheet = xlFile.Worksheets("��Ŀ��Ϣ")
		'��ȡExcel�е�����
		xkzxx=""
		For i=1 To 18
			Redim arr(1000)
			ikey = xlApp.Cells(i,1)
			str =xlApp.Cells(i,2)
			keys = "���õ����,��浥λ��ַ,ί�е�λ"
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
	'���������Ϣ
	Function DT(xkzid)
		EmptyDTInfo()
		Set xlsheet = xlFile.Worksheets("���������Ϣ")
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
	
'����Ϣ���
Function JMZB()
	EmptyMJInfo()
	Set xlsheet = xlFile.Worksheets("���彨��ָ��")
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
			If xlApp.Cells(y+2,5) ="����" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID & "," & JSGHXKZGUID & "," & JZWMCGUID & "," &arr(y)& "," & xlApp.Cells(y+2,4)
				InsertRecord  Table_DTMJ,sInfile,Values
			ElseIf xlApp.Cells(y+2,5) ="����" Then
				FeatureGUID = GenNewGUID
				Values = FeatureGUID  & "," & YDHXGUID & "," & JSGHXKZGUID & "," & JZWMCGUID & "," &arr(y)& "," & xlApp.Cells(y+2,4)
				InsertRecord  Table_DTMJ,xInfile,Values
			End If
	Next
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
		InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function


		
'������֤����Ϣ
Function EmptyHXZInfo()
	sql ="SELECT * FROM JG_���蹤�̹滮���֤��Ϣ���Ա� where JG_���蹤�̹滮���֤��Ϣ���Ա�.ID > " &"0" &";"
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
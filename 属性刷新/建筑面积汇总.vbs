'�˷�������Χ����
Dim RenfCount
'
Dim RenfValues

'��ں���
Sub Onclick()
    GetRenfCount()
    SetArea 9450033
    EmptyRFFHDYXXInfo()
    SQLRefresh()
End Sub ' Onclick

'��ȡ�˷���Χ�������
Function GetRenfCount()
    SSProcess.PushUndoMark
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_code", "==", 9450013
	SSProcess.SelectFilter
    RenfCount = SSProcess.GetSelGeoCount()
    RenfCount = transform(RenfCount)
End Function ' GetInnerIds

'���Ա�ˢ���ڱ������
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

'��ά��ˢ��
Function SQLRefresh()
    Fields = "MC,BH,ZSGN,PROTECTIONL,FHDJ,KBDYS,KBSL,SZCS,PSGN,TCWSL,FJDCSL,BZ,JZMJ,ID_ZRZ,ID_LJZ,ID_LC,ID_FHDY,YBMJ,GYMJ"
    sql = "select  " & Fields & "  from RF_�˷�������Ԫ��Χ���Ա� inner join GeoAreaTB on RF_�˷�������Ԫ��Χ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
    GetInfo "RF_�˷�������Ԫ��Χ���Ա�",Fields,arSQLRecord,iRecordCount
    File = "MC,BH,ZSGN,PROTECTIONLEVEL,FHDJ,KBDYS,KBSL,SZCS,PSGN,TCWSL,FJDCSL,BZ,JZMJ,ID_ZRZ,ID_LJZ,ID_LC,ID_FHDY,YBMJ,GYMJ"
    SqlString = "select  " & File & "  from RFFHDYXX "
    'MsgBox RenfValues
    For i = 0 To iRecordCount -1
        InsertInfo SqlString,File,arSQLRecord(i)
    Next
End Function ' SQLRefresh

'==========================================�����ຯ��==================================================

'��������ת��
Function transform(content)
	If content <> "" Then
		content = CDbl(content)
	Else 
		MsgBox "���ڿ�ֵ"
	End If
	transform = content
End Function

'��ȡ���µ�FeatureGUID
Function GenNewGUID()
	set TypeLib = CreateObject("Scriptlet.TypeLib")
	GenNewGUID = Left(TypeLib.Guid,38)
	set TypeLib=nothing
End Function

'������������
Function EmptyRFFHDYXXInfo()
	sql = "SELECT * FROM RFFHDYXX " & ";"
	mdbName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql 
	
	while  SSProcess.AccessIsEOF (mdbName, sql)=false
			SSProcess.DelAccessRecord mdbName, sql 
	wend 
		SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
		SSProcess.CloseAccessMdb mdbName
End Function

'��ȡ����Ϣ
Function GetInfo(Tablename,fields, ByRef arSQLRecord(), ByRef iRecordCount)
	MdbName=SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb MdbName 
	SqlString ="SELECT  " & fields & "  From  " & Tablename & "  inner join GeoAreaTB on RF_�˷�������Ԫ��Χ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0"
	GetSQLRecordAll MdbName,SqlString,arSQLRecord,iRecordCount
    SSProcess.CloseAccessMdb MdbName
End Function 

'������Ϣ
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
		msgbox "��ѯ���Ϊ�գ�����ֹͣ��",48
	end if
	iRecordCount = -1
	'SQL���
	sql =StrSqlStatement
	'�򿪼�¼��
	SSProcess.OpenAccessRecordset mdbName, sql
	'��ȡ��¼����
	RecordCount =SSProcess.GetAccessRecordCount (mdbName, sql)
	if RecordCount >0 then
		iRecordCount =0
		ReDim arSQLRecord(RecordCount)
		'����¼�α��Ƶ���һ��
		SSProcess.AccessMoveFirst mdbName, sql
		iRecordCount = 0
		'�����¼
		While SSProcess.AccessIsEOF (mdbName, sql) = 0
			fields = ""
			values = ""
			'��ȡ��ǰ��¼����
			SSProcess.GetAccessRecord mdbName, sql, fields, values
			arSQLRecord(iRecordCount) =values
			iRecordCount =iRecordCount +1		
			'�ƶ���¼�α�
			SSProcess.AccessMoveNext mdbName, sql
		Wend
	end if
	'�رռ�¼��
	SSProcess.CloseAccessRecordset mdbName, sql
End FUnction
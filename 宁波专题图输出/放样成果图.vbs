rem autor: <wsw> 
rem email: XXXX@xxx.com 
rem �ű��ļ���: C:\Users\wsw\Desktop\����\����\1111.vbs
rem ��Ӧ�����ļ���:F:\����΢\2023����\Z�㽭\N����\����\EPS����һ����\DeskTop\����һ\����ģ��\�����ɹ�ͼ���.Map
rem ��������:�����ɹ�ͼ
rem ���ű��ļ�Ӧ������ EPS��װĿ¼\desktop\XX̨��\Script\����\�����ɹ�ͼ.vbs
rem framework: gq 
rem framework: 471b1e20fe69040339fca38c3d3a189b 



rem special:[�����ɹ�ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap0(MSGID,mapName,selectID)
	  
	 rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 rem return = 1 ֹͣ����ɹ�ͼ
	 rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		TKFZ1 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... ,������ͼ��·��ÿ�λ���ýű����ص�·�������̲���ͬ����ͨ�������÷�Χ�ߵ������չ����ƴ��
        strProjectName=SSProcess.GetProjectFileName()
        FileFolder=replace(strProjectName,".edb","")
			'FileFolder = SSProcess.GetSysPathName (5) 
			CreateFolders FileFolder
			SaveFile = FileFolder&"\���蹤��ʵ�ط���ƽ��ͼ.edb"
			SSParameter.SetParameterSTR "printMap","NewedbName",SaveFile
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 ElseIf MSGID = 2 Then '// �¹����Զ���Ŀ¼��ͼ(����ѡ�񱣴�·��) 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 
	  
	  
	  
rem special:[�����ɹ�ͼ] ��ͼ����ɴ˽���
Function VBS_postMap0(MSGID,mapName,selectID)
	  
	 rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
	
	 '// ������ĳɹ�ͼ������� 
	TKFZ2 
	fxtl
	 rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 
		 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

	
function TKFZ1()
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	

		SSProcess.ObjectDeal id, "GotoPoints", "", result

		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb  mdbName
		sql = "select VALUE from PROJECTINFO where KEY='��浥λ'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount > 0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC = ""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='������ʼʱ��'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount

		If nSeletionCount > 0 Then
			HTRY=FormatDateTime(arSeletionRecord(0),1)
		Else
			HTRY=""
		End If

		
		sql = "select VALUE from PROJECTINFO where KEY='������Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount

		If nSeletionCount > 0 Then
			JCRY=arSeletionRecord(0)
		Else
			JCRY=""
		End If
		
		strtemp = XMMC&","& HTRY &","&JCRY

		SSProcess.CloseAccessMdb mdbName 

		SSProcess.SetObjectAttr id,"SSObj_DataMark",strtemp
	next

end function

function TKFZ2( )
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9310093
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	
 		ids = SSProcess.SearchInnerObjIDs(id,1,"9410001",0)
		idsList=split(ids,",")
		strtemp = SSProcess.GetObjectAttr (idsList(0),"SSObj_DataMark")
		artemp = split(strtemp,",")
		SSProcess.SetObjectAttr id, "[���ߵ�λ]", artemp(0)
		SSProcess.SetObjectAttr id, "[��������]", artemp(1)
		SSProcess.SetObjectAttr id, "[��ͼԱ]", artemp(2)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'ͼ����������
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next
	SSProcess.DeleteLayer "TKZSM"	
end function

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
			arSQLRecord(iRecordCount) =values										'��ѯ��¼
			iRecordCount =iRecordCount +1													'��ѯ��¼��
			'�ƶ���¼�α�
			SSProcess.AccessMoveNext mdbName, sql
		Wend
	end if
	'�رռ�¼��
	SSProcess.CloseAccessRecordset mdbName, sql
End FUnction
	  
	  
	  
Dim g_MapList,g_MapPrePtrfun,g_MapPostPtrfun 
rem �����������޸�
Sub OnClick() 
	 
	rem ��ʼ�� 
	 g_MapList = Array("�����ɹ�ͼ")
	 g_MapPrePtrfun = Array("VBS_preMap0")
	 g_MapPostPtrfun = Array("VBS_postMap0")
	 
	 rem ϵͳ��������Ϣ,�û�ѡ��ķ�Χ��ID,�ɹ�ͼ����
	 Dim str_msg,str_selectObjid,str_mapName 
	 
	 rem ��ȡϵͳ����--�û�ѡ��Χ��ID
	 SSParameter.GetParameterINT "printMap", "SelectID", -1, str_selectObjid 

	 rem ��ȡϵͳ����--ϵͳ��Ϣ ��0���¹��̶̹�Ŀ¼��ͼ��ʼ����Ϣ  1�������̳�ͼ��ʼ����Ϣ  2: �¹����Զ���Ŀ¼��ͼ��ʼ����Ϣ  3����ͼ����ɽ����ڽű�����ϸ�ڣ�
	 SSParameter.GetParameterINT "printMap", "printMSG", -1, str_msg  

	 rem ��ȡϵͳ����--ר������
	 SSParameter.GetParameterSTR "printMap", "SpecialMapName", "", str_mapName 

	 DistributeMSG str_msg,str_mapName,str_selectObjid 


End Sub




rem ���������������޸�
Function DistributeMSG(MSGid,str_MapName,selectID)
	 dim pFun
	 
	 For i = 0 to ubound(g_MapList) 
		 IF Ucase(g_MapList(i)) = Ucase(str_MapName) Then 
			  IF MSGid = 3 Then 
	 
				  Set pFun = GetRef(g_MapPostPtrfun(i)) 
				  Call pFun(MSGid,str_MapName,selectID) 
	 
			  ELSE 
	 
				  Set pFun = GetRef(g_MapPrePtrfun(i)) 
				  Call pFun(MSGid,str_MapName,selectID) 
	 
			  END IF  
			 Exit For  
		 End IF 
	 Next 
End Function 

'// ���ɹ�Ŀ¼�Ƿ���ڡ���������ڷ�����ͼ
Function CheckReportPath(path_print)

	Dim fso
	Set fso = CreateObject("scripting.filesystemobject")
	
	
	Dim path_thisedb
	strProjectName=SSProcess.GetProjectFileName()
	path_print=replace(strProjectName,".edb"," ")
	

	b1 = fso.FolderExists(path_print)

	
	If  b1 = False  Then 
		
		CheckReportPath = False 
	Else 
		CheckReportPath = True 
	
	End If 

End Function 

'// ��ȡ��������Ŀ����
Function GetXMMC(xmmc)

	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9410001" 
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()

	If geocount <> 1 Then GetXMMC =  0 : Exit Function 
	
	xmmc = SSProcess.GetSelGeoValue(0,"[XiangMMC]")

	If xmmc = "" Or xmmc = "*" Then Exit Function 
	
	GetXMMC = 1

End Function 

'// �ж��ļ��Ƿ����
Function FileExists(fileName)
	Dim fso
	Set fso = CreateObject("scripting.filesystemobject")
	FileExists = fso.FileExists(fileName)
End Function 

'�����ļ���
function CreateFolders(pathname)
	Set fso = CreateObject("Scripting.FileSystemObject")
	newpathname= pathname
	if Not fso.folderExists(newpathname)  then
		fso.CreateFolder   newpathname   '�����ļ���
	end if
	Set fso = Nothing
end function 


function fxtl()
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9310093
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	if geoCount>0 then
		for i = 0 to geoCount-1
			TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
			SSProcess.GetObjectPoint TKID, 0, x, y, z, pointtype, name
			ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"9310082,9310091,FX001,9310022,9310011,9310021,9310001,9310092,9310072,9310032,9310062,9310052,9410021,9410031,9410041,9410051,9410061,9410011,9410001,9310032", 0)
			if ids<> "" then
				SSFunc.ScanString ids, ",", vArray, nCount
				vArray=split(ids,",")
				nCount=ubound(vArray)+1
				ZDrawCode = ""
				FOR j=0 to nCount-1
					DrawCode=SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
					DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
					DrawName = SSProcess.GetFeatureCodeInfo (DrawCode,"ObjectName")
					IF ZDrawCode="" THEN
						ZDrawCode = DrawCode
						ZDrawColor = DrawColor
						ZDrawName = DrawName
					ELSE
					  if replace(ZDrawCode,DrawCode,"")=ZDrawCode then
						ZDrawCode = ZDrawCode&","&DrawCode
						ZDrawColor = ZDrawColor&","&DrawColor
						ZDrawName = ZDrawName&","&DrawName
						end if 
					END IF
			  Next
			end if 
			'LvDiTuLiZPT x-16,y,TKID,ZDrawCode,ZDrawColor,ZDrawName
LvDiTuLiZPT x,y,TKID,ZDrawCode,ZDrawColor,ZDrawName
		next
	end if
End function


function LvDiTuLiZPT(x0,y0,polygonID,ZDrawCode,ZDrawColor,ZDrawName)

		wid1 = 228 : heig1 = 286
		wid2 = 200 : heig2 = 200
		arDrawCode = split(ZDrawCode,",")
		arDrawColor = split(ZDrawColor,",")
		arDrawName = split(ZDrawName,",")
		count5 = ubound(arDrawCode)+2
       '����
         makeLine x0,y0,x0,y0+count5*3+2.5,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID

			makeLine x0+18,y0,x0+18,y0+count5*3+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
		 '����
			'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0,x0+18,y0,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0+count5*3+2.5,x0+18,y0+count5*3+2.5,1, "RGB(255,255,255)", polygonID
			makeNote x0+8,y0+count5*3+1 , 0, "RGB(255,255,255)", wid2, heig2, "ͼ��",polygonID

			for j= 0 to ubound(arDrawCode)
			 '����
               CodeType=SSProcess.GetFeatureCodeInfo(arDrawCode(j), "Type") 
               'makeLine x0+1,y0+j*2+1.5,x0+7,y0+j*2+1.5,arDrawCode(j), arDrawColor(j), polygonID
			      'makeLine x0,y0+j*2+2.5,x0+16,y0+j*2+2.5, 1,"RGB(255,255,255)", polygonID
					'makeNote x0+10,y0+1.5+ j*2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
               if CodeType=3 or  CodeType=2 or  CodeType=1  then '��
               makeLine x0+1,y0+j*3+1.5,x0+5,y0+j*3+1.5,arDrawCode(j), arDrawColor(j), polygonID
					makeNote x0+9,y0+1.5+ j*3, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
               
                elseif CodeType=0  then
               makePoint x0+2.5 ,y0+1.5 +j*3,arDrawCode(j), arDrawColor(j), polygonID
					makeNote x0+9,y0+1.5+ j*3, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID

                elseif CodeType=5  then
					makeArea x0+0.5,y0+0.5 +j*3,x0+5,y0+0.5+ j*3 ,x0+5,y0+2.5+ j*3,x0+0.5,y0+2.5 +j*3,arDrawCode(j), arDrawColor(j), polygonID
					makeNote x0+20,y0+1.5+ j*3, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
                end if 
			next

end function 

function makePoint(x,y,code,color,polygonID)
		SSProcess.CreateNewObj 0
		SSProcess.SetNewObjValue "SSObj_Code", code
		SSProcess.SetNewObjValue "SSObj_Color", color
		SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
		SSProcess.SetNewObjValue "SSObj_LayerName", "����ƽ��ͼͼ����Ϣ"
		SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
		SSProcess.AddNewObjPoint x, y, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 

function makeLine(x1,y1,x2,y2,code, color, polygonID)
		SSProcess.CreateNewObj 1
		SSProcess.SetNewObjValue "SSObj_Code", code
		SSProcess.SetNewObjValue "SSObj_Color", color
		SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
		SSProcess.SetNewObjValue "SSObj_LayerName", "����ƽ��ͼͼ����Ϣ"
		SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
		SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
		SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 

function makeArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID)
		SSProcess.CreateNewObj 2
		SSProcess.SetNewObjValue "SSObj_Code", code
		SSProcess.SetNewObjValue "SSObj_Color", color
		SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
		SSProcess.SetNewObjValue "SSObj_LayerName", "����ƽ��ͼͼ����Ϣ"
		SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
		SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
		SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
		SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
		SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
		SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 

function makeNote(x, y, code, color, width, height, fontString,polygonID)
		SSProcess.CreateNewObj 3
		SSProcess.SetNewObjValue "SSObj_FontClass", "FX001"
		SSProcess.SetNewObjValue "SSObj_FontString", fontString
		SSProcess.SetNewObjValue "SSObj_Color", color
		SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
		SSProcess.SetNewObjValue "SSObj_LayerName", "����ƽ��ͼͼ����Ϣ"
		SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
		SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
		SSProcess.SetNewObjValue "SSObj_FontWidth", width
		SSProcess.SetNewObjValue "SSObj_FontHeight", height
		SSProcess.AddNewObjPoint x, y, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 
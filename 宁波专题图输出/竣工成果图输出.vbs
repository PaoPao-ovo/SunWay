	  	Dim  fileName 
	  	Dim xmmc 
		DIM arID(100000),arID1(100000),arID2(100000)
		dim vArray1(20000), vArray2(20000), vArray3(20000)
		dim cvArray1(20000), cvArray2(20000), cvArray3(20000),vArray(30000)
Rem special:[��ƽͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap0(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
	   fileName= FileFolder & "\���������ƽ��ͼ.edb"
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  JGZPTKEY	selectID
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 

Rem special:[��ƽͼ] ��ͼ����ɴ˽���
Function VBS_postMap0(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()
	SSProcess.SetMapScale "500"

	'DaHui
	'DeleteFeature "9410091","9420033"
	'DeleteFeature "9420035","9999403"
	DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,����ע��,TKZSX,TKZSM,DEFAULT"
	CreateKEYZPT
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[�Ա�ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap1(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		 SSParameter.SetParameterINT "printMap", "return", 1
		 Dim path_print
		If CheckReportPath(path_print) = False    Then 

	  		MsgBox "�޷���ɳ�ͼ���ɹ�Ŀ¼δ�������޷���ɳ�ͼ"
	  		Exit Function 
	  	End If 
	  	

	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 


	    fileName= path_print & "\" & xmmc & "�����ߴ�Ա�ͼ.edb"
	    
	  	If FileExists(fileName) Then 
	  		MsgBox fileName & "�ļ��Ѵ��ڡ��޷������������ֶ����ɾ��������"
	  		Exit Function 
	  	End If 	 
	  	SSParameter.SetParameterINT "printMap", "return", 0
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[�Ա�ͼ] ��ͼ����ɴ˽���
Function VBS_postMap1(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 


	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()



		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[���غ���] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap2(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		 SSParameter.SetParameterINT "printMap", "return", 1
		 Dim path_print
		If CheckReportPath(path_print) = False    Then 

	  		MsgBox "�޷���ɳ�ͼ���ɹ�Ŀ¼δ�������޷���ɳ�ͼ"
	  		Exit Function 
	  	End If 
	  	

	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 


	    fileName= path_print & "\" & xmmc & "���غ������ͼ.edb"
	    
	  	If FileExists(fileName) Then 
	  		MsgBox fileName & "�ļ��Ѵ��ڡ��޷������������ֶ����ɾ��������"
	  		Exit Function 
	  	End If 	 
	  	SSParameter.SetParameterINT "printMap", "return", 0
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[���غ���] ��ͼ����ɴ˽���
Function VBS_postMap2(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 


	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()



		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[����ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap3(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		'����ؿ���ͼ�������
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420025
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	


		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb  mdbName
		sql = "select VALUE from PROJECTINFO where KEY='��Ŀ����'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='������Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			HTRY=arSeletionRecord(0)
		Else
			HTRY=""
		End If

		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			JCRY=arSeletionRecord(0)
		Else
			JCRY=""
		End If

		strtemp = XMMC&","& HTRY &","&JCRY
		SSProcess.CloseAccessMdb mdbName 

		SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
	next


			strProjectName=SSProcess.GetProjectFileName()
			FileFolder=replace(strProjectName,".edb","")
			if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
			LiMianCL "��������ƽ��ʾ��ͼ",FileFolder, intCount
			fileName= FileFolder & "\��������ƽ��ʾ��ͼ"&intCount&".edb"
			SSParameter.SetParameterSTR "printMap","NewedbName",fileName 


	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 

	  
	 End If 
	  
End Function 
	  
	  
Rem special:[����ͼ] ��ͼ����ɴ˽���
Function VBS_postMap3(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 
    GHXKZGUID = SSProcess.GetObjectAttr (selectID,"[JSGHXKZGUID]")
	 jzdtguid = SSProcess.GetObjectAttr (selectID,"[JZWMCGUID]")
	 GHXKZHoutmap = SSProcess.GetObjectAttr (selectID,"[GuiHXKZBH]")
    JianZWMC=SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
    JiDMJ=SSProcess.GetObjectAttr (selectID,"[JiDMJ]")

		SSProcess.SetObjectAttr tk_id,"[JSGHXKZGUID]",GHXKZGUID
		SSProcess.SetObjectAttr tk_id,"[JZWMCGUID]",jzdtguid
		SSProcess.SetObjectAttr tk_id,"[GuiHXKZBH]",GHXKZHoutmap
		SSProcess.SetObjectAttr tk_id,"[JianZWMC]",JianZWMC
		SSProcess.SetObjectAttr tk_id,"[JiDMJ]",JiDMJ
	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()
	CreateKEYJD()
	reset()

	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420032
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	

		ids = SSProcess.SearchInnerObjIDs(id, 2, "9420025", 0)

		If ids<>"" Then
			idsList=split(ids,",")
			strtemp = SSProcess.GetObjectAttr(idsList(0), "SSObj_DataMark")
			artemp = split(strtemp,",")
		Else
			strtemp = ",,"
			artemp = split(strtemp,",")
		End If

		SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
		SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
		SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'ͼ����������
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next

	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 


Rem special:[�ֲ�ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap4(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		TKFZ1
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder

		ZRZH = SSProcess.GetObjectAttr (selectID, "[LD]")
	   fileName= FileFolder &"\"&ZRZH &"�������ܷ�������ʵ��ƽ��ͼ.edb"
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
		SSParameter.SetParameterINT "printMap", "return", 1	  	
	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 
		'SSProcess.WriteEpsIni "��ǰ�����ֲ�ͼ", "��Ŀ����" , xmmc

	  	dh = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
	  	If (dh = "" Or dh = "*") Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ[����������]�Ƿ�Ϸ���"
	  		Exit Function 
	  	End If 
	  	  	SSParameter.SetParameterINT "printMap", "return", 0
	  
	 End If 

End Function 


	  
	  
Rem special:[�ֲ�ͼ] ��ͼ����ɴ˽���
Function VBS_postMap4(MSGID,mapName,selectID)

	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterSTR "printMap", "TKIDS", -1, tk_ids
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 



	 '// ������ĳɹ�ͼ������� 



 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()
	'����
	reset()
	'ͼ����ֵ
	TKFZ2 mark
	if mark=true then
		'�ֲ�ͼ
		FChandle()
		'��ע
		TextEXE()
		'ɾ��¥��
		FCTDeleteLC()
		'����ͼ��
		CreateKEY()
	end if
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[����ͼ ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap5(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder

	   fileName= FileFolder &"\"&"����ͼ.edb"
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
		SSParameter.SetParameterINT "printMap", "return", 1	  	
	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 
		'SSProcess.WriteEpsIni "��ǰ�����ֲ�ͼ", "��Ŀ����" , xmmc

	  	Dim dh
	  	dh = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
	  	
	  	If (dh = "" Or dh = "*") Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ[����������]�Ƿ�Ϸ���"
	  		Exit Function 
	  	End If 
	  	  	SSParameter.SetParameterINT "printMap", "return", 0
	  
	 End If 
	  
End Function 

Rem special:[����ͼ] ��ͼ����ɴ˽���
Function VBS_postMap5(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	  	JianZWMC = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
      GuiHXKZBH=SSProcess.GetObjectAttr (selectID,"[GuiHXKZBH]")
      SSProcess.SetObjectAttr tk_id,"[JianZWMC]",JianZWMC
    SSProcess.SetObjectAttr tk_id,"[GuiHXKZBH]",GuiHXKZBH
	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()



		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 


Rem special:[����ͣ��λ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap6(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		 SSParameter.SetParameterINT "printMap", "return", 1
		 Dim path_print
		If CheckReportPath(path_print) = False    Then 

	  		MsgBox "�޷���ɳ�ͼ���ɹ�Ŀ¼δ�������޷���ɳ�ͼ"
	  		Exit Function 
	  	End If 
	  	

	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 


	    fileName= path_print & "\" & xmmc & "����ͣ��λ�ֲ�ͼ.edb"
	    
	  	If FileExists(fileName) Then 
	  		MsgBox fileName & "�ļ��Ѵ��ڡ��޷������������ֶ����ɾ��������"
	  		Exit Function 
	  	End If 	 
	  	SSParameter.SetParameterINT "printMap", "return", 0
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[����ͣ��λ] ��ͼ����ɴ˽���
Function VBS_postMap6(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 



	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()



		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[����ͣ��λ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap7(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	  
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		 SSParameter.SetParameterINT "printMap", "return", 1
		 Dim path_print
		If CheckReportPath(path_print) = False    Then 

	  		MsgBox "�޷���ɳ�ͼ���ɹ�Ŀ¼δ�������޷���ɳ�ͼ"
	  		Exit Function 
	  	End If 
	  	

	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 


	    fileName= path_print & "\" & xmmc & "����ͣ��λ�ֲ�ͼ.edb"
	    
	  	If FileExists(fileName) Then 
	  		MsgBox fileName & "�ļ��Ѵ��ڡ��޷������������ֶ����ɾ��������"
	  		Exit Function 
	  	End If 	 
	  	SSParameter.SetParameterINT "printMap", "return", 0
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 

Rem special:[����ͣ��λͼ] ��ͼ����ɴ˽���
Function VBS_postMap7(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterSTR "printMap", "TKIDS", -1, tk_ids
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 
    GHXKZGUID = SSProcess.GetObjectAttr (selectID,"[JSGHXKZGUID]")
	 jzdtguid = SSProcess.GetObjectAttr (selectID,"[JZWMCGUID]")
	 GHXKZHoutmap = SSProcess.GetObjectAttr (selectID,"[GuiHXKZBH]")
	' GetXKZXX GHXKZHoutmap,JZWMCoutmap,GHXKZGUID,jzdtguid
	 BLC = SSProcess.GetMapScale

'msgbox jzdtguid
	 '// ������ĳɹ�ͼ������� 
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_ID", "==", tk_ids
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	dim arcc(1000)
	For i = 0 To geoCount -1
		tk_id = SSProcess.GetSelGeoValue(i,"SSObj_ID")

		cc = SSProcess.GetSelGeoValue(i,"[CengC]")

	'	GetCGXX GHXKZHoutmap,JZWMCoutmap,cc,GHCD,SJCG,CMXX,sjcs

		If InStr(cc,"-") = 1 Then
			DSDXBS = "����"
		Else
			DSDXBS = "����"
		End If
		'If GHCD <> "" Then  SSProcess.SetObjectAttr tk_id,"[PiZCG]",GHCD
		'If SJCG <> "" Then  SSProcess.SetObjectAttr tk_id,"[ShiCCG]",SJCG
		SSProcess.SetObjectAttr tk_id,"[BiLC]",BLC
		'SSProcess.SetObjectAttr tk_id,"[DiSDXBS]",DSDXBS
		'SSProcess.SetObjectAttr tk_id,"[CengS]",sjcs
		SSProcess.SetObjectAttr tk_id,"[JSGHXKZGUID]",GHXKZGUID
		SSProcess.SetObjectAttr tk_id,"[JZWMCGUID]",jzdtguid
		SSProcess.SetObjectAttr tk_id,"[GuiHXKZBH]",GHXKZHoutmap
		'SSProcess.SetObjectAttr tk_id,"[BeiZ]","˵����1���ò㽨�������ʵ��ߴ���㡣     \2��ʵ��ߴ���֪�۳�Ĩ�Һ�ȣ�Ĩ�Һ��ƽ��0.03m����"
		If InStr(cc,"-") > 0 Then 
			cc = Split(cc,"+")
			If UBound(cc) =1 Then 
				if instr(cc(0),".")= 0 then 
					str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
				else	
					str0 = cc(0)
				end if 
				if instr(cc(1),".") = 0 then 
					str1 = SSFunc.GetChineseDigit(Abs(cc(1)))
				else
					str1 = cc(1)
				end if 
				str111 = "����" & str0 & "����" & str1 & "��"
			ElseIf UBound(cc) = 0 Then 
				 if instr(cc(0),".") = 0 then 
					str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
				else
					str0 = cc(0)
				end if 
				str111  = "����" & str0 & "��"
			End If 
		Else 
			SSFunc.ScanString cc, ",", arcc, arccCount
			For c = 0 to arccCount-1
					cc = Split(arcc(c),"+")
					If UBound(cc) =1 Then 
						if instr(cc(0),".") = 0 then 
							str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
						else
							str0 = cc(0)
						end if 

						if instr(cc(1),".") = 0 then 
							str1 = SSFunc.GetChineseDigit(Abs(cc(1)))
						else
							str1 = cc(1)
						end if 
						str =  str0 & "����" & str1 & "��"
					ElseIf UBound(cc) = 0 Then 
						if instr(cc(0),".") = 0 then 
							str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
						else
							str0 = cc(0)
						end if 
						str  =  str0 & "��"
					End If 
				If str111 ="" Then
					str111 = str
				Else 
					str111 = str111&"��"&str
				End If
				str =""
			Next

		End If 

		SSProcess.SetObjectAttr tk_id,"[CengM]",str111

		str111=""

		ids = SSProcess.SearchInnerObjIDs( tk_id, 10, "1", 0 ) 
		if ids <> "" then 
			if change_ids = "" then 
					change_ids = ids
			else
					change_ids = change_ids & "," & ids 
			end if 

		end if 

	Next 


 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()
		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[�̵�ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap8(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
		 SSParameter.SetParameterINT "printMap", "return", 1
		 Dim path_print
		If CheckReportPath(path_print) = False    Then 

	  		MsgBox "�޷���ɳ�ͼ���ɹ�Ŀ¼δ�������޷���ɳ�ͼ"
	  		Exit Function 
	  	End If 
	  	

	  	If GetXMMC(xmmc) = False Then 
	  		MsgBox "�޷���ɳ�ͼ��������Ŀ�����Ƿ���ȷ��"
	  		Exit Function 
	  	End If 


	    fileName= path_print & "\" & xmmc & "�̵����ͳ��ͼ.edb"
	    
	  	If FileExists(fileName) Then 
	  		MsgBox fileName & "�ļ��Ѵ��ڡ��޷������������ֶ����ɾ��������"
	  		Exit Function 
	  	End If 	 
	  	SSParameter.SetParameterINT "printMap", "return", 0
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[�̵�ͼ] ��ͼ����ɴ˽���
Function VBS_postMap8(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()



		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 


Rem special:[����ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap9(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	


		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb  mdbName
		sql = "select VALUE from PROJECTINFO where KEY='��Ŀ����'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount

		If nSeletionCount>0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC = ""
		End If
		

		sql = "select VALUE from PROJECTINFO where KEY='������Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			HTRY=arSeletionRecord(0)
		Else
			HTRY = ""
		End If

		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			JCRY=arSeletionRecord(0)
		Else
			JCRY = ""
		End If
	
		sql = "select VALUE from PROJECTINFO where KEY='��浥λ'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			CLDW=arSeletionRecord(0)
		Else
			CLDW = ""
		End If

		strtemp = XMMC&","& HTRY &","&JCRY&","&CLDW
		SSProcess.CloseAccessMdb mdbName 

		SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
	next

	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
	   fileName= FileFolder & "\����ͼ.edb"
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[����ͼ] ��ͼ����ɴ˽���
Function VBS_postMap9(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()

		'DeleteFeature "9410101","9420032"
		'DeleteFeature "9420034","9999403"
		'DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,����ע��,TKZSX,TKZSM"
		DaHui
		SSProcess.SetMapScale "500"

	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420033
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	

		ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)

		If ids<>"" Then
			idsList=split(ids,",")
			strtemp = SSProcess.GetObjectAttr(idsList(0), "SSObj_DataMark")
			artemp = split(strtemp,",")
		Else
			strtemp = ",,,"
			artemp = split(strtemp,",")
		End If
		
		SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
		SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
		SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
		SSProcess.SetObjectAttr id, "[������λ����]", artemp(3)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'ͼ����������
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next
	DeleteFeatureLayerName "�滮��,GHX"

	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 


Rem special:[�����滮����ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap10(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	


		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb  mdbName
		sql = "select VALUE from PROJECTINFO where KEY='��Ŀ����'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='������Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			HTRY=arSeletionRecord(0)
		Else
			HTRY=""
		End If

		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			JCRY=arSeletionRecord(0)
		Else
			JCRY=""
		End If

		sql = "select VALUE from PROJECTINFO where KEY='��浥λ'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			CLDW=arSeletionRecord(0)
		Else
			CLDW=""
		End If


		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			SHRY=arSeletionRecord(0)
		Else
			SHRY=""
		End If


		strtemp = XMMC&","& HTRY &","&JCRY&","&CLDW&","&SHRY
		SSProcess.CloseAccessMdb mdbName 

		SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
	next
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
	   fileName= FileFolder & "\�����滮����ͼ.edb"
	  	
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[�����滮����ͼ] ��ͼ����ɴ˽���
Function VBS_postMap10(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()

		'DaHui
		DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,����ע��,TKZSX,TKZSM"
		SSProcess.SetMapScale "500"

	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420035
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	

		ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)

		If ids<>"" Then
			idsList=split(ids,",")
			strtemp = SSProcess.GetObjectAttr (idsList(0), "SSObj_DataMark")
			artemp = split(strtemp,",")
		Else
			strtemp = ",,,,"
			artemp = split(strtemp,",")
		End If

		SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
		SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
		SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
		SSProcess.SetObjectAttr id, "[������λ����]", artemp(3)
		SSProcess.SetObjectAttr id, "[ShenHY]", artemp(4)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'ͼ����������
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next		

		CreateKEYZPT()
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[�õظ���ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap11(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	


		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb  mdbName
		sql = "select VALUE from PROJECTINFO where KEY='��Ŀ����'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='������Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			HTRY=arSeletionRecord(0)
		Else
			HTRY=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			JCRY=arSeletionRecord(0)
		Else
			JCRY=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='��浥λ'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			CLDW=arSeletionRecord(0)
		Else
			CLDW=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount>0 Then
			SHRY=arSeletionRecord(0)
		Else
			SHRY=""
		End If

		strtemp = XMMC&","& HTRY &","&JCRY&","&CLDW&","&SHRY
		SSProcess.CloseAccessMdb mdbName 

		SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
	next
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
	   fileName= FileFolder & "\�õظ���ͼ.edb"
	  	
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[�õظ���ͼ] ��ͼ����ɴ˽���
Function VBS_postMap11(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()

		DaHui
		DeleteFeature "9410011","9420035"
		DeleteFeature "9420037","9999403"
		DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,����ע��,TKZSX,TKZSM"

		JZDSC
		JZX

	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420036
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	

		ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)

		If ids<>"" Then
			idsList=split(ids,",")
			strtemp = SSProcess.GetObjectAttr (idsList(0), "SSObj_DataMark")
			artemp = split(strtemp,",")
		Else
			strtemp = ",,,,"
			artemp = split(strtemp,",")
		End If

		


		SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
		SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
		SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
		SSProcess.SetObjectAttr id, "[������λ����]", artemp(3)
		SSProcess.SetObjectAttr id, "[ShenHY]", artemp(4)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'ͼ����������
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next		
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[������ƽ�沼�ú�ʵ����ƽ��ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap12(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
	   fileName= FileFolder & "\������ƽ�沼�ú�ʵ����ƽ��ͼ.edb"
	  	
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[������ƽ�沼�ú�ʵ����ƽ��ͼ] ��ͼ����ɴ˽���
Function VBS_postMap12(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()

		DaHui

		DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,����ע��,TKZSX,TKZSM"

		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 

Rem special:[��ƽ�������ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap13(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder
	   fileName= FileFolder & "\��ƽ�������ͼ.edb"


		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[��ƽ�������ͼ] ��ͼ����ɴ˽���
Function VBS_postMap13(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()

		DaHui
		DeleteFeature "9410011","9420037"
		DeleteFeature "9420039","9420108"
		DeleteFeature "9450013","9999403"
		DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,����ע��,TKZSX,TKZSM"

		
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 


Rem special:[����ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap14(MSGID,mapName,selectID)
	  
	 Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
	 Rem return = 1 ֹͣ����ɹ�ͼ
	 Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
	 If MSGID = 0 Then '// �¹��̳�ͼ 
		 '// ������Ĵ���.... 
		 '// ���ó�ͼ�������ơ��������.... 
	   strProjectName=SSProcess.GetProjectFileName()
	   FileFolder=replace(strProjectName,".edb","")
		if IsfolderExists(FileFolder) = false then CreateFolders FileFolder

		LiMianCL "����ͼ",FileFolder, intCount
	   fileName= FileFolder & "\����ͼ"&intCount&".edb"
	  	
		SSParameter.SetParameterSTR "printMap","NewedbName",fileName 
	  
	  
	 ElseIf MSGID = 1 Then '// �����̳�ͼ 
		 '// ������Ĵ���.... 
	  
	 End If 
	  
End Function 


	  
	  
Rem special:[����ͼ] ��ͼ����ɴ˽���
Function VBS_postMap14(MSGID,mapName,selectID)
	  
	 Rem ͼ��ID,�ű����������
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 Rem �ű�����������,�ű����������,�ű�������Ӳ��� 
	 Dim str_Name,str_para,str_paraex	  
	 Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 Rem ��ȡͼ���ڵ���IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 Rem ��ȡ�ű����������
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
		 '// ������ĳɹ�ͼ������� 

	SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc

	
 '// ������ĳɹ�ͼ������� 
 	debug_print String(50,"-")
	debug_print "�����ɡ�"
	debug_print String(50,"-")
	ViewExtend()

	
	
	 Rem �ɹ�ͼϸ�ڷֿ�����
	 For i = 0 to ScriptChangeCount -1
		 Rem ��ȡ����������
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 Rem ��ȡ���������
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 Rem ��ȡ������Ӳ���
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 '// �˴��޴��롢˵��û�нű�������..
	 Next 
	  
End Function 


function LiMianCL(byval mapname,byval FileFolder,byref intCount)
	intCount = 1
	'����FileSystemObject����
	Set objFso= CreateObject("Scripting.FileSystemObject")
	'ʹ��GetFolder()����ļ��ж���
	Set objGetFolder = objFso.GetFolder(FileFolder)
	'����Files���ϲ���ʾ�ļ��������е��ļ���
	For Each strFile in objGetFolder.Files		
			if objFso.GetExtensionName(strFile)="edb" then
				if instr(strFile.Name ,mapname)>0 then intCount = intCount + 1
			end if
	Next

end function


Dim g_MapList,g_MapPrePtrfun,g_MapPostPtrfun 

Sub OnClick() 
	 
	rem ��ʼ�� 
	 g_MapList = Array("��ƽ��ͼ","�����ߴ�Ա�ͼ","���غ���ͼ","��������ͼ","�ֲ��������ͼ","��������ʾ��ͼ","����ͣ��λ�ֲ�ͼ","����ͣ��λ�ֲ�ͼ","�̵����ͳ��ͼ","����ͼ","�����滮����ͼ","�õظ���ͼ","������ƽ�沼�ú�ʵ����ƽ��ͼ","��ƽ�������ͼ","����ͼ")
	 g_MapPrePtrfun = Array("VBS_preMap0","VBS_preMap1","VBS_preMap2","VBS_preMap3","VBS_preMap4","VBS_preMap5","VBS_preMap6","VBS_preMap7","VBS_preMap8","VBS_preMap9","VBS_preMap10","VBS_preMap11","VBS_preMap12","VBS_preMap13","VBS_preMap14")
	 g_MapPostPtrfun = Array("VBS_postMap0","VBS_postMap1","VBS_postMap2","VBS_postMap3","VBS_postMap4","VBS_postMap5","VBS_postMap6","VBS_postMap7","VBS_postMap8","VBS_postMap9","VBS_postMap10","VBS_postMap11","VBS_postMap12","VBS_postMap13","VBS_postMap14")
	 
	 rem ϵͳ��������Ϣ,�û�ѡ��ķ�Χ��ID,�ɹ�ͼ����
	 Dim str_msg,str_selectObjid,str_mapName 
	 
	 rem ��ȡϵͳ����--�û�ѡ��Χ��ID
	 SSParameter.GetParameterINT "printMap", "SelectID", -1, str_selectObjid 

	 rem ��ȡϵͳ����--ϵͳ��Ϣ ��0�������̳�ͼ��ʼ����Ϣ 1���¹��̶̹�Ŀ¼��ͼ��ʼ����Ϣ  2: �¹����Զ���Ŀ¼��ͼ��ʼ����Ϣ  -1����ͼ����ɽ����ڽű�����ϸ�ڣ�
	 SSParameter.GetParameterINT "printMap", "printMSG", -1, str_msg  

	 rem ��ȡϵͳ����--ר������
	 SSParameter.GetParameterSTR "printMap", "SpecialMapName", "", str_mapName 
	 DistributeMSG str_msg,str_mapName,str_selectObjid 
End Sub

Sub ViewExtend()

		'ͼ�η�Χȫ��

		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0

		'ͼ����������

		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0

End Sub

'// �ж��ļ��Ƿ����
Function FileExists(fileName)


	Dim fso
	Set fso = CreateObject("scripting.filesystemobject")
	FileExists = fso.FileExists(fileName)

End Function 

'�����ļ���
Function CreateFolders(path)
    Set fso = CreateObject("scripting.filesystemobject")
    CreateFolderEx fso,path
    set fso = Nothing
End Function
 
Function CreateFolderEx(fso,path)
    If fso.FolderExists(path) Then
        Exit Function
    End If
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolderEx fso,fso.GetParentFolderName(path)
    End If
    fso.CreateFolder(path)
End Function

'�ж��ļ����Ƿ����
Function IsfolderExists(folder)
	Dim fso
	Set fso=CreateObject("Scripting.FileSystemObject")        
	If fso.folderExists(folder) Then
		IsfolderExists = True
	Else 
		IsfolderExists = False
	End If 
End Function 


rem ���������������޸�
Function DistributeMSG(MSGid,str_MapName,selectID)
	 dim pFun
	 
	 For i = 0 to ubound(g_MapList) 
		 IF Ucase(g_MapList(i)) = Ucase(str_MapName) Then 
			 IF MSGid =3  Then 

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


Function debug_print(str)

	SSProcess.MapCallBackFunction "OutputMsg", STR & "	" & Now , 0 

End Function 

'// ���ɹ�Ŀ¼�Ƿ���ڡ���������ڷ�����ͼ
Function CheckReportPath(path_print)

	Dim fso
	Set fso = CreateObject("scripting.filesystemobject")
	
	Dim path_thisedb
	path_thisedb = SSProcess.GetSysPathName( 5)
	
	path_print = path_thisedb 
	
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

'��ȡ�滮���guid
Function GetXKZXX(ghxkzbh,jzwmc,GHXKZGUID,jzdtguid)

	Dim arID(20)
  sql = "SELECT JG_���蹤�̽���������Ϣ���Ա�.GuiHXKZGUID,JZWMCGUID,GuiHXKZBH FROM JG_���蹤�̽���������Ϣ���Ա�  WHERE (JG_���蹤�̽���������Ϣ���Ա�.GHXKZBH = '"&ghxkzbh&"' AND JG_���蹤�̽���������Ϣ���Ա�.JianZWMC = '"&jzwmc&" ');"
  projectname= SSProcess.GetProjectFileName
  SSProcess.OpenAccessMdb projectname
  SSProcess.OpenAccessRecordset projectname, sql
  recordCount = SSProcess.GetAccessRecordCount (projectname, sql ) 
	if recordCount > 0 then
		SSProcess.AccessMoveFirst projectname,sql
		while (SSProcess.AccessIsEOF (projectName, sql ) = False)
			SSProcess.GetAccessRecord projectName, sql, fields, values							  
			SSFunc.ScanString values, ",", arID, idCount
			GHXKZGUID=arID(0)
			jzdtguid=arID(1)
			ghxkzbh=arID(2)
			SSProcess.AccessMoveNext projectName, sql
		Wend
	End If	
	   SSProcess.CloseAccessRecordset projectName, sql
      SSProcess.CloseAccessMdb projectName 

End Function 

function DaHui
		SSProcess.PushUndoMark 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_LayerName", "==", "DEFAULT,��ע��,�������Ƶ�,��ѧ����,ˮϵ��,ˮϵ��,ˮϵ��,����ص�,�������,�������,��ͨ��,��ͨ��,��ͨ��,���ߵ�,������,������,�����,������,������,��ò��,��ò��,��ò��,ֲ�������ʵ�,ֲ����������,ֲ����������,��������,ˮϵ������,��·������,ˮϵע��,�����ע��,��ͨע��,����ע��,����ע��,��òע��,ֲ��ע��,����������,��������,ͼ����,������,��ʩ��,��ʩ��,��ʩ��,����������,����������,ԭʼ�۲��,���ƺ�,�ȸ���,�̵߳�,��ά�ӽǵ��,������ʩ��,������ʩ��,������ʩ��,����������,�������̵�,����������,DMTZ,GXYZ,KZD,JMD,DLDW,ZBTZ,SXSS,GCD,DGX,ZDH,ZJ,��λ��" 
		SSProcess.SelectFilter
		geocount = SSProcess.GetSelGeoCount()
		for i=0 to geocount-1
			geoID= SSProcess.GetSelGeoValue(i, "SSObj_ID") 
			SSProcess.SetObjectAttr geoID, "SSObj_Color", RGB(0,0,0)
		next

		SSProcess.PushUndoMark 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_LayerName", "==", "DEFAULT,��ע��,�������Ƶ�,��ѧ����,ˮϵ��,ˮϵ��,ˮϵ��,����ص�,�������,�������,��ͨ��,��ͨ��,��ͨ��,���ߵ�,������,������,�����,������,������,��ò��,��ò��,��ò��,ֲ�������ʵ�,ֲ����������,ֲ����������,��������,ˮϵ������,��·������,ˮϵע��,�����ע��,��ͨע��,����ע��,��òע��,ֲ��ע��,����������,��������,ͼ����,������,��ʩ��,��ʩ��,��ʩ��,����������,����������,ԭʼ�۲��,���ƺ�,�ȸ���,�̵߳�,��ά�ӽǵ��,������ʩ��,������ʩ��,������ʩ��,����������,�������̵�,����������,DMTZ,GXYZ,KZD,JMD,DLDW,ZBTZ,SXSS,GCD,DGX,ZDH,ZJ,��λ��" 
		SSProcess.SelectFilter
		notecount= SSProcess.GetSelNoteCount()
		For i1  =0 To notecount- 1
			id = SSProcess.GetSelNoteValue(i1 ,"SSObj_ID" )
			SSProcess.SetObjectAttr id, "SSObj_Color", RGB(0,0,0)
		Next
end function

function DeleteFeature(StartCode,EndCode)
		SSProcess.PushUndoMark 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_Code", ">=", StartCode
		SSProcess.SetSelectCondition "SSObj_Code", "<=", EndCode
		SSProcess.SelectFilter
		SSProcess.DeleteSelectionObj
end function

function DeleteFeatureLayerName(strLayerName)
		SSProcess.PushUndoMark 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_LayerName", "==", strLayerName
		SSProcess.SelectFilter
		SSProcess.DeleteSelectionObj
end function

function JZDSC
		Const JZDBM ="9510041"
		Const QSMBM ="9410001"
		SSProcess.PushUndoMark
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_Code", "==", "9510041" 
		SSProcess.SelectFilter
				geoecount = SSProcess.GetSelgeoCount
				For i=0 To geoecount-1
				  SSProcess.DelSelgeo i
				Next
		SSProcess.ClearSelection
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
		SSProcess.SelectFilter
		GeoCount =SSProcess.GetSelGeoCount  
		For i =0 To GeoCount -1
			AreaPNum = SSProcess.GetSelGeoPointCount(i)
			'Msgbox AreaPNum
			For j = 0 To AreaPNum -2
				SSProcess.GetSelGeoPoint i, j, x,  y,  z,  ptype,  name 
				ids = SSProcess.SearchNearObjIDs(x, y, 0.001, 0, JZDBM, 0 )
				If ids ="" Then
					'Msgbox ids
					SSProcess.CreateNewObjByCode JZDBM
					SSProcess.AddNewObjPoint x, y, 9999, 0, "J"&J+1
					SSProcess.AddNewObjToSaveObjList
				End If
			Next
		Next
		SSProcess.SaveBufferObjToDatabase
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition


		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_Code", "==", JZDBM 
		SSProcess.SelectFilter
		SSProcess.ChangeSelectionObjAttr "SSObj_PointType", "2" 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
end function

FUNCTION JZX
		SSProcess.ClearSelection
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
		SSProcess.SelectFilter
		GeoCount =SSProcess.GetSelGeoCount  
		For i =0 To GeoCount -1
			geoID= SSProcess.GetSelGeoValue(i, "SSObj_ID")
			SSProcess.ChangeCodeCopy geoID,"9510042"
			Maxid=SSProcess.GetGeoMaxID()
			SSProcess.LineCrack Maxid ,0
		Next
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
end function


function FChandle()
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_ID", "==", 1
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	dim strCopy(1000,1000),strCopyID(10000),strCopyID1(10000),strCopyID2(10000),strCopyID3(10000),strCopyID4(10000)
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
		ids = SSProcess.SearchInnerObjIDs(id, 2, "9420008", 0)
		artemp = split(ids,",")
		'���ĵ�����
		SSProcess.GetObjectFocusPoint  id , x0,  y0 
		'���ĵ㵽ͼ����������
		SSProcess.GetObjectPoint     id, 0, x1, y1, z1, pointtype, name
		SSProcess.GetObjectPoint     id, 1, x2, y2, z2, pointtype, name
		SSProcess.GetObjectPoint     id, 2, x3, y3, z3, pointtype, name
		SSProcess.GetObjectPoint     id, 3, x4, y4, z4, pointtype, name
		Length = (x2-x1)+(x2-x0)
		for j = 0 to ubound(artemp)
			LC = SSProcess.GetObjectAttr (artemp(j), "[LC]")
			LCID = SSProcess.GetObjectAttr (artemp(j), "[ID_LC]")
			'��ȡͼ����¥�㡢������ID
			mdbName=SSProcess.GetProjectFileName   '��ǰ���̹�����
			SSProcess.OpenAccessMdb mdbName
			'������			
			strSQL="SELECT JG_�滮���������Ա�.ID from  JG_�滮���������Ա� inner join GeoAreaTB on JG_�滮���������Ա�.ID = GeoAreaTB.ID where ([GeoAreaTB].[Mark] Mod 2)<>0 and JG_�滮���������Ա�.LC = '"&LC&"'"
			GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
			if nSeletionCount>0 then
				strtemp1 = ""
				for k = 0 to nSeletionCount-1
					strtemp = arSeletionRecord(k)
					if strtemp1 = "" then
						strtemp1 = strtemp
					else
						strtemp1 = strtemp1&","&strtemp
					end if
				next
				strCopyID1(j)= id&","&artemp(j)&","&strtemp1
			ELSE
				strCopyID1(j)= id&","&artemp(j)
			end if
			'����������
			strSQL="SELECT JG_�滮�������������Ա�.ID from  JG_�滮�������������Ա� inner join GeoAreaTB on JG_�滮�������������Ա�.ID = GeoAreaTB.ID where ([GeoAreaTB].[Mark] Mod 2)<>0 and JG_�滮�������������Ա�.LC = '"&LC&"'"
			GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
			if nSeletionCount>0 then
				strtemp1 = ""
				for k = 0 to nSeletionCount-1
					strtemp = arSeletionRecord(k)
					if strtemp1 = "" then
						strtemp1 = strtemp
					else
						strtemp1 = strtemp1&","&strtemp
					end if
				next
				strCopyID2(j)= id&","&artemp(j)&","&strtemp1
			ELSE
				strCopyID2(j)= id&","&artemp(j)
			end if
			'������
			strSQL="SELECT JG_�滮���������Ա�.ID from  JG_�滮���������Ա� inner join GeoAreaTB on JG_�滮���������Ա�.ID = GeoAreaTB.ID where ([GeoAreaTB].[Mark] Mod 2)<>0 and JG_�滮���������Ա�.LC = '"&LC&"'"
			GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
			if nSeletionCount>0 then
				strtemp1 = ""
				for k = 0 to nSeletionCount-1
					strtemp = arSeletionRecord(k)
					if strtemp1 = "" then
						strtemp1 = strtemp
					else
						strtemp1 = strtemp1&","&strtemp
					end if
				next
				strCopyID3(j)= id&","&artemp(j)&","&strtemp1
			ELSE
				strCopyID3(j)= id&","&artemp(j)
			end if
			'ע��
			strSQL="SELECT JG_��ͼע�����Ա�.ID from  JG_��ͼע�����Ա�  where  JG_��ͼע�����Ա�.ID_LC = '"&LCID&"'"
			GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
			if nSeletionCount>0 then
				strtemp1 = ""
				for k = 0 to nSeletionCount-1
					strtemp = arSeletionRecord(k)
					if strtemp1 = "" then
						strtemp1 = strtemp
					else
						strtemp1 = strtemp1&","&strtemp
					end if
				next
				strCopyID4(j)= id&","&artemp(j)&","&strtemp1
			ELSE
				strCopyID4(j)= id&","&artemp(j)
			end if
			strCopyID(j)=strCopyID1(j)&","&strCopyID2(j)&","&strCopyID3(j)&","&strCopyID4(j)
			SSProcess.CloseAccessMdb mdbName 
		next
	next

	'ճ��ͼ��
	for j = 0 to ubound(artemp)
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_ID", "==", strCopyID (j)
		SSProcess.SelectFilter
		SSProcess.SelectionObjToClipBoard
		SSProcess.AddClipBoardObjToMap Length*(j+1), 0
	next

	'ɾ��ԭ����
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_ID", "==", 1
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
		ids = SSProcess.SearchInnerObjIDs(id, 10, "", 0)
		arids = split(ids,",")
		for j = 0 to ubound(arids)	
			SSProcess.DeleteObject arids(j) 
		next
		SSProcess.DeleteObject id
		SSProcess.RefreshView 
	next 

end function


function reset
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9420008,9420025"
	SSProcess.SelectFilter 
	SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_500", "Feature.Code=SSObj_Code", "SSObj_Color=Feature.LineColor,SSObj_LineWidth=Feature.LineWidth,SSObj_LayerName=Feature.LayerName,SSObj_Type=Feature.Type"


end function

function TextEXE
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420031
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	dim strCopy(1000,1000),strCopyID(10000)
	For i = 0 To geoCount -1
		'ע����������
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	
		SSProcess.GetObjectPoint id, 0, x0, y0, z0, pointtype, name
		SSProcess.GetObjectPoint id, 1, x1, y1, z1, pointtype, name
		x = (x1-x0)/2 + x0
		y = y0+10
		'����ע��
		ids = SSProcess.SearchInnerObjIDs(id, 2, "9420008", 0)
		LC = SSProcess.GetObjectAttr (ids, "[LC]")
		CQC = SSProcess.GetObjectAttr (ids, "[CQC]")
		if instr(LC,"��")=0 then
			IF CQC<> "�ݶ���"  then
				artemplc = split(LC,".")
				if ubound(artemplc) = 0 then
					'��׼��
					strText = ""
					GetLCXX LC,strText
					strText = strText&"ƽ��ͼ"
					CreateText strText,x,y,z
				elseif ubound(artemplc)>0 then
					'�в�
					LC = mid(LC,1,InStr(LC,".")-1)
					strText = ""
					GetLCXX LC,strText
					strText = strText&"�в�ƽ��ͼ"
					CreateText strText,x,y,z
				end if
			else
					strText = "�ݶ���ƽ��ͼ"
					CreateText strText,x,y,z
			end if
		else
			artemplc=split(LC,"��")
			strText = ""
			GetLCXX artemplc(0),strText0
			GetLCXX	artemplc(1),strText1
			strText = strText0&"��"&strText1&"ƽ��ͼ"
			CreateText strText,x,y,z
		end if
	next
end function


function NumberChange(Number,BigNumber)
		strNumer = "1,2,3,4,5,6,7,8,9"
		strBigNumber = "һ,��,��,��,��,��,��,��,��"
		artempNumber = split(strNumer,",")
		artempBigNumber = split(strBigNumber,",")
		for i = 0 to 8
			if  artempNumber(i) = Number  then
				BigNumber = artempBigNumber(i)
			end if
		next
end function

function CreateText(strText,x,y,z)
		SSProcess.CreateNewObjByClass "0"
		SSProcess.SetNewObjValue "SSObj_FontString", strText
		SSProcess.SetNewObjValue "SSObj_FontWidth", 1000
		SSProcess.SetNewObjValue "SSObj_FontHeight", 1000

		SSProcess.AddNewObjPoint x, y, z, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function

function GetLCXX(LC,strText)
		if Len(LC) = 1 then
			for j = 0 to Len(LC)-1
				IntLC = mid(LC,j+1,1)
				NumberChange IntLC,BigNumber
				strText = strText&BigNumber&"��"
			next
		elseif Len(LC)>1 and instr(LC,"-")=0 then 
			for j = 0 to Len(LC)-1
				IntLC = mid(LC,j+1,1)
				NumberChange IntLC,BigNumber
				if strText = "" then
					strText = BigNumber
				elseif IntLC<>0 then
					strText = strText&"ʮ"&BigNumber&"��"
				elseif IntLC=0 then
					strText = strText&"ʮ"&"��"
				end if
			next
			if  mid(LC,1,1) = 1 then
				strText = mid(strText,2,len(strText))
			end if
		elseif Len(LC)>1 and instr(LC,"-")=1 then 
			for j=1 to Len(LC)-1
				IntLC = mid(LC,j+1,1)
				NumberChange IntLC,BigNumber
				strText = "����"&BigNumber&"��"
			next
		end if

end function 

function TKFZ1
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420004
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	


		mdbName = SSProcess.GetProjectFileName 
		SSProcess.OpenAccessMdb  mdbName
		sql = "select VALUE from PROJECTINFO where KEY='��Ŀ����'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount > 0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='������Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount > 0 Then
			HTRY=arSeletionRecord(0)
		Else
			HTRY=""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='�����Ա'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount > 0 Then
			JCRY=arSeletionRecord(0)
		Else
			JCRY=""
		End If
		
		strtemp = XMMC&","& HTRY &","&JCRY
		SSProcess.CloseAccessMdb mdbName 

		SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
	next
end function

function TKFZ2(byref mark)
	mark=true
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420031
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	For i = 0 To geoCount -1
		id = SSProcess.GetSelGeoValue(i, "SSObj_ID")	

		ids = SSProcess.SearchInnerObjIDs(id, 2, "9420004", 0)

		FWJG = SSProcess.GetObjectAttr (ids, "[FWJG]")
		ZRZH = SSProcess.GetObjectAttr (ids, "[ZRZH]")
		ZCS = SSProcess.GetObjectAttr (ids, "[ZCS]")
		FWZL = SSProcess.GetObjectAttr (ids, "[FWZL]")
		
		idsList=split(ids,",")
		if ubound(idsList)>0 then msgbox "��ͼλ�����ص���Ȼ������ȷ�������Ƿ���ȷ��":mark=false:exit function
		strtemp = SSProcess.GetObjectAttr (ids, "SSObj_DataMark")
		artemp = split(strtemp,",")
		If UBound(artemp)<0 Then
			ReDim artemp(2)
			artemp(0) = ""
			artemp(1) = ""
			artemp(2) = ""
		End If
		SSProcess.SetObjectAttr id, "[FWJG]", FWJG
		SSProcess.SetObjectAttr id, "[ZRZH]", ZRZH
		SSProcess.SetObjectAttr id, "[ZCS]", ZCS
		SSProcess.SetObjectAttr id, "[FWZL]", FWZL
		SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
		SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
		SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'ͼ����������
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next
end function

function FCTDeleteLC
	SSProcess.PushUndoMark
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9420003,9420008,9410001"
	SSProcess.SelectFilter
	SSProcess.DeleteSelectionObj
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
'�ֲ�ͼͼ��
function CreateKEY
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420031
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	if geoCount>0 then
		for i = 0 to geoCount-1
			TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
			SSProcess.GetObjectPoint TKID, 1, x, y, z, pointtype, name
			ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"9420021,9420022,9420023,9420024", 0)
			ZGNQMC = ""
			ZDrawCode=""
			if ids<> "" then
				SSFunc.ScanString ids, ",", vArray, nCount
				FOR j=0 to nCount-1
					DrawCode=SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
					DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
					GNQMC=SSProcess.GetObjectAttr(vArray(j), "[MC]")
					if GNQMC<>"" then
						IF ZGNQMC="" THEN
							ZDrawCode = DrawCode
							ZGNQMC=GNQMC
							ZDrawColor = DrawColor
						ELSE
						  if replace(ZGNQMC,GNQMC,"")=ZGNQMC then
								ZGNQMC=ZGNQMC&","&GNQMC
								ZDrawCode = ZDrawCode&","&DrawCode
								ZDrawColor = ZDrawColor&","&DrawColor
							end if 
						END IF
					end if 
			  Next
			end if 
			LvDiTuLi x-16,y,ZGNQMC,TKID,ZDrawCode,ZDrawColor
		next
	end if
end function


'�ֲ�ͼͼ��
function LvDiTuLi(x0,y0,ZGNQMC,polygonID,ZDrawCode,ZDrawColor)
		wid1 = 228 : heig1 = 286
		wid2 = 228 : heig2 = 286
		SSFunc.ScanString ZGNQMC, ",", cvArray1, count5
		arDrawCode = split(ZDrawCode,",")
		arDrawColor = split(ZDrawColor,",")
       '����
         makeLine x0,y0,x0,y0+count5*2+2.5,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID

			makeLine x0+16,y0,x0+16,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
		 '����
			'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0,x0+16,y0,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0+count5*2+2.5,x0+16,y0+count5*2+2.5,1, "RGB(255,255,255)", polygonID
			makeNote x0+2.5,y0+count5*2+1.5 , 0, "RGB(255,255,255)", wid2, heig2, "ͼ��",polygonID
			makeNote x0+10,y0+count5*2+1.5 , 0, "RGB(255,255,255)", wid2, heig2, "��ע",polygonID

			 for j= 0 to count5-1
			 '����
               makeArea x0+1,y0+j*2+0.7,x0+7,y0+j*2+0.7,x0+7,y0+j*2+2.3,x0+1,y0+j*2+2.3,arDrawCode(j), arDrawColor(j), polygonID
			      makeLine x0,y0+j*2+2.5,x0+16,y0+j*2+2.5, 1,"RGB(255,255,255)", polygonID
					makeNote x0+10,y0+1.5+ j*2, 0, "RGB(255,255,255)", wid2, heig2, cvArray1(j),polygonID
			  next
end function 

'�����滮��ƽͼ,�滮����ͼ
function CreateKEYZPT
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9420034,9420035"
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	if geoCount>0 then
		for i = 0 to geoCount-1
			TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
			DaYBL = SSProcess.GetSelGeoValue( i, "[DaYBL]" )
			SSProcess.GetObjectPoint TKID, 1, x, y, z, pointtype, name
			ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"9410001,9410011,9410021,9410031,9410041,9410051,9410061,9410071,9410091,9410101,9410104,9410105", 0)
			if ids<> "" then
				SSFunc.ScanString ids, ",", vArray, nCount
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
			LvDiTuLiZPT x-16,y,TKID,ZDrawCode,ZDrawColor,ZDrawName,500
		next
	end if
	SSProcess.SetMapScale "500"
end function
'�����滮��ƽͼ
function LvDiTuLiZPT(x0,y0,polygonID,ZDrawCode,ZDrawColor,ZDrawName,DaYBL)
		wid1 = (228 *500)/DaYBL: heig1 = (286*500)/DaYBL
		wid2 = (228*500)/DaYBL : heig2 = (286*500)/DaYBL
		arDrawCode = split(ZDrawCode,",")
		arDrawColor = split(ZDrawColor,",")
		arDrawName = split(ZDrawName,",")
		count5 = ubound(arDrawCode)+2
       '����
         makeLine x0,y0,x0,y0+count5*2+2.5,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID

			makeLine x0+16,y0,x0+16,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
		 '����
			'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0,x0+16,y0,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0+count5*2+2.5,x0+16,y0+count5*2+2.5,1, "RGB(255,255,255)", polygonID
			makeNote x0+7,y0+count5*2+1 , 0, "RGB(255,255,255)", wid2, heig2, "ͼ��",polygonID

			for j= 0 to ubound(arDrawCode)
			 '����
               makeLine x0+1,y0+j*2+1.5,x0+7,y0+j*2+1.5,arDrawCode(j), arDrawColor(j), polygonID
			      'makeLine x0,y0+j*2+2.5,x0+16,y0+j*2+2.5, 1,"RGB(255,255,255)", polygonID
					makeNote x0+10,y0+1.5+ j*2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
			next
end function 

'����ͼ
function CreateKEYJD
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9420032
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	if geoCount>0 then
		for i = 0 to geoCount-1
			TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
			DaYBL=SSProcess.GetSelGeoValue( i, "[DaYBL]" )
			SSProcess.GetObjectPoint TKID, 1, x, y, z, pointtype, name
			ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"9420025,9420026,9420027", 0)
			if ids<> "" then
				SSFunc.ScanString ids, ",", vArray, nCount
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
			LvDiTuLiJD x-20,y,TKID,ZDrawCode,ZDrawColor,ZDrawName,DaYBL
		next
	end if
end function
'����ͼ
function LvDiTuLiJD(x0,y0,polygonID,ZDrawCode,ZDrawColor,ZDrawName,DaYBL)
		wid1 = 228*500/DaYBL : heig1 = 286*500/DaYBL 
		wid2 = 228*500/DaYBL  : heig2 = 286*500/DaYBL 
		arDrawCode = split(ZDrawCode,",")
		arDrawColor = split(ZDrawColor,",")
		arDrawName = split(ZDrawName,",")
		count5 = ubound(arDrawCode)+2
       '����
         makeLine x0,y0,x0,y0+count5*2+2.5,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID

			makeLine x0+20,y0,x0+20,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
		 '����
			'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0,x0+16,y0,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0+count5*2+2.5,x0+20,y0+count5*2+2.5,1, "RGB(255,255,255)", polygonID
			makeNote x0+8,y0+count5*2+1 , 0, "RGB(255,255,255)", wid2, heig2, "ͼ��",polygonID

			for j= 0 to ubound(arDrawCode)
			 '����
				makeNote x0+1,y0+1.5+ j*2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j)&":",polygonID
				makeArea x0+10,y0+j*2+0.7,x0+17,y0+j*2+0.7,x0+17,y0+j*2+2.3,x0+10,y0+j*2+2.3,arDrawCode(j), arDrawColor(j), polygonID
			next
end function

function makePoint(x,y,code,color,polygonID)
		SSProcess.CreateNewObj 0
		SSProcess.SetNewObjValue "SSObj_Code", code
		SSProcess.SetNewObjValue "SSObj_Color", color
		SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
		SSProcess.SetNewObjValue "SSObj_LayerName", "���������ɹ�ͼͼ����Ϣ"
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
		SSProcess.SetNewObjValue "SSObj_LayerName", "���������ɹ�ͼͼ����Ϣ"
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
		SSProcess.SetNewObjValue "SSObj_LayerName", "���������ɹ�ͼͼ����Ϣ"
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
		SSProcess.SetNewObjValue "SSObj_FontClass", "0"
		SSProcess.SetNewObjValue "SSObj_FontString", fontString
		SSProcess.SetNewObjValue "SSObj_Color", color
		SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
		SSProcess.SetNewObjValue "SSObj_LayerName", "���������ɹ�ͼͼ����Ϣ"
		SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
		SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
		SSProcess.SetNewObjValue "SSObj_FontWidth", width
		SSProcess.SetNewObjValue "SSObj_FontHeight", height
		SSProcess.AddNewObjPoint x, y, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 

dim adoConnection
Function InitDB() 
		accessName=SSProcess.GetProjectFileName
		set adoConnection=createobject("adodb.connection")
		strcon="DBQ="& accessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
		adoConnection.Open strcon
End Function

'//�ؿ�
Function ReleaseDB()
		adoConnection.Close
		Set adoConnection = Nothing
End Function
'//�жϱ��Ƿ����
Function IsTableExits(byval  strMdbName,byval  strTableName_s)
		strMdbName=SSProcess.GetProjectFileName
		IsTableExits=false 
		strTableName_s=ucase(strTableName_s)
		'�ж��ļ�DB�ļ��Ƿ����
		Set fso=CreateObject("Scripting.FileSystemObject")   
		if fso.fileExists(strMdbName)=false then exit function 
		'��ȡDB�ļ���׺��
		set f=fso.getfile(strMdbName)
		dbType= fso.GetExtensionName(f)
		set f = nothing 
		set fso = nothing 
		'��DB���Ͳ���
		if dbType="dbf" then 
			strMdbName=Replace(strMdbName,"/","\")
			ipos=InStrRev(strMdbName,"\")
			strMdbName=Left(strMdbName,ipos)
			strcon="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMdbName + ";Extended Properties=dBASE IV;User ID=;Password="
		else
			strcon="DBQ="& strMdbName &";DRIVER={Microsoft Access Driver (*.mdb)};"
		end if 
		'����
		'''set adoConnection=createobject("adodb.connection")
		'''adoConnection.Open strcon
		Set rsSchema = adoConnection.OpenSchema(20)
		'��ȡ��ǰDB�ļ����б�
		strAllTableName=""
		Do While Not rsSchema.EOF
				strTableName =""&UCASE (rsSchema.Fields("TABLE_NAME"))&""
				if left(strTableName,4) <> "MSYS" then
					if strTableName_s=strTableName then IsTableExits=true:exit do
				end if 
				rsSchema.MoveNext
		Loop
		rsSchema.Close
		Set rsSchema = Nothing
		if IsTableExits=false then addloginfo "��"&strTableName_s&"������edb�в�����"
		''adoConnection.Close
		''Set adoConnection = Nothing
End Function	

Function GetProjectTableList(byval strTableName,byval strFields,byval strAddCondition,byval strTableType,byval strGeoType,byref rs(),byref fieldCount)
		GetProjectTableList=0:   values="":rsCount = 0:fieldCount=0
		if strTableName="" or strFields="" then Exit function
		if IsTableExits("",strTableName)=false then Exit Function
		'strFields=GetTableAllFields ("", strTableName, strFields)
		if  strFields="" then Exit function
		'���õ�������
		if strGeoType="0" then 
			GeoType="GeoPointTB"
		elseif strGeoType="1" then
			GeoType="GeoLineTB"
		elseif strGeoType="2" then 
			GeoType="GeoAreaTB"
		elseif strGeoType="3" then 
			GeoType="MarkNoteTB"
		else
			GeoType="GeoAreaTB"
		end if 
		if strTableType="SpatialData" then 
			strCondition=" ("&GeoType&".Mark Mod 2)<>0"
			if strAddCondition<>"" then strCondition=" ("&GeoType&".Mark Mod 2)<>0 and "&strAddCondition&""	
			sql = "select  "&strFields&" from "&strTableName&"  INNER JOIN "&GeoType&" ON "&strTableName&".ID = "&GeoType&".ID WHERE "&strCondition&""
		else 
			if strAddCondition<>"" then 	 
				strCondition=strAddCondition
				sql = "select  "&strFields&" from "&strTableName&"  WHERE  "&strCondition&""
			else 
				sql = "select  "&strFields&" from "&strTableName&""
			end if 
		end if

		''addloginfo sql
		'if instr(sql,"scpcjzmj")>0 then  addloginfo sql
		'��ȡ��ǰ����edb���¼
		AccessName=SSProcess.GetProjectFileName
		'�жϱ��Ƿ����
		'if  IsTableExits(AccessName,strTableName)=false then exit function 
		'set adoConnection=createobject("adodb.connection")
		'strcon="DBQ="& AccessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
		'adoConnection.Open strcon
		Set adoRs=CreateObject("ADODB.recordset")
		count=0
		adoRs.cursorLocation =3 
		adoRs.cursorType =3
		adoRs.open sql,adoConnection,3,3
		rcdCount = adoRs.RecordCount
		fieldCount= adoRs.Fields.Count
		redim rs(rcdCount,fieldCount)
		'erase rs
		while adoRs.Eof=false
				nowValues=""
				For i=0 To fieldCount-1
						value=adoRs(i)
						if isnull(value) then value=""
						value=replace(value,",","��")
						rs(rsCount,i)=value
				Next
				rsCount=rsCount+1
				adoRs.MoveNext
		wend
		adoRs.Close
		Set adoRs = Nothing
		'adoConnection.Close
		'Set adoConnection = Nothing
		GetProjectTableList=rsCount
End Function
function YDHXYDMJ(YongDMJ)
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9410001"
	SSProcess.SelectFilter
	geocount = SSProcess.GetSelGeoCount()
	if geocount= 1 then
		for i =  0 to geocount-1
			id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
			YongDMJ = SSProcess.GetObjectAttr (id, "SSObj_Area")
		next
	end if
	YDHXYDMJ = YongDMJ
end function

Function GetSubArea(cellList,cellCount,byval scArea,byval ghArea,byval numberDigit,byval startCol)
		if isnumeric(numberDigit)=false then numberDigit=2
		scArea=GetFormatNumber(scArea,numberDigit)
		ghArea=GetFormatNumber(ghArea,numberDigit)
		subNum=cdbl(scArea)-cdbl(ghArea)
		subNum=GetFormatNumber(subNum,numberDigit)'��ֵ-�����������
		if scArea = "0.00" or scArea = "0" then scArea = ""
		if ghArea = "0.00" or ghArea = "0" then ghArea = ""
		if subNum = "0.00" or subNum = "0" then subNum = ""

		if startCol=2 then  	cellValue=scArea&"||"&ghArea&"||"&subNum &"||"&""  else 	cellValue=scArea&"||"&ghArea&"||"&subNum
		redim preserve cellList(cellCount): cellList(cellCount)=cellValue:cellCount=cellCount+1
End Function
Function OutputTable11( )
		cellCount=0:redim cellList(cellCount)
		'**************************************************************���õ����
		ydhxTableName="JG_�õغ�����Ϣ���Ա�"
		fields="GuiHSPZYDMJ"
		listCount=GetProjectTableList (ydhxTableName,"GuiHSPZYDMJ","","SpatialData","1",list,fieldCount)
		if listCount=1 then gh_YongDMJ=list(0,0)
		gh_YongDMJ=GetFormatNumber(gh_YongDMJ,2)'�滮-���õ����
		sc_YongDMJ = YDHXYDMJ(YongDMJ)
		if sc_YongDMJ<>"" then sc_YongDMJ = GetFormatNumber(sc_YongDMJ,2)
		GetSubArea cellList,cellCount, sc_YongDMJ, gh_YongDMJ,2,1
		
		'**************************************************************�ܽ������
		zrzCount=GetProjectTableList ("FC_��Ȼ����Ϣ���Ա�","sum(SCJZMJ)","","SpatialData","2",zrzList,fieldCount)
		if zrzCount=1 then sc_SCJZMJ=zrzList(0,0)
		sc_SCJZMJ=GetFormatNumber(sc_SCJZMJ,2)'ʵ��-�ܽ������
		
		ghxkTableName="JG_���蹤�̹滮���֤��Ϣ���Ա�"
		'exCondition="YDHXGUID In (select YDHXGUID from "&ydhxTableName&"  INNER JOIN GeoLineTB ON "&ydhxTableName&".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0)"

		exCondition="ID>0"
		ghxkCount=GetProjectTableList (ghxkTableName,"sum(GuiHSPZJZMJ)",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then gh_SCJZMJ=ghxkList(0,0)
		gh_SCJZMJ=GetFormatNumber(gh_SCJZMJ,2)'�滮-�ܽ������
		GetSubArea cellList,cellCount, sc_SCJZMJ, gh_SCJZMJ,2,1
		
		iniRow=1:iniCol=1
		startRow=iniRow:startCol=iniCol
		ghgnqTableName="JG_�滮���������Ա�"
		'**************************************************************���Ͻ������
		GetGnqAreaList cellList,cellCount, ghxkTableName, ghgnqTableName, "int(sjcs)>0","GuiHSPDSJZMJ",exCondition,copyCount,sumDsGNQMJ
		'���Ƶ��Ϲ�����
		startRow=iniRow+2
		startCol=iniCol+1
	
		'**************************************************************���½������
		startRow=startRow+copyCount+1
		GetGnqAreaList cellList,cellCount, ghxkTableName, ghgnqTableName, "int(sjcs)<0","GuiHSPDXJZMJ",exCondition,copyCount1,sumDxGNQMJ
		'���Ƶ��¹�����

		
		
		'**************************************************************�����������
		jdCount=GetProjectTableList ("JG_��������������Ա�","sum(JDMJ)","","SpatialData","2",jdList,fieldCount)
		if jdCount=1 then sc_JDMJ=jdList(0,0)
		sc_JDMJ=GetFormatNumber(sc_JDMJ,2)'ʵ��-�����������
		ghxkCount=GetProjectTableList (ghxkTableName,"sum(GuiHSPJDMJ),sum(GuiHSPRJL),sum(GuiHSPJZMD),sum(GuiHSPLHL),sum(ZpsJZMJ),sum(ScZZHS),sum(GhZZHS)",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then 
			gh_JDMJ=ghxkList(0,0):gh_GuiHSPRJL=ghxkList(0,1):gh_GuiHSPJZMD=ghxkList(0,2)
			gh_GuiHSPLHL=ghxkList(0,3):gh_ZpsJZMJ=ghxkList(0,4)
			ScZZHS=ghxkList(0,5):GhZZHS=ghxkList(0,6)
		end if 
		gh_JDMJ=GetFormatNumber(gh_JDMJ,2)'�滮-�����������
		GetSubArea cellList,cellCount, sc_JDMJ, gh_JDMJ,2,1

		ldCount=GetProjectTableList ("GH_�̻�Ҫ�����Ա�","sum(LHMJ)","ID>0","","",sclhmjList,fieldCount)
		if ldCount = 1 then sc_lhmj=sclhmjList(0,0)
		gh_lhmj = ""
		GetSubArea cellList,cellCount, sc_lhmj, gh_lhmj,2,1'�̵����
		
		if  sc_YongDMJ=0 then sc_Rjl=0 else    sc_Rjl=sumDsGNQMJ/sc_YongDMJ
		GetSubArea cellList,cellCount, sc_Rjl, gh_GuiHSPRJL,2,1'�ݻ���
		
		if  sc_YongDMJ=0 then sc_Jzmd=0 else    sc_Jzmd=(sc_JDMJ/sc_YongDMJ)*100
		GetSubArea cellList,cellCount, sc_Jzmd, gh_GuiHSPJZMD,2,1'�����ܶ�
		
		ldCount=GetProjectTableList ("GH_�̻�Ҫ�����Ա�","sum(LHMJ/ZSBL)","ID>0","","",sclhYdmjList,fieldCount)
		if ldCount = 1 then sc_lhYdmj=sclhYdmjList(0,0)
		if  sc_YongDMJ=0 then sc_lhl=0 else    sc_lhl=(sc_lhYdmj/sc_YongDMJ)*100
		GetSubArea cellList,cellCount, sc_lhl, gh_GuiHSPLHL,2,1'�̵���
		
		sc_ZpsJZMJ=""
		if gh_ZpsJZMJ = 0 then gh_ZpsJZMJ=""
		GetSubArea cellList,cellCount, sc_ZpsJZMJ, gh_ZpsJZMJ,2,1'װ��ʽ�������
		
		cwTableName="CWSCXX"
		cwCount=GetProjectTableList (cwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX='��ͨ������λ'","","",cwList,fieldCount)
		if  cwCount=1 then    sc_Jdcw=cwList(0,0):sc_ds_Jdcw=cwList(0,1):sc_dx_Jdcw=cwList(0,2)
		
		ghcwTableName="CWGHXX"
		cwCount=GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX='��ͨ������λ'","","",ghcwList,fieldCount)
		if  cwCount=1 then    gh_Jdcw=ghcwList(0,0):gh_ds_Jdcw=ghcwList(0,1):gh_dx_Jdcw=ghcwList(0,2)
		
		GetSubArea cellList,cellCount, sc_Jdcw, gh_Jdcw,0,1'������λ
		GetSubArea cellList,cellCount, sc_ds_Jdcw, gh_ds_Jdcw,0,2'���ϻ�����λ
		GetSubArea cellList,cellCount, sc_dx_Jdcw, gh_dx_Jdcw,0,2'���»�����λ
		GetSubArea cellList,cellCount, ScZZHS, GhZZHS,0,1'סլ����
		
		cwCount=GetProjectTableList (cwTableName,"sum(DSCWSL)+sum(DXCWSL)","CWLX='�ǻ�����λ'","","",cwList,fieldCount)
		if  cwCount=1 then    sc_Fjdcw=cwList(0,0)
		ghcwCount=GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL)","CWLX='�ǻ�����λ'","","",ghcwList,fieldCount)
		if  ghcwCount=1 then    gh_Fjdcw=ghcwList(0,0)
		GetSubArea cellList,cellCount, sc_Fjdcw, gh_Fjdcw,0,1'�ǻ�����λ
		
	
End Function

'//��ȡ�������������
Function GetGnqAreaList(cellList,cellCount,byval ghxkTableName,byval ghgnqTableName,byval strConditon,byval field,exCondition,copyCount,sc_GNQMJ)
		copyCount=0
		'**************************************************************�������
		ghgnqCount=GetProjectTableList (ghgnqTableName,"SUM(GNQMJ)",strConditon,"SpatialData","2",ghgnqList,fieldCount)
		if ghgnqCount=1  then sc_GNQMJ=ghgnqList(0,0)
		sc_GNQMJ=GetFormatNumber(sc_GNQMJ,2)'ʵ��-�������
		
		ghxkCount=GetProjectTableList (ghxkTableName,"sum("&field&")",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then gh_GNQMJ=ghxkList(0,0)
		gh_GNQMJ=GetFormatNumber(gh_GNQMJ,2)'�滮-�������
		GetSubArea cellList,cellCount, sc_GNQMJ, gh_GNQMJ,2,2
		'**************************************************************�������-�����������
		ghgnqCount=GetProjectTableList (ghgnqTableName,"SUM(JZMJ),YT",""&strConditon&" group by YT","SpatialData","2",ghgnqList,fieldCount)

		ghldxxCount = GetProjectTableList ("GHLDXX","SUM(JZMJ),GHYT","GHYT<>'' group by GHYT","AttributeData","0",ghldxxList,ghldxxfieldCount)

		if ghgnqCount>0 then
			for i=0 to ghgnqCount-1
				sc_gnq_GNQMJ=ghgnqList(i,0):sc_gnq_GNQMJ=GetFormatNumber(sc_gnq_GNQMJ,2)
				gnqName=ghgnqList(i,1)
				ghldxx_gnqmj=""
				if ghldxxCount> 0 then
					for i1 = 0 to ghldxxCount-1
						if ghldxxList(i1,1) = gnqName then	ghldxx_gnqmj=ghldxxList(i1,0):ghldxx_gnqmj=GetFormatNumber(ghldxx_gnqmj,2)
					next
				end if
				if sc_gnq_GNQMJ="" then sc_gnq_GNQMJ=0
				if ghldxx_gnqmj="" then ghldxx_gnqmj=0
				change_gnqmj = GetFormatNumber(sc_gnq_GNQMJ-ghldxx_gnqmj,2)
				if sc_gnq_GNQMJ = "0.00" or sc_gnq_GNQMJ = "0" then sc_gnq_GNQMJ = ""
				if ghldxx_gnqmj = "0.00" or ghldxx_gnqmj = "0" then ghldxx_gnqmj = ""
				if change_gnqmj = "0.00" or change_gnqmj = "0" then change_gnqmj = ""
				cellValue=gnqName&"||"&sc_gnq_GNQMJ&"||"&ghldxx_gnqmj&"||"&change_gnqmj
				redim preserve cellList(cellCount): cellList(cellCount)=cellValue:cellCount=cellCount+1
				copyCount=copyCount+1
			next
		else
				cellValue=gnqName&"||"&""&"||"&""&"||"&""
				redim preserve cellList(cellCount): cellList(cellCount)=cellValue:cellCount=cellCount+1
		end if
End Function

function JGZPTKEY(byval TKID)

InitDB() 
	sSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()

	xmmc=SSProcess.GetSelGeoValue( 0, "[XiangMMC]" )
'��JG_�滮���������Ա��yt �ֶ� ȥ�� �õ�����a
		ydhxTableName="JG_�滮���������Ա�"
		GNQCount=GetProjectTableList (ydhxTableName,"DISTINCT(YT)","","SpatialData","2",list,fieldCount)

		'���������� ����
'������= a+13
		ZHS=GNQCount+12
'�����
'��ȡͼ������λ��

	SSProcess.GetObjectPoint TKID, 2, x, y, z, pointtype, name
	x1= x-10
	y1=y-10
	ztmc="����"
	ztdx=150
	ztkd=200
	yzl=y1-7-10-GNQCount*2-24-1

'��ͷ
	makeNote1 x1-10, y1+2, code, color, ztdx, ztkd, xmmc&"��Ŀ�滮��ʵ��",TKID,ztmc
	makeArea x1+3,y1,x1-17,y1,x1-17,yzl+1,x1+3,yzl+1,1,color,TKID 

'��һ������

	makeLine x1-15,y1,x1-15,yzl+1,1, color, TKID
'��һ������
	makeLine x1-17,y1-3,x1+3,y1-3,1, color, TKID
'���
	makeNote1 x1-16.3, y1-0.5, code, color, ztdx, ztkd, "��",TKID,ztmc
	makeNote1 x1-16.3, y1-1.5, code, color, ztdx, ztkd, "��",TKID,ztmc
'�ڶ�������
	makeLine x1-3,y1,x1-3,yzl+10,1, color, TKID
	makeNote1 x1-14, y1-1.5, code, color, ztdx, ztkd, "�õ���Ϣ����ָ������",TKID,ztmc
	makeNote1 x1-1.5, y1-1.5, code, color, ztdx, ztkd, "��ʵ����",TKID,ztmc

'�ڶ�������
	makeLine x1-17,y1-5,x1+3,y1-5,1, color, TKID
	makeNote1 x1-16, y1-3.5, code, color, ztdx, ztkd, "1",TKID,ztmc
	makeNote1 x1-10, y1-3.5, code, color, ztdx, ztkd, "�ݻ���",TKID,ztmc
	ghgnqCount=GetProjectTableList ("JG_�滮���������Ա�","SUM(JZMJ)","int(sjcs)>0","SpatialData","2",ghgnqList,fieldCount)
	if ghgnqCount=1  then sumDsGNQMJ=ghgnqList(0,0)
	if sumDsGNQMJ="" then sumDsGNQMJ=0
	sumDsGNQMJ=GetFormatNumber(sumDsGNQMJ,2)'ʵ��-�������
	sc_YongDMJ = YDHXYDMJ(YongDMJ)
	if sc_YongDMJ<>"" then sc_YongDMJ = GetFormatNumber(sc_YongDMJ,2)
	if  sc_YongDMJ=0 then sc_Rjl=0 else    sc_Rjl=GetFormatNumber(sumDsGNQMJ/sc_YongDMJ,2)
	makeNote1 x1-1.5, y1-4, code, color, ztdx, ztkd, sc_Rjl,TKID,ztmc
'����������
	makeLine x1-17,y1-7,x1+3,y1-7,1, color, TKID
	makeNote1 x1-16, y1-6, code, color, ztdx, ztkd, "2",TKID,ztmc
	makeNote1 x1-12, y1-5.5, code, color, ztdx, ztkd, "�����ݻ��ʽ������",TKID,ztmc
	makeNote1 x1-1.5, y1-5.5, code, color, ztdx, ztkd, sumDsGNQMJ,TKID,ztmc
'���ĸ�����
	makeLine x1-15,y1-18,x1+3,y1-18,1, color, TKID
	makeNote1 x1-16, y1-12-GNQCount, code, color, ztdx, ztkd, "3",TKID,ztmc

'��1���̺���
	makeLine x1-15,y1-9,x1+3,y1-9,1, color, TKID
	makeNote1 x1-10.5, y1-7.5, code, color, ztdx, ztkd, "�ܽ������",TKID,ztmc
	zrzCount=GetProjectTableList ("FC_��Ȼ����Ϣ���Ա�","sum(JZMJ)","","SpatialData","2",zrzList,fieldCount)
	if zrzCount=1 then sc_SCJZMJ=zrzList(0,0)
	sc_SCJZMJ=GetFormatNumber(sc_SCJZMJ,2)'ʵ��-�ܽ������
	makeNote1 x1-1.5, y1-7.5, code, color, ztdx, ztkd, sc_SCJZMJ,TKID,ztmc

'��2���̺���
	makeLine x1-9,y1-13,x1+3,y1-13,1, color, TKID
	makeNote1 x1-8, y1-10, code, color, ztdx, ztkd, "���Ͻ���",TKID,ztmc
	makeNote1 x1-7, y1-12, code, color, ztdx, ztkd, "���",TKID,ztmc
	makeNote1 x1-1.5, y1-11, code, color, ztdx, ztkd, sumDsGNQMJ,TKID,ztmc

	makeNote1 x1-8, y1-14, code, color, ztdx, ztkd, "���½���",TKID,ztmc
	makeNote1 x1-7, y1-16, code, color, ztdx, ztkd, "���",TKID,ztmc
	ghgnqCount1=GetProjectTableList ("JG_�滮���������Ա�","SUM(jzmj)","int(sjcs)<0","SpatialData","2",ghgnqList1,fieldCount1)
	if ghgnqCount1=1  then sumDsGNQMJ1=ghgnqList1(0,0)
	sumDsGNQMJ1=GetFormatNumber(sumDsGNQMJ1,2)'ʵ��-�������
	makeNote1 x1-1.5, y1-15, code, color, ztdx, ztkd, sumDsGNQMJ1,TKID,ztmc
	makeNote1 x1-14, y1-12, code, color, ztdx, ztkd, "���ռ�",TKID,ztmc
	makeNote1 x1-14, y1-14, code, color, ztdx, ztkd, "λ�÷���",TKID,ztmc

	makeNote1 x1-13, y1-17-GNQCount-2, code, color, ztdx, ztkd, "��ʹ��",TKID,ztmc
	makeNote1 x1-13, y1-17-GNQCount-4, code, color, ztdx, ztkd, "��;����",TKID,ztmc
'ѭ��������;
	for j= 0 to GNQCount-1 
		ytname=list(j,0)
		makeLine x1-9,y1-17-j*2-3,x1+3,y1-17-j*2-3,1, color, TKID
		makeNote1 x1-8, y1-17-j*2-1.5, code, color, ztdx, ztkd,ytname ,TKID,ztmc
		'���Ӧ���
		ytCount=GetProjectTableList (ydhxTableName,"sum(jzmj)","yt='"&ytname&"'","SpatialData","2",list1,fieldCount1)
		ytmj=list1(0,0)
		makeNote1 x1-1.5, y1-17-j*2-1.5, code, color, ztdx, ztkd,ytmj,TKID,ztmc
	next
	'����
	makeLine x1-9,y1-9,x1-9,y1-7-10-GNQCount*2-1,1, color, TKID
'���������
	makeLine x1-15,y1-7-10-GNQCount*2-1,x1+3,y1-7-10-GNQCount*2-1,1, color, TKID
'����������
	makeLine x1-17,y1-7-10-GNQCount*2-5,x1+3,y1-7-10-GNQCount*2-5,1, color, TKID
	makeNote1 x1-16, y1-7-10-GNQCount*2-3, code, color, ztdx, ztkd, "4",TKID,ztmc
'�̺���
	makeLine x1-15,y1-7-10-GNQCount*2-3,x1+3,y1-7-10-GNQCount*2-3,1, color, TKID
	makeNote1 x1-10, y1-7-10-GNQCount*2-1.5, code, color, ztdx, ztkd, "���ϳ�λ",TKID,ztmc
	cwTableName="CWSCXX"
	cwCount=GetProjectTableList (cwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX='��ͨ������λ'","","",cwList,fieldCount)
	if  cwCount=1 then    sc_Jdcw=cwList(0,0):sc_ds_Jdcw=cwList(0,1):sc_dx_Jdcw=cwList(0,2)
	cwCount=GetProjectTableList (cwTableName,"sum(DSCWSL)","CWLX='�ǻ�����λ'","","",cwList,fieldCount)
	if  cwCount=1 then    sc_dsFjdcw=cwList(0,0)
	cwCount1=GetProjectTableList (cwTableName,"sum(DXCWSL)","CWLX='�ǻ�����λ'","","",cwList1,fieldCount1)
	if  cwCount1=1 then    sc_dXFjdcw=cwList1(0,0)
	if sc_ds_Jdcw<>"" then sc_ds_Jdcw=cdbl(sc_ds_Jdcw)
	if sc_dsFjdcw<>"" then sc_dsFjdcw=cdbl(sc_dsFjdcw)
	if sc_dx_Jdcw<>"" then sc_dx_Jdcw=cdbl(sc_dx_Jdcw)
	if sc_dXFjdcw<>"" then sc_dXFjdcw=cdbl(sc_dXFjdcw)
	dscezsl=sc_ds_Jdcw+sc_dsFjdcw
	dxcezsl=sc_dx_Jdcw+sc_dXFjdcw
	makeNote1 x1-1.5, y1-7-10-GNQCount*2-1.5, code, color, ztdx, ztkd, dscezsl,TKID,ztmc
	makeNote1 x1-10, y1-7-10-GNQCount*2-3.5, code, color, ztdx, ztkd, "���³�λ",TKID,ztmc
	makeNote1 x1-1.5, y1-7-10-GNQCount*2-3.5, code, color, ztdx, ztkd,dxcezsl,TKID,ztmc
'���߸�����
	makeLine x1-17,y1-7-10-GNQCount*2-7,x1+3,y1-7-10-GNQCount*2-7,1, color, TKID
	makeNote1 x1-16, y1-7-10-GNQCount*2-6, code, color, ztdx, ztkd, "5",TKID,ztmc
	makeNote1 x1-10, y1-7-10-GNQCount*2-5.5, code, color, ztdx, ztkd, "�̵���",TKID,ztmc

	ldCount=GetProjectTableList ("GH_�̻�Ҫ�����Ա�","sum(LHMJ/ZSBL)","ID>0","","",sclhYdmjList,fieldCount)
	if ldCount = 1 then sc_lhYdmj=sclhYdmjList(0,0)
	if sc_lhYdmj="" then sc_lhYdmj=0
	if  sc_YongDMJ=0 then sc_lhl=0 else    sc_lhl=(sc_lhYdmj/sc_YongDMJ)*100
	sc_lhl=GetFormatNumber(sc_lhl,2)
	makeNote1 x1-1.5, y1-7-10-GNQCount*2-5.5, code, color, ztdx, ztkd, sc_lhl&"%",TKID,ztmc
'�ڰ˸�����
	makeLine x1-15,y1-7-10-GNQCount*2-9,x1+3,y1-7-10-GNQCount*2-9,1, color, TKID
	makeNote1 x1-16, y1-7-10-GNQCount*2-11, code, color, ztdx, ztkd, "6",TKID,ztmc
	makeNote1 x1-10, y1-7-10-GNQCount*2-7.5, code, color, ztdx, ztkd, "���ؿ������",TKID,ztmc
	'������makeNote x1, y1-7-10-GNQCount*2-7, code, color, ztdx, ztkd, "���ؿ����������",TKID,ztmc
'�ھŸ�����
	makeLine x1-17,y1-7-10-GNQCount*2-15,x1+3,y1-7-10-GNQCount*2-15,1, color, TKID
'����֮������
	makeLine x1-7,y1-7-10-GNQCount*2-9,x1-7,yzl+1,1, color, TKID
	makeNote1 x1-13, y1-7-10-GNQCount*2-11.5, code, color, ztdx, ztkd, "���ڵط���",TKID,ztmc
'�̺���
	makeLine x1-7,y1-7-10-GNQCount*2-11,x1+3,y1-7-10-GNQCount*2-11,1, color, TKID
	makeNote1 x1-6, y1-7-10-GNQCount*2-9.5, code, color, ztdx, ztkd, "�ڵ�һ",TKID,ztmc
	'�ڵ�������ȡֵʱ ����ſ�
	'makeNote x1, y1-7-10-GNQCount*2-9, code, color, ztdx, ztkd, "�ڵ�һ����",TKID,ztmc
'�̺���
	makeLine x1-7,y1-7-10-GNQCount*2-13,x1+3,y1-7-10-GNQCount*2-13,1, color, TKID
	makeNote1 x1-6, y1-7-10-GNQCount*2-11.5, code, color, ztdx, ztkd, "�ڵض�",TKID,ztmc
	'�ڵ�������ȡֵʱ ����ſ�
	'makeNote x1, y1-7-10-GNQCount*2-11, code, color, ztdx, ztkd, "�ڵض�����",TKID,ztmc
'�̺���
	makeNote1 x1-6, y1-7-10-GNQCount*2-13.5, code, color, ztdx, ztkd, "�ڵ���",TKID,ztmc
	'�ڵ�������ȡֵʱ ����ſ�
	'makeNote x1, y1-7-10-GNQCount*2-13, code, color, ztdx, ztkd, "�ڵض�����",TKID,ztmc
'��ʮ������
	makeLine x1-17,y1-7-10-GNQCount*2-19,x1+3,y1-7-10-GNQCount*2-19,1, color, TKID
	makeNote1 x1-16, y1-7-10-GNQCount*2-17, code, color, ztdx, ztkd, "7",TKID,ztmc
	makeNote1 x1-13, y1-7-10-GNQCount*2-15.5, code, color, ztdx, ztkd, "������Ȩ֤��",TKID,ztmc
	makeNote1 x1-13.5, y1-7-10-GNQCount*2-17.5, code, color, ztdx, ztkd, "������֤��֤��",TKID,ztmc
	''������makeNote x1,  y1-7-10-GNQCount*2-16, code, color, ztdx, ztkd, "����",TKID,ztmc
'���һ��
	makeNote1 x1-16, y1-7-10-GNQCount*2-21, code, color, ztdx, ztkd, "8",TKID,ztmc
	makeNote1 x1-13, y1-7-10-GNQCount*2-21, code, color, ztdx, ztkd, "������;",TKID,ztmc
''������makeNote x1,  y1-7-10-GNQCount*2-19, code, color, ztdx, ztkd, "����",TKID,ztmc
'����
	makeNote1 x1-17, y1-7-10-GNQCount*2-27, code, color, ztdx, ztkd, "˵����",TKID,ztmc
	makeNote1 x1-17, y1-7-10-GNQCount*2-29, code, color, ztdx, ztkd, "1������������ָ���е�����ˮ��������ˮ����ʩ��",TKID,ztmc
	makeNote1 x1-17, y1-7-10-GNQCount*2-31, code, color, ztdx, ztkd, "2����Ŀ����������ָΪ����Ŀ���׵ĵ��������š�����ˮ���豸���÷���",TKID,ztmc
   makeNote1 x1-17, y1-7-10-GNQCount*2-33, code, color, ztdx, ztkd, "3���ո�Ϊ������Ŀ�ľ���������־�������д�����ݡ�",TKID,ztmc
	makeNote1 x1-17, y1-7-10-GNQCount*2-35, code, color, ztdx, ztkd, "4����ʹ����;�ķ�����ռ�λ�����ʩ��ƽ��ͼ��",TKID,ztmc
	makeNote1 x1-17, y1-7-10-GNQCount*2-37, code, color, ztdx, ztkd, "5����������ǰ��ա��������̽����������Ϳ����ۺϲ���������̡�",TKID,ztmc
	makeNote1 x1-17, y1-7-10-GNQCount*2-39, code, color, ztdx, ztkd, "��DB33/T 1152-2018�����к�ʵ",TKID,ztmc
	makeNote1 x1-17, y1-7-10-GNQCount*2-41, code, color, ztdx, ztkd, "6�������������ա��οա��в��δ��ʾ���������ƽ��ͼ����һ��",TKID,ztmc

'���ڿ�

ReleaseDB()
end function

function makeNote1(x, y, code, color, width, height, fontString,polygonID,ztmc)
  SSProcess.CreateNewObj 3
  SSProcess.SetNewObjValue "SSObj_FontClass", "0"
SSProcess.SetNewObjValue "SSObj_FontInterval", "80"
  SSProcess.SetNewObjValue "SSObj_FontString", fontString
  SSProcess.SetNewObjValue "SSObj_Color", color
  SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
	SSProcess.SetNewObjValue "SSObj_FontName", ztmc
  SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
  SSProcess.SetNewObjValue "SSObj_FontInterval", "8"
 SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
  SSProcess.SetNewObjValue "SSObj_FontAlignment", "1"
  SSProcess.SetNewObjValue "SSObj_FontWidth",width
  SSProcess.SetNewObjValue "SSObj_FontHeight", height
  SSProcess.AddNewObjPoint x, y, 0, 0, ""
  SSProcess.AddNewObjToSaveObjList
  SSProcess.SaveBufferObjToDatabase
end function 

Function GetFormatNumber(byval number,byval numberDigit)
		if isnumeric(numberDigit)=false then numberDigit=2
		if isnumeric(number)=false then number=0 
		number= formatnumber(round(number+0.00000001,numberDigit),numberDigit,-1,0,0)
		GetFormatNumber=(number)
End Function


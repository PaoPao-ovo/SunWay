Sub OnClick()
'��Ӵ���
	geoMaxID = SSProcess.GetGeoMaxID  
	SSProcess.ClearFunctionParameter 
	'���ҵ㴦���޾�
	SSProcess.AddFunctionParameter "limitdist=0.001"
	'���˻��α���
	SSProcess.AddFunctionParameter "SrcArcCodes=9420023,9420021,9420024"
	'ɾ��Դ����
	SSProcess.AddFunctionParameter "DelSrcArc=0"
	'ɾ���ϴ����ɵ��ص�����
	SSProcess.AddFunctionParameter "DelNewArc=0"
	'ɾ���ϴ����ɵ�ԭ������
	SSProcess.AddFunctionParameter "DelOldTopArea=0"
	'���ݴ�����Ƿ����
	SSProcess.AddFunctionParameter "SaveDB=1"
	'�Ƿ�����������
	SSProcess.AddFunctionParameter "CreateTopArea=1"
	'�������������  ���Ե����1,�����1,ͼ������1/���Ե����2,�����2,ͼ������2
	SSProcess.AddFunctionParameter "NewObject=-1,9420023,���������������Ϣ"
	'�ж����Ե��ظ��Ĺؼ���
	SSProcess.AddFunctionParameter "LabelKeyFields="
	'�������˻���ѡ��
	'0 �����ɻ���
	'1 ����ͳһ���뻡�Σ�������UniqueArcCodeָ��
	'2 ���ɻ���, ���ж�����״�����ص�ʱ���� ReserveArcOrder���õı���˳�����ȴ�ǰѡȡ
	'3 �Զ����������������ص����»���, ��CreateOverlayArc����
	SSProcess.AddFunctionParameter "CreateTopArc=0"
	SSProcess.TopProcess "��������˹���"
	
	SSProcess.ClearSelection  
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition  "SSObj_Type","=","AREA"
	SSProcess.SetSelectCondition  "SSObj_ID",">",geoMaxID
	SSProcess.SetSelectCondition  "SSObj_LayerName","=","���������������Ϣ"
	SSProcess.SelectFilter          
	geoCount =  SSProcess.GetSelGeoCount 
	
	SSProcess.ClearSelection  
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition  "SSObj_Type","=","AREA"
	SSProcess.SetSelectCondition  "SSObj_ID",">",geoMaxID
	SSProcess.SetSelectCondition  "SSObj_LayerName","=","���������������Ϣ"
	SSProcess.SelectFilter          
	geoCount =   SSProcess.GetSelGeoCount  
	innerObjGetPointMode = 1 '�жϽ���
	Dim idList(1000), idCount
	OpenBar "����ƥ�����������", geoCount
	for i=0 to geoCount-1
			 If (i mod 10) = 0 Then  RollBar "����ƥ�����������", CStr(geoCount-i)
			 geoID = SSProcess.GetSelGeoValue (i, "SSObj_ID") 
			 ids = SSProcess.SearchInnerObjIDs (geoID, 2, "9420023,9420021,9420024", innerObjGetPointMode)
			 If ids<>"" Then
					ScanString ids, ",", idList, idCount
					'ͬ�����־��λ
					posXY = SSRETools.GetAreaLabelPos (idList(0))
					SSRETools.SetAreaLabelPos geoID,posXY 
					'ͬ������
					SSProcess.CopyObjectAttr idList(0), geoID, 0, 1

					'ɾ��ԭ��
					SSProcess.DeleteObject idList(0) 
					'��ȡ����
					MJKMC=SSProcess.GetSelGeoValue (i, "[MianJKMC]") 
					MJKID=SSProcess.GetSelGeoValue (i, "SSObj_ID") 
					SYGN=SSProcess.GetSelGeoValue (i, "[YT]") 
					
					if  SYGN = "סլ"  Then col = RGB(255,0,0)
					if  SYGN = "��ҵ��ͨ�ִ�"  Then col = RGB(255,255,0)
					if  SYGN = "��ҵ������Ϣ"  Then col = RGB(0,255,0)
					if  SYGN = "����"  Then col = RGB(255,255,0)
					if  SYGN = "����ҽ����������"  Then col = RGB(0,255,255)
					if  SYGN = "�Ļ���������"  Then col = RGB(0,0,255)
					if  SYGN = "�칫"  Then col = RGB(255,0,255)
					if  SYGN = "����"  Then col = RGB(128,128,128)
					if  SYGN = "δ���������"  Then col = RGB(255,255,255)
					if  SYGN = "����"  Then col = RGB(192,192,192)
					SSProcess.SetObjectAttr MJKID,"SSObj_Color",col
					
			  
			End If
		Next
		CloseBar
End Sub


'����ϵͳ������
Function OpenBar(byval barname, range)
   SSProcess.EpsProgressCreate range,barname
   SSProcess.EpsProgressSetStep  1
End Function
Function CloseBar()
      SSProcess.EpsProgressDelete  
End Function
Function RollBar(barname,dispmsg)
    SSProcess.EpsProgressStepIt   
   SSProcess.EpsProgressUpdateMsg  barname  & dispmsg
End Function

'�ֽ��ַ���
Function ScanString(ByVal str, ByVal sep, ByRef strs(), ByRef count)
    Dim sepidx1, sepidx2,  strtemp
    count  = 0
    sepidx1 = 1
    sepidx2 = InStr(sepidx1 , str, sep, 1)
	  While (sepidx2 > 0)
       strs(count) = Mid( str, sepidx1, sepidx2-sepidx1)
        sepidx1 = sepidx2+1
       sepidx2 = InStr(sepidx1, str, sep, 1)
       count = count + 1
    Wend
    strs(count) = Mid( str, sepidx1, Len(str)+1-sepidx1)
    count = count + 1
End Function

'����ͼ���Ƿ�͸������ɫģʽ
Function SetLayerMode(byval layernamestr)
    Dim strs1(3000),scount1
	mapHandle  = SSProject.GetActiveMap  
	datasourceHandle = SSProject.GetActiveDatasource (mapHandle )
	lycount = SSProcess.GetLayerCount
	For  i = 0 to lycount-1
		strLayerName =  SSProcess.GetLayerName(i)
		layerHandle = SSProject.GetDataSourceLayerByname (datasourceHandle, strLayerName)
		If strLayerName=layernamestr  Then
			SSProject.SetLayerInfo layerHandle, "DrawAreaMode" , "1"  '0 ͸�� 1 ��͸�� 2 ��͸��
         SSProject.SetLayerInfo layerHandle, "ColorMode" , "5"  '�����
		End If
	Next
End Function

Function Setcg (mark)
	SSProcess.PushUndoMark
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==","9400403"
	SSProcess.SelectFilter
		Dim arID(1000), idCount,cgid(10000)
		JL = 100
		geoCount = SSProcess.GetSelGeoCount()
			For i = 0 To geoCount - 1
				cggeoCount =  0
				id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
				WHILE  (cggeoCount <> 1)  
					pointCount = SSProcess.GetSelGeoPointCount(i)
					For j = 0 To pointCount - 1
						SSProcess.GetSelGeoPoint i, j, x, y, z, pointType, name
						ids = SSProcess.SearchNearObjIDs (x, y, JL, 1, "9400603", 0) 
					Next
						SSFunc.Scanstring ids,",",cgid(10000),cggeoCount
						JL = JL-5
						if cggeoCount = 0  then
						JL = JL+10
						end if
				WEND


'�����Ҫ�޸���ȡ������ֵ
			SCCG = SSProcess.GetObjectAttr (ids, "[ShiCCG]")
			CC = SSProcess.GetObjectAttr (ids, "[��Ŀ����]")
			HXGUID = SSProcess.GetObjectAttr (ids, "[CS]")
			XKZGUID = SSProcess.GetObjectAttr (ids, "[JSGHXKZGUID]")
			JZWGUID= SSProcess.GetObjectAttr (ids, "[JZWMCGUID]")
			YDHXBH= SSProcess.GetObjectAttr (ids, "[GuiHYDXKZBH]")
			GHXKZBH= SSProcess.GetObjectAttr (ids, "[GuiHXKZBH]")
			JZWMC= SSProcess.GetObjectAttr (ids, "[JianZWMC]")

			SSProcess.SetObjectAttr id, "[YDHXGUID]", HXGUID
			SSProcess.SetObjectAttr id, "[JSGHXKZGUID]", XKZGUID
			SSProcess.SetObjectAttr id, "[JZWMCGUID]", JZWGUID
			SSProcess.SetObjectAttr id, "[GuiHYDXKZBH]", YDHXBH
			SSProcess.SetObjectAttr id, "[GuiHXKZBH]", GHXKZBH
			SSProcess.SetObjectAttr id, "[JianZWMC]", JZWMC
			SSProcess.SetObjectAttr id, "[CengG]", SCCG
			SSProcess.SetObjectAttr id, "[CengC]", CC
			Next
			mark = true
end Function
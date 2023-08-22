rem autor: <wsw> 
rem email: XXXX@xxx.com 
rem 脚本文件名: C:\Users\wsw\Desktop\放线\放线\1111.vbs
rem 对应方案文件名:F:\王邵微\2023工作\Z浙江\N宁波\程序\EPS多测合一宁波\DeskTop\多测合一\功能模板\放样成果图输出.Map
rem 方案名称:放样成果图
rem 本脚本文件应放置于 EPS安装目录\desktop\XX台面\Script\放样\放样成果图.vbs
rem framework: gq 
rem framework: 471b1e20fe69040339fca38c3d3a189b 



rem special:[放样成果图] 出图前（初始化调用）由此进入
Function VBS_preMap0(MSGID,mapName,selectID)
	  
	 rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
	 rem return = 1 停止输出成果图
	 rem return = 0 继续输出成果图（无需设置、默认值为0）
	  
	 If MSGID = 0 Then '// 新工程出图 
		TKFZ1 
		 '// 添加您的代码.... 
		 '// 设置出图工程名称、必须调用.... ,批量出图的路径每次会调用脚本传回的路径，工程不能同名，通常可以用范围线地物的扩展属性拼接
        strProjectName=SSProcess.GetProjectFileName()
        FileFolder=replace(strProjectName,".edb","")
			'FileFolder = SSProcess.GetSysPathName (5) 
			CreateFolders FileFolder
			SaveFile = FileFolder&"\建设工程实地放线平面图.edb"
			SSParameter.SetParameterSTR "printMap","NewedbName",SaveFile
	  
	 ElseIf MSGID = 1 Then '// 本工程出图 
		 '// 添加您的代码.... 
	  
	 ElseIf MSGID = 2 Then '// 新工程自定义目录出图(自主选择保存路径) 
		 '// 添加您的代码.... 
	  
	 End If 
	  
End Function 
	  
	  
	  
rem special:[放样成果图] 出图完成由此进入
Function VBS_postMap0(MSGID,mapName,selectID)
	  
	 rem 图廓ID,脚本处理项个数
	 Dim tk_id,tk_innerids,ScriptChangeCount
	 rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数 
	 Dim str_Name,str_para,str_paraex	  
	 rem 获取分层图图廓IDS,多个英文逗号相隔
	 SSParameter.GetParameterINT "printMap", "TKID", -1, tk_id 
	 rem 获取图廓内地物IDS
	 SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids 
	 rem 获取脚本处理项个数
	 SSParameter.GetParameterINT "printMap", "ScriptChangeCount", -1, ScriptChangeCount
	
	
	 '// 添加您的成果图处理代码 
	TKFZ2 
	fxtl
	 rem 成果图细节分开处理
	 For i = 0 to ScriptChangeCount -1
		 rem 获取处理项名称
		 SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name 
		 rem 获取处理项参数
		 SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para 
		 rem 获取处理项附加参数
		 SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex 
		 
		 
		 '// 此处无代码、说明没有脚本处理项..
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
		sql = "select VALUE from PROJECTINFO where KEY='测绘单位'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
		If nSeletionCount > 0 Then
			XMMC=arSeletionRecord(0)
		Else
			XMMC = ""
		End If
		
		sql = "select VALUE from PROJECTINFO where KEY='测量开始时间'"
		GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount

		If nSeletionCount > 0 Then
			HTRY=FormatDateTime(arSeletionRecord(0),1)
		Else
			HTRY=""
		End If

		
		sql = "select VALUE from PROJECTINFO where KEY='编制人员'"
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
		SSProcess.SetObjectAttr id, "[放线单位]", artemp(0)
		SSProcess.SetObjectAttr id, "[放线日期]", artemp(1)
		SSProcess.SetObjectAttr id, "[制图员]", artemp(2)
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
		'图形重新生成
		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
	next
	SSProcess.DeleteLayer "TKZSM"	
end function

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
			arSQLRecord(iRecordCount) =values										'查询记录
			iRecordCount =iRecordCount +1													'查询记录数
			'移动记录游标
			SSProcess.AccessMoveNext mdbName, sql
		Wend
	end if
	'关闭记录集
	SSProcess.CloseAccessRecordset mdbName, sql
End FUnction
	  
	  
	  
Dim g_MapList,g_MapPrePtrfun,g_MapPostPtrfun 
rem 主函数无需修改
Sub OnClick() 
	 
	rem 初始化 
	 g_MapList = Array("放样成果图")
	 g_MapPrePtrfun = Array("VBS_preMap0")
	 g_MapPostPtrfun = Array("VBS_postMap0")
	 
	 rem 系统传来的消息,用户选择的范围线ID,成果图名称
	 Dim str_msg,str_selectObjid,str_mapName 
	 
	 rem 获取系统参数--用户选择范围线ID
	 SSParameter.GetParameterINT "printMap", "SelectID", -1, str_selectObjid 

	 rem 获取系统参数--系统消息 （0：新工程固定目录出图初始化消息  1：本工程出图初始化消息  2: 新工程自定义目录出图初始化消息  3：出图已完成交付于脚本处理细节）
	 SSParameter.GetParameterINT "printMap", "printMSG", -1, str_msg  

	 rem 获取系统参数--专题名称
	 SSParameter.GetParameterSTR "printMap", "SpecialMapName", "", str_mapName 

	 DistributeMSG str_msg,str_mapName,str_selectObjid 


End Sub




rem 此虑数函数无需修改
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

'// 检查成果目录是否存在、如果不存在放弃出图
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

'// 获取本工程项目名称
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

'// 判断文件是否存在
Function FileExists(fileName)
	Dim fso
	Set fso = CreateObject("scripting.filesystemobject")
	FileExists = fso.FileExists(fileName)
End Function 

'创建文件夹
function CreateFolders(pathname)
	Set fso = CreateObject("Scripting.FileSystemObject")
	newpathname= pathname
	if Not fso.folderExists(newpathname)  then
		fso.CreateFolder   newpathname   '创建文件夹
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
       '竖线
         makeLine x0,y0,x0,y0+count5*3+2.5,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID

			makeLine x0+18,y0,x0+18,y0+count5*3+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
			'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
		 '横线
			'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0,x0+18,y0,1, "RGB(255,255,255)", polygonID
			'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
			makeLine x0,y0+count5*3+2.5,x0+18,y0+count5*3+2.5,1, "RGB(255,255,255)", polygonID
			makeNote x0+8,y0+count5*3+1 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID

			for j= 0 to ubound(arDrawCode)
			 '竖线
               CodeType=SSProcess.GetFeatureCodeInfo(arDrawCode(j), "Type") 
               'makeLine x0+1,y0+j*2+1.5,x0+7,y0+j*2+1.5,arDrawCode(j), arDrawColor(j), polygonID
			      'makeLine x0,y0+j*2+2.5,x0+16,y0+j*2+2.5, 1,"RGB(255,255,255)", polygonID
					'makeNote x0+10,y0+1.5+ j*2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
               if CodeType=3 or  CodeType=2 or  CodeType=1  then '线
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
		SSProcess.SetNewObjValue "SSObj_LayerName", "放线平面图图廓信息"
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
		SSProcess.SetNewObjValue "SSObj_LayerName", "放线平面图图廓信息"
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
		SSProcess.SetNewObjValue "SSObj_LayerName", "放线平面图图廓信息"
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
		SSProcess.SetNewObjValue "SSObj_LayerName", "放线平面图图廓信息"
		SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
		SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
		SSProcess.SetNewObjValue "SSObj_FontWidth", width
		SSProcess.SetNewObjValue "SSObj_FontHeight", height
		SSProcess.AddNewObjPoint x, y, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 
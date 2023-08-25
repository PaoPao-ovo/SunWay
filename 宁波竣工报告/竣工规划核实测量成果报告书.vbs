dim g_docObj'doc全局变量
f_subYWMC = "竣工多测合一:规划核实测量,用地复核测量,停车场（库）核实测量,消防核实测量,绿地核实测量,人防核实测量,地下管线测量"
'业务节点索引从0开始
strSectionIndexList0 = "消防核实测量:37,36;人防核实测量:35,34,33,32,31,30;绿地核实测量:29;"
strSectionIndexList1="绿地核实测量:24;地下管线测量:23,22,21;停车场（库）核实测量:20;规划核实测量:19;用地复核测量:18;规划核实测量:17;规划核实表:16;规划核实测量:15,14,13,12,11;停车场（库）核实测量:10;规划核实测量:9,8,7,6,5,4"
ExpDocYWMC="消防核实测量,人防核实测量,绿地核实测量,地下管线测量,停车场（库）核实测量,用地复核测量,规划核实测量"
Sub OnClick()
		strTempFileName="建设工程竣工规划核实成果报告书0818.doc"
		strTempFilePath = SSProcess.GetSysPathName (7) &strTempFileName
		Set g_docObj = CreateObject ("asposewordscom.asposewordshelper")
		If  TypeName (g_docObj) = "AsposeWordsHelper" Then 
			g_docObj.CreateDocumentByTemplate strTempFilePath
		Else
			msgbox "请先注册Aspose.Word插件":Exit Sub
		End If
		strSubYWMC =   SSPro.VbsCheckDlg (f_subYWMC)
		if strSubYWMC="" then msgbox "请选择需要输出成果报告的竣工业务！":exit sub else strSubYWMC = replace(strSubYWMC,"竣工多测合一:","")
		pathName = GetFilePath'SSProcess.SelectPathName()
		g_docObj.CreateDocumentByTemplate  strTempFilePath
		fwnr=SSProcess.ReadEpsIni("签章GUID", "fwnr" ,"")
		
		InitDB() 
			'字符替换 
			ReplaceValue
			'封面
			ReplaceValueFM
			'测绘项目技术人员
			OutputTable0 1
			'总体目标表
			OutputTable01 2 ,strSubYWMC
			'控制点坐标表
			OutputTable02 3
			'输出报告重组业务数组排序
			strSubYWMC=DocArrSort( strSubYWMC, ExpDocYWMC , str)
			arSubYWMC = split(strSubYWMC,",")
			for i = 0 to ubound(arSubYWMC)
				if arSubYWMC(i)= "规划核实测量" then
					'规划新增节点个数
					GHnode=0
					'输出 建筑物建筑面积汇总表
					OutputTable14 5
					'输出 建筑面积变更规划核实表
					OutputTable15 7
					'输出 特殊部位计算说明规划核实表
					OutputTable16 8
					'输出 建筑基底平面示意图
					OutputBook "建筑基底平面示意图","建筑基底平面示意图","9420032",10
					'输出 建筑功能分区竣工实测平面图
					OutputBook "建筑功能分区竣工实测平面图","建筑功能分区竣工实测平面图","9420031",11
					'输出 竣工图
					OutputBook "竣工图","竣工图","9420033",""
					'输出 竣工规划总平面图
					OutputBook "竣工测绘总平面图","竣工测绘总平面图","9420034",""
					'输出 竣工规划复核图
					OutputBook "竣工规划复核图","竣工规划复核图","9420035",""
					'输出 基地总平面布置核实测量平面图
					OutputBook "基地总平面布置核实测量平面图","基地总平面布置核实测量平面图","9420037",""
					'输出 规划测量分幢与规划许可比对结果表
					OutputTable17 9,GHnode1
					'输出 主要经济技术指标比对表
					OutputTable11 6
					'输出 建（构）筑物竣工标高测量一览表
					OutputTable10 4,GHnode2
					'实景三维模型概览图
					OutMap "实景三维模型概览图",dmark
				elseif arSubYWMC(i)= "用地复核测量" then
					'输出 用地复核图
					OutputBook "用地复核图","用地复核图","9420036",""
				elseif arSubYWMC(i)= "停车场（库）核实测量" then
					'插入 停车场（库）核实测量平面图
					'OutputTable6 10,"停车场（库）核实测量平面图","9460093"
					OutputBook "停车场（库）核实测量平面图","停车场（库）核实测量平面图","9460093",""
					'输出 停车位序号一览表
					OutputTable7 10
				elseif arSubYWMC(i)= "消防核实测量" then
					'插入 总平面测量略图
					'OutputTable6 25,"总平面测量略图","9430093"
					OutputBook "总平面测量略图","总平面测量略图","9430093",""
					'输出 总平面布局测量表
					OutputTable8 27
				elseif arSubYWMC(i)= "绿地核实测量" then
					lvDSection=""
					'插入 绿地竣工地形图
					'OutputTable6 17,"绿地竣工图","9470105"
					OutputBook "绿地竣工图","绿地竣工图","9470105",""
					'绿地面积统算表
					OutputTable5 19
					'输出 垂直绿化面积明细表
					OutputTable4 18,lvDSection
					'输出 休憩场所面积明细表
					OutputTable3 17,lvDSection
					'输出 地下设施面积屋顶绿地面积明细表
					OutputTable2 16,lvDSection
					'输出 地面绿地面积明细表
					OutputTable1 15
				elseif arSubYWMC(i)= "人防核实测量" then
					'插入 其他测量图
					'OutputTable6 23,"其他测量图","9450083"
					OutputBook "其他测量图","其他测量图","9450083",""
					'插入 防护单元核实面积测量图
					'OutputTable6 22,"防护单元核实面积测量图","9450073"
					OutputBook "防护单元核实面积测量图","防护单元核实面积测量图","9450073",""
					'插入 人防总平面示意图
					'OutputTable6 21,"人防总平面示意图","9450063"
					OutputBook "人防总平面示意图","人防总平面示意图","9450063",""
					'输出 人防面积测绘计算表
					OutputTable13 23
					'输出 人防测量成果表
					OutputTable12 22
					'输出		人防基本信息表
					OutputTable18	21
				elseif arSubYWMC(i)= "地下管线测量" then
					'输出 地下管线测量成果表
					OutputTable9 14
				end if
			next
			GHnode=GHnode1+GHnode2
			'插入签章与水印
			InsertSignature
			'删除不输出的业务
			strSectionIndexList=strSectionIndexList0&";"&lvDSection&";"&strSectionIndexList1
			DeleteYWMCSection strSubYWMC,strSectionIndexList,GHnode,dmark
			'目录
			' g_docObj.UpdateFields()
		ReleaseDB()
		'签章
		if fwnr="" then
			fwnr = fwnrGUid()
			SSProcess.WriteEpsIni "签章GUID", "fwnr" ,fwnr
		else
			fwnr = fwnr
		end if
		
		strFileSavePath=pathName&replace(strTempFileName,".doc",".doc")
		'strFileSavePath=replace(strFileSavePath,".docx",".pdf")
		g_docObj.SaveEx  strFileSavePath
		bRes=ProtectDoc(strFileSavePath,true,fwnr)
		set g_docObj=nothing
		msgbox "输出完成"
End Sub
#include ".\function\SQLOperateVbsFunc.vbs"

'//插入 成果图
function OutMap(byval MapName,byref dmark)
	mdbName = SSProcess.GetSysPathName (5)
	filePath=replace(mdbName,".edb","")&"\"&MapName&"\"
	dim imageList(10000):listCount=0
	GetAllFiles filePath,"jpg",listCount,imageList
	if listCount=0 then dmark=false:exit function else dmark=true
	for i=0 to listCount-1
		imageFile=imageList(i)
		name=GetFileName(imageFile)
		extensionName=GetFileExtensionName(imageFile)
		name=replace(name,"."&extensionName,"")
		'if instr(name,fileName)>0 then 
				insertImageFile=imageFile 
				if FileExists( insertImageFile)  =true then  
					g_docObj.MoveToBookmark	name
					RES = g_docObj.InsertImage (insertImageFile,350,350,0)   
				end if
		'end if 
	next

end function

function DocArrSort(byval strSubYWMC,byval ExpDocYWMC ,byref str)
	arSubYWMC = split(strSubYWMC,",")
	arExpDocYWMC=split(ExpDocYWMC,",")
	str=""
	for i=0 to ubound(arExpDocYWMC)
		for i1=0 to ubound(arSubYWMC)
			if arExpDocYWMC(i)=arSubYWMC(i1) then
				if str="" then
					str=arExpDocYWMC(i)
				else
					str=str&","&arExpDocYWMC(i)
				end if
			end if
		next
	next
	DocArrSort=str
end function

'//签章guid
function fwnrGUid()
		set TypeLib = CreateObject("Scriptlet.TypeLib")
		fwnrGUid=TypeLib.Guid
		fwnrGUid=replace(fwnrGUid,"-","")
		fwnrGUid=replace(fwnrGUid,"{","")
		fwnrGUid=replace(fwnrGUid,"}","")
		fwnrGUid = Left(fwnrGUid,10)
		set TypeLib=nothing
end function

'//加密解密docx文件
'//strFilePath 需要加密的doc文件路径
'//isProtectDoc true 加密 false解密
'//password 密码
Function ProtectDoc(byval strFilePath, byval  isProtectDoc,byval password)
	bRes=false
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(strFilePath))=true  Then
		Set pDocObj = CreateObject ("asposewordscom.asposewordshelper")
		If  TypeName (pDocObj) = "AsposeWordsHelper" Then 
			pDocObj.OpenDocument strFilePath
		Else
			bRes=false
			Exit Function
		End If

		if isProtectDoc=true then
			str=pDocObj.ProtectDoc (password)
		else
			str=pDocObj.UnProtectDoc (password)
		end if 
		pDocObj.SaveEx  strFilePath
		set pDocObj=nothing
	end if 
	set fso=nothing
	if instr(str,"成功") then bRes=true
	ProtectDoc=bRes
End Function

'//输出 测绘项目技术人员
Function OutputTable0(byval tableIndex)
  g_docObj.MoveToTable tableIndex,false
  '获取人员信息表单元格
  cellCount=0:redim cellList(cellCount)
  strField="姓名,职称或职业资格,上岗证书编号或职业资格证书号,主要工作职责,备注"
  listCount=GetProjectTableList ("info_RYXX",strField,"ID>0","","",list,fieldCount)
  for i=0 to listCount-1
   cellValue=""
   for j= 0 to fieldCount-1
    value=list(i,j)
    if j=0 then  cellValue=value else cellValue=cellValue&"||"&value
   next
   cellValue=i+1&"||"&cellValue
   redim preserve cellList(cellCount):cellList(cellCount)=cellValue:cellCount=cellCount+1
  next
  
  '填充人员信息表单元格
  iniRow=1:iniCol=0
  startRow=iniRow:startCol=iniCol
  if cellCount>1 then   g_docObj.CloneTableRow tableIndex, iniRow, 1,cellCount-1, false
  for i= 0 to cellCount-1
   startCol=iniCol
   cellValue=cellList(i)
   cellValueList=split(cellValue,"||")
   for j= 0 to ubound(cellValueList)
    g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
    startCol=startCol+1
   next
   startRow=startRow+1
  next
End Function

'//删除不输出的业务
Function DeleteYWMCSection(strSubYWMC,strSectionIndexList,byval GHnode,byval dmark)
	arSubYWMC = split(strSubYWMC,",")
	arSection = split(strSectionIndexList,";")
	strReplaceList = ""
	strDeleteList = ""
	'字符串去掉输出业务
	for i = 0 to ubound(arSection)
		for i1 = 0 to ubound(arSubYWMC)
			arSection(i) = replace(arSection(i),arSubYWMC(i1),"*")
		next
		if strReplaceList = "" then
			strReplaceList = arSection(i)
		else
			strReplaceList = strReplaceList&";"&arSection(i)
		end if
	next
	'筛选删除业务分节符索引
	arReplaceList = split(strReplaceList,";")
	for i = 0 to ubound(arReplaceList)
		if replace(arReplaceList(i),"*","")=arReplaceList(i) then
			if strDeleteList = "" then
				strDeleteList = arReplaceList(i)
			else
				strDeleteList = strDeleteList&";"&arReplaceList(i)
			end if
		end if
	next

	'删除文档分节
	arDeleteList = split(strDeleteList,";")
	for i = 0 to ubound(arDeleteList)
		strSubDeleteList = mid(arDeleteList(i),instr(arDeleteList(i),":")+1,len(arDeleteList(i))-instr(arDeleteList(i),":"))
		arSubDeleteList = split(strSubDeleteList,",")
		for i1 = 0 to ubound(arSubDeleteList)
			g_docObj.RemoveSection(arSubDeleteList(i1)+GHnode)
		next
	next
	'删除实景三维概览图
	if dmark=false then 
		if instr(strReplaceList,"停车场")>0 then
			g_docObj.RemoveSection(11+GHnode-1)
		else
			g_docObj.RemoveSection(11+GHnode)
		end if
	end if

End Function

'//插入签章
Function InsertSignature
		folderPath =  SSProcess.GetSysPathName (0)&"\签章\"
		names="水印":nameList=split(names,",")
		for i= 0 to ubound(nameList)
			name=nameList(i)
			imageFile=folderPath&name&".png"
			if name="水印" then 
				if IsFileExists(imageFile)=true then    g_docObj.SetImgWatermark imageFile, 400, 400,0
			else
				g_docObj.MoveToBookmark name
				if IsFileExists(imageFile)=true then    g_docObj.InsertImageEx imageFile,  0, 250, 0, 390, 150, 150,3, 0
			end if 
		next
End Function

'//判断文件是否存在
Function IsFileExists(filespec)
	IsFileExists=false
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec))=true Then
         IsFileExists = true
   End If
   Set fso = Nothing
End Function

'//获取成果目录路径
Function  GetFilePath
		projectFileName=SSProcess.GetProjectFileName()
		filePath=replace(projectFileName,".edb","")
		filePath=filePath&"\"
		CreateFolder filePath
		GetFilePath=filePath
End Function




'//递归创建多级目录
Function CreateFolder(path)
		Set fso = CreateObject("scripting.filesystemobject")
		If fso.FolderExists(path) Then
			Exit Function
		End If
		If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
			CreateFolder fso.GetParentFolderName(path)
		End If
		fso.CreateFolder(path)
		set fso=nothing
End Function


Function AddLoginfo(msg)
     SSProcess.MapCallBackFunction "OutputMsg", "[" & now & "] " & msg, 1 
End Function


'//字符替换 
Function ReplaceValue
		'人防工程基本情况表
		strTableName="projectinfo"
		values="项目名称,项目地址"
		valuesList=split(values,",")
		for i= 0 to ubound(valuesList)
			strFieldValue=""
			strField=valuesList(i)
			listCount=GetProjectTableList (strTableName,"value","key='"&strField&"'","","",list,fieldCount)
			if listCount=1 then strFieldValue=list(0,0)
			g_docObj.Replace "{"&strField&"}",strFieldValue,0
		next
		
		'人防测量成果表
		strTableName="RFPROJECTINFO"
		values="建筑结构,住宅户数,地上建筑面积(㎡),地上住宅建筑面积(㎡),地上其他建筑面积(㎡),地上层数,地下平时功能,地下建筑面积(㎡),地下层数,互联互通面积,防空警报控制室面积,外墙最薄掩体厚度（小于10米时填写）,板坪高差（顶板底面高出室外时填写）,编制人,检查人"
		valuesList=split(values,",")
		for i= 0 to ubound(valuesList)
			strFieldValue=""
			strField=valuesList(i)
			listCount=GetProjectTableList (strTableName,"value","key='"&strField&"'","","",list,fieldCount)
			if listCount=1 then strFieldValue=list(0,0)
			g_docObj.Replace "{"&strField&"}",strFieldValue,0
		next
End Function


'//输出 地面绿地面积明细表
Function OutputTable1(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		
		writeRowCount=12
		copyCount=0
		code="9470103":tableType =2
		strTableName=SSProcess.GetCodeAttrTableName(code,tableType)
		listCount=GetProjectTableList (strTableName,strTableName&".id,lhbh,lhmj","lhlx='地面绿化' order by  lhbh asc","SpatialData",tableType,list,fieldCount)
		redim cellList(listCount)
		for i = 0 to listCount-1
			objid=list(i,0):lhbh=list(i,1):lhmj=list(i,2)
			lhmj=GetFormatNumber(lhmj,2)
			cellValue=lhbh&"||"&lhmj
			cellList(cellCount)=cellValue
			cellCount=cellCount+1
			if i>0 and i mod writeRowCount*4=0 then copyCount=copyCount+1
		next
		
		'数组按 lhbh 从小到大冒泡排序
		for i= 0 to cellCount-1
			for j= 0 to cellCount-1-1
				cellValue=cellList(j):cellValueList=split(cellValue,"||")
				cellValue1=cellList(j+1):cellValueList1=split(cellValue1,"||")
				num1=right(cellValueList(0),len(cellValueList(0))-1)
				num2=right(cellValueList1(0),len(cellValueList1(0))-1)
				if isnumeric(num1)=true and isnumeric(num2) =true then 
					if cdbl(num1)>cdbl(num2) then 
						temp=cellList(j)
						cellList(j)=cellList(j+1)
						cellList(j+1)=temp
					end if 
				end if 
			next
		next
		
		'根据数据个数复制表格
		for i=0 to copyCount-1
			g_docObj.CloneTable  tableIndex, 1,0,false
		next
		
		'填充表格单元格
		iniRow=2:iniCol=0
		startRow=iniRow:startCol=iniCol
		colIndex=0:index1=0:index2=0
		for i=0 to listCount-1
			cellValue=cellList(i):cellValueList=split(cellValue,"||")
			lhbh=cellValueList(0):lhmj=cellValueList(1)
			
			'根据数组个数动态计算表、行、列索引
			if i>0 and i mod writeRowCount=0 then startRow=iniRow:index1=index1+1
			if i>0 and i mod writeRowCount*4=0 then tableIndex=tableIndex+1 :index2=index2+1
			if index1 mod 4=0 then startCol=iniCol
			if index1-(index2*4)=1 then startCol=iniCol+2
			if index1-(index2*4)=2 then startCol=iniCol+4
			if index1-(index2*4)=3 then startCol=iniCol+6
			g_docObj.SetCellText tableIndex,startRow,startCol,lhbh,true,false
			g_docObj.SetCellText tableIndex,startRow,startCol+1,lhmj,true,false
			startRow=startRow+1
		next
End Function


'//输出 地下设施面积屋顶绿地面积明细表
Function OutputTable2(byval tableIndex,byref lvDSection)
		g_docObj.MoveToTable tableIndex,false
		
		code="9470103":tableType =2
		'查找 地下设施顶面绿化 属性
		strTableName=SSProcess.GetCodeAttrTableName(code,tableType)
		listCount=GetProjectTableList (strTableName,strTableName&".id,lhbh,lhhd,lhmj","lhlx='地下设施顶面绿化'  order by lhbh asc","SpatialData",tableType,list,fieldCount)
		redim cellList1(10000):cellCount1=0
		for i = 0 to listCount-1
			objid=list(i,0):lhbh=list(i,1):lhhd=list(i,2):lhmj=list(i,3)
			area=SSProcess.GetObjectAttr (objid,"SSObj_Area"):area=GetFormatNumber(area,2)
			lhmj=GetFormatNumber(lhmj,2)
			lhhd=GetFormatNumber(lhhd,1)
			cellValue1=lhbh&"||"&area&"||"&lhhd&"||"&lhmj
			cellList1(i)=cellValue1
			cellCount1=cellCount1+1
		next
		'查找 屋顶绿地 属性
		listCount=GetProjectTableList (strTableName,strTableName&".id,lhbh,lhhd,lhmj","lhlx='屋顶绿地'  order by lhbh asc","SpatialData",tableType,list,fieldCount)
		redim cellList2(10000):cellCount2=0
		for i = 0 to listCount-1
			objid=list(i,0):lhbh=list(i,1):lhhd=list(i,2):lhmj=list(i,3)
			area=SSProcess.GetObjectAttr (objid,"SSObj_Area"):area=GetFormatNumber(area,2)
			lhmj=GetFormatNumber(lhmj,2)
			lhhd=GetFormatNumber(lhhd,1)
			cellValue2=lhbh&"||"&area&"||"&lhhd&"||"&lhmj
			cellList2(i)=cellValue2
			cellCount2=cellCount2+1
		next

		if cellCount1+cellCount2=0 then LvDSection=LvDSection&";"&"地下设施面积屋顶绿地面积明细表:25"else LvDSection=LvDSection&";"&"绿地核实测量:25"
		'属性重新组织后写入输出数组
		if cellCount2>cellCount1 then cellCount=cellCount2 else cellCount=cellCount1
		redim cellList(cellCount)
		for i= 0 to cellCount-1
			cellValue1=cellList1(i)
			cellValue2=cellList2(i)
			if  cellValue1="" then  cellValue1=""&"||"&""&"||"&""&"||"&""
			if  cellValue2="" then  cellValue2=""&"||"&""&"||"&""&"||"&""
			cellValue= cellValue1&"||"&cellValue2
			cellList(i)=cellValue
		next
		
		writeRowCount=11
		copyCount=0
		for i= 0 to cellCount-1
			if i>0 and i mod writeRowCount=0 then copyCount=copyCount+1
		next
		
		'根据数据个数复制表格
		for i=0 to copyCount-1
			g_docObj.CloneTable  tableIndex, 1,0,false
		next
		
		'填充表格单元格
		iniRow=3:iniCol=0
		startRow=iniRow:startCol=iniCol
		for i=0 to cellCount-1
			cellValue=cellList(i):cellValueList=split(cellValue,"||")
			if i>0 and i mod writeRowCount=0 then tableIndex=tableIndex+1 :startRow=iniRow
			
			if  ubound(cellValueList)=7 then 
				for  j= 0 to ubound(cellValueList)
					g_docObj.SetCellText tableIndex,startRow,startCol+j,cellValueList(j),true,false
				next
			end if
			startRow=startRow+1
		next
End Function


'//输出 休憩场所面积明细表
Function OutputTable3(byval tableIndex,byref LvDSection)
		g_docObj.MoveToTable tableIndex,false
		
		code="9470103":tableType =2
		strTableName=SSProcess.GetCodeAttrTableName(code,tableType)
		'查找 景观水体 属性
		redim cellList1(10000)
		GetCellList strTableName, tableType,"景观水体",cellList1,listCount1,sumArea1
		'查找 园路及园林铺装 属性
		redim cellList2(10000)
		GetCellList strTableName, tableType,"园路及园林铺装",cellList2,listCount2,sumArea2
		'查找 园林小品 属性
		redim cellList3(10000)
		GetCellList strTableName, tableType,"园林小品",cellList3,listCount3,sumArea3
		'查找 其他 属性
		redim cellList4(10000)
		GetCellList strTableName, tableType,"其他",cellList4,listCount4,sumArea4
		if listCount1+listCount2+listCount3+listCount4=0 then LvDSection=LvDSection&";"&"休憩场所面积明细表:26"else LvDSection=LvDSection&";"&"绿地核实测量:26"
		countValue=listCount1&","&listCount1&","&listCount2&","&listCount4
		countValueList=split(countValue,",")
		'数组元素个数冒泡排序
		for i= 0 to ubound(countValueList)
			for j= 0 to ubound(countValueList)-1
				if isnumeric(countValueList(j))=true and isnumeric(countValueList(j+1)) =true then 
					if cdbl(countValueList(j))>(cdbl(countValueList(j+1))) then 
						temp=countValueList(j)
						countValueList(j)=countValueList(j+1)
						countValueList(j+1)=temp
					end if 
				end if 
			next
		next
		'属性重新组织后写入输出数组
		cellCount=countValueList(ubound(countValueList))
		redim cellList(cellCount)
		for i= 0 to cellCount-1
			cellValue1=cellList1(i)
			cellValue2=cellList2(i)
			cellValue3=cellList3(i)
			cellValue4=cellList4(i)
			if  cellValue1="" then  cellValue1=""&"||"&""
			if  cellValue2="" then  cellValue2=""&"||"&""
			if  cellValue3="" then  cellValue3=""&"||"&""
			if  cellValue4="" then  cellValue4=""&"||"&""
			cellValue= cellValue1&"||"&cellValue2&"||"&cellValue3&"||"&cellValue4
			cellList(i)=cellValue
		next
		
		writeRowCount=22
		copyCount=0
		for i= 0 to cellCount-1
			if i>0 and i mod writeRowCount=0 then copyCount=copyCount+1
		next
		
		'根据数据个数复制表格
		for i=0 to copyCount-1
			g_docObj.CloneTable  tableIndex, 1,0,false
		next
		
		'填充表格单元格
		iniRow=3:iniCol=0
		startRow=iniRow:startCol=iniCol
		for i=0 to cellCount-1
			cellValue=cellList(i):cellValueList=split(cellValue,"||")
			if i>0 and i mod writeRowCount=0 then tableIndex=tableIndex+1 :startRow=iniRow
			if  ubound(cellValueList)=7 then 
				for  j= 0 to ubound(cellValueList)
					g_docObj.SetCellText tableIndex,startRow,startCol+j,cellValueList(j),true,false
				next 
			end if
			startRow=startRow+1
		next
		
		g_docObj.Replace "{合计1}",sumArea1,0
		g_docObj.Replace "{合计2}",sumArea2,0
		g_docObj.Replace "{合计3}",sumArea3,0
		g_docObj.Replace "{合计4}",sumArea4,0
End Function


Function GetCellList(byval strTableName,byval tableType,byval lhzlx,cellList,listCount,sumArea)
		'查找 屋顶绿地 属性
		sumArea=0
		listCount=GetProjectTableList (strTableName,strTableName&".id,lhbh,lhmj","lhlx='休憩场所' and lhzlx='"&lhzlx&"'  order by lhbh asc","SpatialData",tableType,list,fieldCount)
		'redim cellList(listCount)
		for i = 0 to listCount-1
			objid=list(i,0):lhbh=list(i,1):lhmj=list(i,2)
			lhmj=GetFormatNumber(lhmj,2)
			cellValue=lhbh&"||"&lhmj
			cellList(i)=cellValue
			sumArea=sumArea+cdbl(lhmj)
		next
		sumArea=GetFormatNumber(sumArea,2)
End Function


'//输出 垂直绿化面积明细表
Function OutputTable4(byval tableIndex,byref LvDSection)
		g_docObj.MoveToTable tableIndex,false
		
		strTableName="VerticalGreening"
		listCount=GetProjectTableList (strTableName,"LDBH,ZZCD,CTKD,FTHD,PJDMMJ,GJCKMJ,ZSLDMJ",strTableName&".ID>0","","",list,fieldCount)
		if  listCount=0 then LvDSection="垂直绿化面积明细表:27" else LvDSection="绿地核实测量:27"
		redim cellList(listCount)
		for i = 0 to listCount-1
			cellValue=""
			for j= 0 to fieldCount-1
				value=list(i,j)
				if j=4 or j=5 or j=6 then value=GetFormatNumber(value,2)
				if j=0 then cellValue=value else cellValue=cellValue&"||"&value
			next
			cellList(cellCount)=cellValue
			cellCount=cellCount+1
		next
		
		writeRowCount=18
		copyCount=0
		for i= 0 to cellCount-1
			if i>0 and i mod writeRowCount=0 then copyCount=copyCount+1
		next
		
		'根据数据个数复制表格
		for i=0 to copyCount-1
			g_docObj.CloneTable  tableIndex, 1,0,false
		next
		
		iniRow=1:iniCol=0
		startRow=iniRow:startCol=iniCol
		
		'填充表格单元格
		for i=0 to cellCount-1
			cellValue=cellList(i):cellValueList=split(cellValue,"||")
			if i>0 and i mod writeRowCount=0 then tableIndex=tableIndex+1 :startRow=iniRow
			
			if  ubound(cellValueList)=6 then 
				for  j= 0 to ubound(cellValueList)
					g_docObj.SetCellText tableIndex,startRow,startCol+j,cellValueList(j),true,false
				next
			end if
			startRow=startRow+1
		next
End Function


'//输出 绿地面积统算表
Function OutputTable5(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		listCount=GetProjectTableList ("LHHF","MC"," id>0 ","","",list,fieldCount)
		if listCount>0 then 

			'**************************************************写单块绿地-->地面及地下设施顶面绿化
			iniRow=1:iniCol=2
			startRow=iniRow:startCol=iniCol
			GetAreaList "'地面绿化','地下设施顶面绿化','休憩场所'","'单块绿地'",areaList1,count1,sumJrmj1
			iniRowCount1=6
			if count1>iniRowCount1 then   g_docObj.CloneTableRow tableIndex,  startRow+1, 1,count1-iniRowCount1, false
			for i= 0 to count1-1
				values= areaList1(i)
				valuesList=split(values,"||")
				if ubound(valuesList)=5 then 
					startCol=iniCol
					for j= 0 to ubound(valuesList)-1
						g_docObj.SetCellText tableIndex,startRow,startCol+j,valuesList(j),true,false
					next
				end if 
				mergeMark=valuesList(5)
				if mergeMark="合并" then 
					g_docObj.MergeCell tableIndex,  startRow,  iniCol,  startRow+1,  iniCol,false
					g_docObj.MergeCell tableIndex,  startRow,  iniCol+4,  startRow+1,  iniCol+4,false
				end if 
				startRow=startRow+1
			next

			
			'**************************************************写单块绿地-->屋顶及垂直绿化
			if  count1<7  then  startRow=iniRowCount1+iniRow+1   else    	startRow=startRow+1
			GetAreaList "'屋顶绿地'","'单块绿地'",areaList1,count2,sumJrmj2
			iniRowCount2=4
			if count2>iniRowCount2 then   g_docObj.CloneTableRow tableIndex,  startRow+1, 1,count2-iniRowCount2, false
			for i= 0 to count2-1
				values= areaList1(i)
				valuesList=split(values,"||")
				if ubound(valuesList)=5 then 
					startCol=iniCol
					for j= 0 to ubound(valuesList)-1
						if j=3   then if valuesList(j)=0 then valuesList(j)="删除单元格"
						g_docObj.SetCellText tableIndex,startRow,startCol+j,valuesList(j),true,false
					next
				end if 
				mergeMark=valuesList(5)
				if mergeMark="合并" then 
					g_docObj.MergeCell tableIndex,  startRow,  iniCol,  startRow+1,  iniCol,false
					g_docObj.MergeCell tableIndex,  startRow,  iniCol+4,  startRow+1,  iniCol+4,false
				end if 
				startRow=startRow+1
			next
			
			'**************************************************写集中绿地
			if count2<5  then startRow=iniRowCount1+iniRow+iniRowCount2+2 else    startRow=startRow+1
			GetAreaList "'屋顶绿地','地下设施顶面绿化','地面绿化'","'集中绿地'",areaList1,count1,sumJrmj3
			iniRowCount=2
			if count1>iniRowCount then   g_docObj.CloneTableRow tableIndex,  startRow+1, 1,count1-iniRowCount, false
			for i= 0 to count1-1
				values= areaList1(i)
				valuesList=split(values,"||")
				if ubound(valuesList)=5 then 
					startCol=iniCol-1
					for j= 0 to ubound(valuesList)-1
						if j=3   then if valuesList(j)=0 then valuesList(j)="删除单元格"
						g_docObj.SetCellText tableIndex,startRow,startCol+j,valuesList(j),true,false
					next
				end if 
				mergeMark=valuesList(5)
				if mergeMark="合并" then 
					g_docObj.MergeCell tableIndex,  startRow,  iniCol-1,  startRow+1,  iniCol-1,false
					g_docObj.MergeCell tableIndex,  startRow,  iniCol+3,  startRow+1,  iniCol+3,false
				end if 
				startRow=startRow+1
			next
			sumAllJrmj=cdbl(sumJrmj1)+cdbl(sumJrmj2)+cdbl(sumJrmj3):sumAllJrmj=GetFormatNumber(sumAllJrmj,2)
			

			'筛选出需要删除的单元行
			rowCount=0:redim  deleteRowList(rowCount)
			tableRowCount=g_docObj.GetTableRowCount(tableIndex, false)
			for i= 0 to tableRowCount-1 
				tableColCount=g_docObj.GetTableColCount( tableIndex, i, false)
				for j= 0 to tableColCount-1
					str=g_docObj.GetCellText( tableIndex, i, j,false)
					if str="计入面积" then mark=mark&","&"'"&i&"'"
					if str="删除单元格" then   
						redim preserve  deleteRowList(rowCount): deleteRowList(rowCount)=i:rowCount=rowCount+1
					elseif str="{计入面积1}"   then 
						if sumJrmj1=0 then redim preserve  deleteRowList(rowCount): deleteRowList(rowCount)=i:rowCount=rowCount+1
					elseif str="{计入面积2}"   then 
						if sumJrmj2=0 then redim preserve  deleteRowList(rowCount): deleteRowList(rowCount)=i:rowCount=rowCount+1
					elseif str="{计入面积3}"   then 
						for m= 1 to i-1
							if instr(mark,"'"&m&"'")=0 then 	
								str1=g_docObj.GetCellText( tableIndex, m, 5,false)
								if str1="" then 
									redim preserve  deleteRowList(rowCount): deleteRowList(rowCount)=m:rowCount=rowCount+1
								end if 
							end if 
						next
						if sumJrmj3=0 then redim preserve  deleteRowList(rowCount): deleteRowList(rowCount)=i:rowCount=rowCount+1
					end if 
				next
			next
			'单元行冒泡排序
			for i= 0 to rowCount-1
				for j= 0 to rowCount-1-1
					if isnumeric(deleteRowList(j))=true and isnumeric(deleteRowList(j+1)) =true then 
						if cdbl(deleteRowList(j))<(cdbl(deleteRowList(j+1))) then 
							temp=deleteRowList(j)
							deleteRowList(j)=deleteRowList(j+1)
							deleteRowList(j+1)=temp
						end if 
					end if 
				next
			next
			'删除单元行
			for i= 0 to rowCount-1
				g_docObj.DeleteRow tableIndex, deleteRowList(i), false
			next
		end if
		listCount=GetProjectTableList ("JGSCHZXX","YDMJ","ID>0","","",list,fieldCount)
		if listCount=1 then sc_YongDMJ = list(0,0)
		if sc_YongDMJ="" then sc_YongDMJ=0.00 else sc_YongDMJ=GetFormatNumber(sc_YongDMJ,2)
		g_docObj.Replace "{计入面积1}",sumJrmj1,0
		g_docObj.Replace "{计入面积2}",sumJrmj2,0
		g_docObj.Replace "{计入面积3}",sumJrmj3,0
		g_docObj.Replace "{绿地合计}",sumAllJrmj,0
		g_docObj.Replace "{绿地总用地面积}",sc_YongDMJ,0
		ldCount=GetProjectTableList ("JGSCHZXX","LVL","ID>0","","",sclhYdmjList,fieldCount)
		if ldCount = 1 then sc_lhl=sclhYdmjList(0,0)
		if sc_lhl="" then sc_lhl=0.00 else sc_lhl=GetFormatNumber(sc_lhl,2)
		g_docObj.Replace "{绿化率}",sc_lhl,0
End Function


Function GetAreaList(byval lhlx1,byval mc,areaList,count,sumJrmj1)
		if  mc ="'单块绿地'" then   strCondition="lhlx in ("&lhlx1&") and " else strCondition=""
		code="9470103":tableType =2
		strTableName=SSProcess.GetCodeAttrTableName(code,tableType)
		listCount=GetProjectTableList (strTableName,"bh,sum(lhmj)"," "&strCondition&" MC="&mc&"  group by bh","SpatialData",tableType,list,fieldCount)
		count=0:redim areaList(count)
		for i= 0 to listCount-1
			sumArea=0
			bh=list(i,0):sumLhmj=list(i,1):sumLhmj=GetFormatNumber(sumLhmj,2)'绿地 绿化面积合计
			'休憩场所 绿化面积合计
			'listCount1=GetProjectTableList (strTableName,"bh,sum(lhmj)","lhlx in ('休憩场所') and MC="&mc&" and bh='"&bh&"'  group by bh","SpatialData",tableType,list1,fieldCount)
			'xSumLhmj= list1(0,1):xSumLhmj=GetFormatNumber(xSumLhmj,2)
			'当前编号下的 绿化面积总计
			sumArea=cdbl(sumLhmj):sumArea=GetFormatNumber(sumArea,2)
			sumJrmj1=sumJrmj1+cdbl(sumArea)
			
			'计算当前编号下【绿地】的图形面积总和，并统计其涉及的绿化编号
			listCount2=GetProjectTableList (strTableName,strTableName&".id,lhbh,lhlx","lhlx in ("&lhlx1&") and MC="&mc&" and bh='"&bh&"'  order by lhbh asc","SpatialData",tableType,list2,fieldCount)
			allLhbh="":sumObjArea=0
			for j= 0 to listCount2-1
				objid=list2(j,0):lhbh=list2(j,1):lhlx=list2(j,2)
				objArea=SSProcess.GetObjectAttr (objid,"SSObj_Area")
				sumObjArea=sumObjArea+cdbl(objArea)
				if allLhbh="" then allLhbh=lhbh else allLhbh=allLhbh&","&lhbh
			next
			allLhbh=SplitNoteStr( allLhbh, 10,  chr(10))
			allLhbh=replace(allLhbh,chr(10),"")
			bhlist=split(allLhbh,",")
			for j=0 to ubound(bhlist)
				for j1=0 to ubound(bhlist)-1
					num1=mid(bhlist(j1),2,len(bhlist(j1))-1):num2=mid(bhlist(j1+1),2,len(bhlist(j1+1))-1)
					if isnumeric(num1)=true and isnumeric(num2) =true then 
						if cdbl(num1)>cdbl(num2) then 
							temp=bhlist(j1)
							bhlist(j1)=bhlist(j1+1)
							bhlist(j1+1)=temp
						end if 
					end if 
				next
			next
			str=""
			for j=0 to ubound(bhlist)
				if str="" then
					str=bhlist(j)
				else
					str=str&","&bhlist(j)
				end if
			next
			allLhbh=str
			sumObjArea=GetFormatNumber(sumObjArea,2)
			if lhlx="休憩场所" then
				cellValues=bh&"||"&"休憩场所"&"||"&allLhbh&"||"&sumObjArea&"||"&sumArea&"||"&""
			else
				cellValues=bh&"||"&"绿地"&"||"&allLhbh&"||"&sumObjArea&"||"&sumArea&"||"&""
			end if
			redim preserve areaList(count):areaList(count)=cellValues:count=count+1
			
			'计算当前编号下【休憩场所】的图形面积总和，并统计其涉及的绿化编号
			'listCount3=GetProjectTableList (strTableName,strTableName&".id,lhbh","lhlx in ('休憩场所') and MC="&mc&" and bh='"&bh&"'  order by lhbh asc","SpatialData",tableType,list3,fieldCount)
			'allLhbh1="":sumObjArea1=0
			'for j= 0 to listCount3-1
				'objid=list3(j,0):lhbh=list3(j,1)
				'objArea=SSProcess.GetObjectAttr (objid,"SSObj_Area")
				'sumObjArea1=sumObjArea1+cdbl(objArea)
				'if allLhbh1="" then allLhbh1=lhbh else allLhbh1=allLhbh1&","&lhbh
			'next
			'allLhbh1=SplitNoteStr( allLhbh1, 10,  chr(10))
			'sumObjArea1=GetFormatNumber(sumObjArea1,2)
			'cellValues=""&"||"&"休憩场所"&"||"&allLhbh1&"||"&sumObjArea1&"||"&"0"&"||"&""
			'redim preserve areaList(count):areaList(count)=cellValues:count=count+1
		next
		sumJrmj1=GetFormatNumber(sumJrmj1,2)
End Function


'//注记自动分割换行
'// allNoteStr  注记内容
'// splitCount  注记分割个数
Function SplitNoteStr(byval allNoteStr,byval splitCount,byval  splitMark)
		'//固定换行字符串用 || 分割
		allOutputStr=""
		noteStrList=split(allNoteStr,"||")
		for i= 0 to ubound(noteStrList)
			noteStr=noteStrList(i)
			noteStrLen=len(noteStr)
			count=0
			for j= 0 to noteStrLen
				if 	j  mod splitCount=0  then 
					count=count+1
				end if 
			next
			
			singleOutputStr=""
			for ii= 0 to count-1
				outputStr= mid(noteStr, (ii*splitCount)+1,splitCount)
				if  outputStr<>"" then  if singleOutputStr="" then singleOutputStr=outputStr else singleOutputStr=singleOutputStr&splitMark& outputStr
			next
			if allOutputStr="" then allOutputStr=singleOutputStr else  allOutputStr=allOutputStr&splitMark&singleOutputStr
		next
		SplitNoteStr=allOutputStr
End Function

'//插入 立面图
Function OutputTableLI(byval tableIndex,byval fileName,byval row,byval col)
		g_docObj.MoveToTable tableIndex,false
		
		'查找对应wmf文件并插入word报告中
		dim imageList(10000):listCount=0
		filePath=SSProcess.GetSysPathName (4)
		GetAllFiles filePath,"bmp",listCount,imageList
		insertImageFile=""
		for i= 0 to listCount-1
			imageFile=imageList(i)
			name=GetFileName(imageFile)
			extensionName=GetFileExtensionName(imageFile)
			name=replace(name,"."&extensionName,"")
			if name=fileName then insertImageFile=imageFile:	exit for 
		next
		'if FileExists( insertImageFile)  =true then   g_docObj.SetCellImageEx2 tableIndex,  row, col, 0,  insertImageFile, 0, 0, false

		if FileExists( insertImageFile)  =true then   g_docObj.SetCellImageEx tableIndex,  row, col, 0,  insertImageFile, 650, 200, false


End Function

'//插入 成果图
Function OutputTable6(byval tableIndex,byval fileName,byval code,byval row,byval col)
		g_docObj.MoveToTable tableIndex,false
		
		'查找成果图edb文件
		accessName=SSProcess.GetProjectFileName
		filePath=replace(accessName,".edb","")&"\"
		dim edbList(10000):listCount=0
		GetAllFiles filePath,"edb",listCount,edbList

		outEdbPath=""
		for i= 0 to listCount-1
			edbPath=edbList(i)
			if instr(edbPath,fileName)>0 and instr(fileName,"bak")=0 then 
				outEdbPath=edbPath
				exit for 
			end if 
		next
		if FileExists(outEdbPath)  =false then Exit Function
		'DeleteAllImage

		'打开edb文件并按图廓范围打印wmf
		bRes=SSProcess.OpenDatabase (outEdbPath)
		if bRes=1 then 
			PrintImage code,fileName, ShapeHeight, ShapeWidth,daYZZ
			SSProcess.CloseDatabase()
		end if 
		'查找对应wmf文件并插入word报告中
		dim imageList(10000):listCount=0
		filePath=SSProcess.GetSysPathName (4)
		GetAllFiles filePath,"bmp",listCount,imageList
		insertImageFile=""
		for i= 0 to listCount-1
			imageFile=imageList(i)
			name=GetFileName(imageFile)
			extensionName=GetFileExtensionName(imageFile)
			name=replace(name,"."&extensionName,"")
			if name=fileName then insertImageFile=imageFile:	exit for 
		next
		if FileExists( insertImageFile)  =true then   g_docObj.SetCellImageEx2 tableIndex,  row, col, 0,  insertImageFile, 0, 0, false
End Function

'//插入 成果图
Function OutputBook(byval bookmark,byval fileName,byval code,byval SectionIndex)
		dim imageList(10000):listCount1=0
		g_docObj.MoveToBookmark	bookmark
		
		'查找成果图edb文件
		accessName=SSProcess.GetProjectFileName
		filePath=replace(accessName,".edb","")&"\"
		dim edbList(10000):listCount=0
		GetAllFiles filePath,"edb",listCount,edbList
		outEdbPath=""		
		for j= 0 to listCount-1
			DeleteAllImage
			edbPath=edbList(j)
			if instr(edbPath,fileName)>0 and instr(fileName,"bak")=0 then 
				outEdbPath=edbPath
				if FileExists(outEdbPath)  =false then Exit Function				
				'打开edb文件并按图廓范围打印wmf
				bRes=SSProcess.OpenDatabase (outEdbPath)
				if bRes=1 then 
					PrintImage code,fileName, ShapeHeight, ShapeWidth,daYZZ
					SSProcess.CloseDatabase()
				end if 
				'查找对应wmf文件并插入word报告中
				filePath=SSProcess.GetSysPathName (4)
				GetAllFiles filePath,"bmp",listCount1,imageList
				insertImageFile=""
				for i= 0 to listCount1-1
					imageFile=imageList(i)
					name=GetFileName(imageFile)
					extensionName=GetFileExtensionName(imageFile)
					name=replace(name,"."&extensionName,"")
					nameNumber = replace(name,fileName,"")
					if instr(name,fileName)>0 then 
						insertImageFile=imageFile
						if FileExists( insertImageFile)  =true then 
							'RES = g_docObj.InsertImage (insertImageFile,ShapeWidth,ShapeHeight,0)
							If daYZZ = "A4横向" then
								paperSize =  1
								orientation=2
								pageWidth = -1 : pageHeight = -1
								H=17.1: W=24.2
								width = 26.345*W
								height = 26.345*H
								'设置纸张的大小
								leftMargin=20'毫米
								rightMargin=20
								topMargin=7
								bottomMargin=7
							elseif daYZZ = "A4纵向" then
								paperSize =  1
								orientation=1
								pageWidth = -1: pageHeight = -1
								'设置宽高
								H=26.8: W=21.8
								width = 20.245*W
								height = 10.345*H
								'设置纸张的大小
								leftMargin=10'毫米
								rightMargin=10
								topMargin=10
								bottomMargin=10
							elseif daYZZ  = "A3纵向" then
								paperSize = 0
								orientation=1
								pageWidth = -1 : pageHeight = -1
								H=37.2: W=26.3
								width = 28.345*W
								height = 28.345*H
								'设置纸张的大小
								leftMargin=10'毫米
								rightMargin=10
								topMargin=10
								bottomMargin=10
							elseif daYZZ = "A3横向" then
								paperSize = 0
								orientation=2
								pageWidth = -1 : pageHeight = -1
								H=25.8: W=36.5
								width = 28.345*W
								height = 28.345*H
								'设置纸张的大小
								leftMargin=10'毫米
								rightMargin=10
								topMargin=10
								bottomMargin=10
							elseif daYZZ = "500*500" then
								paperSize =  1
								orientation=1
								pageWidth = 500: pageHeight = 500
								'设置宽高
								H=45.04: W=45.01
								width =30.245*W
								height =28.345*H
								'设置纸张的大小
								leftMargin=10'毫米50
								rightMargin=10
								topMargin=10
								bottomMargin=10
							end if
							if SectionIndex<>"" then
								g_docObj.SectionPageSetup SectionIndex, paperSize, orientation, pageWidth, pageHeight, leftMargin, rightMargin, topMargin, bottomMargin
							end if
						'水平相对位置模式（wrapType非0时起作用） Margin = 0, Page = 1, Column = 2, Default = 2, Character = 3, LeftMargin = 4, RightMargin = 5, InsideMargin = 6, OutsideMargin = 7
							horzPos = 0
							left0 = 0
							'垂直位置相对模式（wrapType非0时起作用） Margin = 0,  TableDefault = 0,  Page = 1,  Paragraph = 2, TextFrameDefault = 2,  Line = 3,  TopMargin = 4,   BottomMargin = 5,  InsideMargin = 6,  OutsideMargin = 7
							vertPos =  0
							top0 = 3
							'图像环绕方式 Inline = 0 嵌入,    TopBottom = 1 上下,   Square = 2 四周,   None = 3 浮于文字上方,    Tight = 4 紧密,  Through = 5 穿越
							wrapType =  0
							'旋转角度
							rotation = 0
							g_docObj.InsertImageEx insertImageFile, horzPos, left0, vertPos, top0, ShapeWidth,ShapeHeight,  wrapType, rotation
						end if
					end if
				next
				'if FileExists( insertImageFile)  =true then   g_docObj.SetCellImageEx2 tableIndex,  0, 0, 0,  insertImageFile, 0, 0, false
			end if 
		next


End Function



'//输出 停车位序号一览表
Function OutputTable7(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		codes="9460003,9460013,9460023,9460033,9460043,9460053"
		cwList=array("小型车位","微型车位","无障碍车位","其他车位")
		
		'获取车位信息
		cellCount=0:redim  cellList(cellCount)
		for i= 0 to ubound(cwList)
			strCw=cwList(i)
			cwbhs=""
			SSProcess.PushUndoMark
			SSProcess.ClearSelection
			SSProcess.ClearSelectCondition
			SSProcess.SetSelectCondition "SSObj_Code", "==", codes
			SSProcess.SelectFilter
			geoCount = SSProcess.GetSelGeoCount()
			If geoCount > 0 Then
				For j=0 To geoCount-1 
					objID = SSProcess.GetSelGeoValue(j, "SSObj_ID") 
					cwlx=SSProcess.GetObjectAttr (objID, "[cwzlx]")
					cwbh=SSProcess.GetObjectAttr (objID, "[cwbh]")
					if strCw=cwlx  and instr(cwbhs,"'"&cwbh&"'")=0  and  cwbh<>""  then 
						if cwbhs="" then cwbhs="'"&cwbh&"'" else cwbhs=cwbhs&","&"'"&cwbh&"'"
					end if 
				Next
			End If
			cwbhs=replace(cwbhs,"'","")
			redim preserve  cellList(cellCount):cellList(cellCount)=cwbhs
			cellCount=cellCount+1
		next

		'数组按 lhbh 从小到大冒泡排序
		for i= 0 to cellCount-1
			cellValueList=split(cellList(i),",")
			for j= 0 to ubound(cellValueList)
				for j1=0 to ubound(cellValueList)-1
					num1=cellValueList(j1)
					num2=cellValueList(j1+1)
					if isnumeric(num1)=true and isnumeric(num2) =true then 
						if cdbl(num1)>cdbl(num2) then 
							temp=cellValueList(j1)
							cellValueList(j1)=cellValueList(j1+1)
							cellValueList(j1+1)=temp
						end if 
					end if 
				next
			next
			str=""
			for j=0 to ubound(cellValueList)
				if str="" then
					str=cellValueList(j)
				else
					str=str&","&cellValueList(j)
				end if
			next 
			cellList(i)=str
		next

		
		'填充表单元格
		startRow=1:startCol=1
		for i= 0 to cellCount-1
			g_docObj.SetCellText tableIndex,startRow,startCol,cellList(i),true,false
			startRow=startRow+1
		next
End Function


'//输出 总平面布局测量表
Function OutputTable8(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		
		iniRow=2:iniCol=1
		startRow=iniRow:startCol=iniCol
		
		'****************************************填充 消防车道 单元格
		GetValueList cellList1,cellCount1,"9430033",2,"XS,ZXJKCC,ZXZWBJ,ZDPD,JLJZWQJLZXZ,JLJZWQJLZDZ"
		if cellCount1>1 then  g_docObj.CloneTableRow tableIndex,  startRow, 1,cellCount1-1, false
		for i= 0 to cellCount1-1
			startCol=iniCol
			cellValue=cellList1(i)
			cellValueList=split(cellValue,"||")
			for j= 0 to ubound(cellValueList)
				g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
		next
		
		'****************************************填充 消防通道 单元格
		GetValueList cellList2,cellCount2,"9430023",2,"MC,JKSJZ,JKSCZ,JGYXZXZ,JGSCZ"
		if cellCount1>1  then  startRow=iniRow+3+ cellCount1-1 else startRow=iniRow+3
		if cellCount2>1 then  g_docObj.CloneTableRow tableIndex,  startRow, 1,cellCount2-1, false
		for i= 0 to cellCount2-1
			startCol=iniCol
			cellValue=cellList2(i)
			cellValueList=split(cellValue,"||")
			for j= 0 to ubound(cellValueList)
				g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
		next
		
		'****************************************填充 消防登高操作场地 单元格
		GetValueList cellList3,cellCount3,"9430013",2,"MC,CCSJZ,CCSCZ,JLWQZXZ,JLWQZDZ,PD"
		if cellCount2=0  then 
			startRow=startRow+3
		elseif  cellCount2=1 then 
			startRow=startRow+2
		else
			startRow=startRow+cellCount2-1
		end if
		if 	cellCount3>1 then 	g_docObj.CloneTableRow tableIndex,  startRow, 1,cellCount3-1, false
		for i= 0 to cellCount3-1
			startCol=iniCol+2
			if i=0 then 
				g_docObj.SetCellText tableIndex,startRow,iniCol+1,cellCount3,true,false
				if cellCount3>1 then  
					g_docObj.MergeCell tableIndex,  startRow,  iniCol+1,  startRow+cellCount3-1,  iniCol+1,false
					g_docObj.MergeCell tableIndex,  startRow,  iniCol,  startRow+cellCount3-1,  iniCol,false
				end if 
			end if 
			cellValue=cellList3(i)
			cellValueList=split(cellValue,"||")
			for j= 0 to ubound(cellValueList)
				g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
		next
		
		strTableName=SSProcess.GetCodeAttrTableName("9430033",2)
		listCount=GetProjectTableList (strTableName,strTableName&".id,QKSM,JSZDZ","","SpatialData",tableType,list,fieldCount)
		allQKSM="" :allJSZDZ=""
		for i= 0 to listCount-1
			QKSM=list(i,1):JSZDZ=list(i,2)
			if QKSM<>"" and QKSM<>"*" then   if allQKSM="" then  allQKSM=QKSM else allQKSM=allQKSM&","&QKSM
			if JSZDZ<>"" and JSZDZ<>"*" then   if allJSZDZ="" then  allJSZDZ=JSZDZ else allJSZDZ=allJSZDZ&","&JSZDZ
		next
		g_docObj.Replace "{QKSM}",allQKSM,0
		g_docObj.Replace "{JSZDZ}",allJSZDZ,0
End Function


Function GetValueList(cellList,cellCount,byval code,byval tableType,byval fields)
		strTableName=SSProcess.GetCodeAttrTableName(code,tableType)
		listCount=GetProjectTableList (strTableName,strTableName&".id,"&fields&"","","SpatialData",tableType,list,fieldCount)
		cellCount=0:redim cellList(cellCount)
		for i= 0 to listCount-1
			cellValue=""
			objid=list(i,0)
			objCode=SSProcess.GetObjectAttr(objid,"SSObj_Code")
			if objCode=code then 
				for j= 0 to fieldCount-1
					if j<>0 then 
						value=list(i,j)
						if j<>1 then   value=GetFormatNumber(value,2)
						if j=1 then cellValue=value  else cellValue=cellValue&"||"&value
					end if 
				next
				redim preserve cellList(cellCount):cellList(cellCount)=cellValue:cellCount=cellCount+1
			end if 
		next
End Function


'//输出 地下管线测量成果表
Function	OutputTable9(byval tableIndex)
		strGxdTableName="GD_管点基本属性表":gxdTableType=0
		strGxxTableName="GX_管线基本属性表":gxxTableType=1
		
		mdbName = SSProcess.GetProjectFileName  
		SSProcess.OpenAccessMdb mdbName 

		initableIndex=tableIndex
		layers="给水,排水,电力,通信,热力,燃气,工业,其他,综合管廊（沟）":layersList=split(layers,",")
		SSProcess.ClearSelection
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_LayerName", "==", layers
		SSProcess.SelectFilter
		geocount = SSProcess.GetSelGeoCount()
		if geocount>0 then
			SSProcess.CreateMapFrame()
			frameCount = SSProcess.GetMapFrameCount()
			for i = 0  to frameCount-1
				SSProcess.GetMapFrameCenterPoint i, x, y
				SSProcess.SetCurMapFrame x, y, 0, ""
				frameID = SSProcess.GetCurMapFrame()
				mapNumber = SSProcess.GetObjectAttr( CLng(frameID), "[MapNumber]")
				ids=SSProcess.SearchInPolyObjIDs(CLng(frameID), 0, "", 1, 1, 1):idsList=split(ids,",")
				redim PIPEid(ubound(layersList))
				for i1 =0 to ubound(idsList)
					idslayername = SSProcess.GetObjectAttr( idsList(i1), "SSObj_LayerName")
					for j = 0 to ubound(layersList)
						if idslayername = layersList(j) then
							if PIPEid(j)="" then PIPEid(j)=idsList(i1) else PIPEid(j)=PIPEid(j)&","&idsList(i1)
						end if
					next
				next

				'复制表格
				tableclonecount = 0
				strPIPEid = ""
				for ii = 0 to ubound(layersList)
					if PIPEid(ii)<>"" then 
						if strPIPEid= "" then
							strPIPEid = PIPEid(ii)
						else
							strPIPEid = strPIPEid&";"&PIPEid(ii)
						end if
						tableclonecount = tableclonecount+1
					end if
				next
				if 	tableclonecount>0 then	
					for i2 = 0 to tableclonecount-1
						cloneres = g_docObj.CloneTable(tableIndex,1,0,false)
					next
					tableIndex = tableIndex+1
				end if
				atempPIPEid = split(strPIPEid,";")
				strTZ = "直通,弯头,三通,四通,多通,变径,变材,进水口,出水口,盖堵,预留口,非普,井边点,进水口,出水口,分支,变坡点"
				strFSW="检修井,阀门井,消防井,水表（井）,排气阀（井）,排污阀（井）,消防栓,阀门,污水检修井,雨水检修井,雨篦,污篦,溢流井,阀门井,跌水井,通风井,冲洗井,沉泥井,渗水井,出气井,水封井,阀门,倒虹井,阀门井,阀门,凝水缸,调压箱,排气装置,计量表,检修井,阴极保护桩,储气柜,检修井,阀门井,阀门,调压装置,凝水井,排气阀门,排污装置,压力计,出水口,上杆,变压器,检修井,控制柜,灯杆,线杆,上（下）杆,路灯,电线塔,电线架（双杆）,检修井,人孔,手孔,分线箱,线杆,上（下）杆,接线箱,电话亭,电杆,检修井,排污装置,阀门井,阴极保护桩,人员出入口,逃生口,排风口,吊装口,进风口,管线分支口"



				'拆分每个图幅里每个管线图层的id
				for iii = 0 to tableclonecount-1
					g_docObj.MoveToTable tableIndex+iii,false
					gxpointid = split(atempPIPEid(iii),",")
					startRow = 3
					for kk = 0 to ubound(gxpointid)
						sql = "select GD_管点基本属性表.PID,GD_管点基本属性表.PhyCode,GD_管点基本属性表.Ctype,GD_管点基本属性表.FSW,GD_管点基本属性表.X,GD_管点基本属性表.Y,GD_管点基本属性表.PPZ,GD_管点基本属性表.WellDepth from GD_管点基本属性表 INNER JOIN GeoPointTB ON GD_管点基本属性表.ID = GeoPointTB.ID where ([GeoPointTB].[Mark] Mod 2)<>0 AND GD_管点基本属性表.ID = "&gxpointid(kk)
						GetSQLRecordAll mdbName, sql, arGXpointRecord, nGXpointCount
						pipetablelayername = SSProcess.GetObjectAttr( gxpointid(kk), "SSObj_LayerName")
						g_docObj.SetCellText tableIndex+iii,0,1,"管线类别："&pipetablelayername,true,false
						g_docObj.SetCellText tableIndex+iii,0,2,"所在图幅："&mapNumber,true,false
						for k1 = 0 to nGXpointCount-1
							artemp = split(arGXpointRecord(k1),",")
							'管点是管线的起点
							sql1 = "select GX_管线基本属性表.EnodeID,GX_管线基本属性表.LEMS,GX_管线基本属性表.LBTG,GX_管线基本属性表.LETG,GX_管线基本属性表.PWidHt,GX_管线基本属性表.Material,GX_管线基本属性表.VentNum,GX_管线基本属性表.Number,GX_管线基本属性表.Pressure,GX_管线基本属性表.Voltage,GX_管线基本属性表.LayMethod,GX_管线基本属性表.LayDate,GX_管线基本属性表.Source,GX_管线基本属性表.ID from GX_管线基本属性表 INNER JOIN GeoLineTB ON GX_管线基本属性表.ID = GeoLineTB.ID where ([GeoLineTB].[Mark] Mod 2)<>0 and GX_管线基本属性表.SnodeID = '"&artemp(0)&"'"
							GetSQLRecordAll mdbName, sql1, arGXlineRecord, nGXlineCount
							for k2 = 0 to nGXlineCount-1
								startCol = 0
								startCol1 = 4
								'管点属性
								for kk1 = 0 to ubound(artemp)
									if kk1=0 and len(artemp(kk1))>6 then artemp(kk1)=mid(artemp(kk1),len(artemp(kk1))-14+1,2)+right(artemp(kk1),4)
									if kk1 = 4 then 
										g_docObj.SetCellText tableIndex+iii,startRow,startCol+1,Round(artemp(kk1),3),true,false:startCol = startCol+2
									elseif kk1 = 5 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol,Round(artemp(kk1),3),true,false:startCol = startCol+1
									elseif kk1 = 7 then 
										g_docObj.SetCellText tableIndex+iii,startRow,startCol,artemp(kk1),true,false
									else
										g_docObj.SetCellText tableIndex+iii,startRow,startCol,artemp(kk1),true,false:startCol = startCol+1
									end if
								next
								'管线属性
								artempline = split(arGXlineRecord(k2),",")
								strPressure = "给水、燃气、热力、工业"
								for kk2 = 0 to ubound(artempline)
									if kk2 =0  and len(artempline(kk2))>6 then artempline(kk2)=mid(artempline(kk2),len(artempline(kk2))-14+1,2)+right(artempline(kk2),4)
									linelayername = SSProcess.GetObjectAttr( artempline(13), "SSObj_LayerName")
									if kk2 = 0 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk2),true,false:startCol1 = startCol1+4
									elseif kk2 = 1 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk2),true,false:startCol1 = startCol1+2
									elseif kk2=6 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk2)&"/"&artempline(kk2+1),true,false:startCol1 = startCol1+1
									elseif kk2 = 7 then
										startCol1 = startCol1
									elseif kk2=8 then
										if replace(strPressure,linelayername,"")<>strPressure then
											g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(8),true,false:startCol1 = startCol1+1
										else
											startCol1 = startCol1
										end if
									elseif kk2 = 9   then
										if  linelayername = "电力" then
											g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(9),true,false:startCol1 = startCol1+1
										else 
											startCol1 = startCol1
										end if
									elseif kk2<13 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk2),true,false:startCol1 = startCol1+1
									end if
								next
								'起点的行数
								startRow = startRow+1
								g_docObj.CloneTableRow tableIndex+iii,startRow,1,1,false
							next

							'管线是管线的终点
							sql1 = "select GX_管线基本属性表.SnodeID,GX_管线基本属性表.LBMS,GX_管线基本属性表.LBTG,GX_管线基本属性表.LETG,GX_管线基本属性表.PWidHt,GX_管线基本属性表.Material,GX_管线基本属性表.VentNum,GX_管线基本属性表.Number,GX_管线基本属性表.Pressure,GX_管线基本属性表.Voltage,GX_管线基本属性表.LayMethod,GX_管线基本属性表.LayDate,GX_管线基本属性表.Source,GX_管线基本属性表.ID from GX_管线基本属性表 INNER JOIN GeoLineTB ON GX_管线基本属性表.ID = GeoLineTB.ID where ([GeoLineTB].[Mark] Mod 2)<>0 and GX_管线基本属性表.EnodeID = '"&artemp(0)&"'"
							GetSQLRecordAll mdbName, sql1, arGXlineRecord, nGXlineCount
							for k22 = 0 to nGXlineCount-1
								startCol = 0
								startCol1 = 4
								'管点属性
								for kk1 = 0 to ubound(artemp)
									if kk1=0 and len(artemp(kk1))>6 then artemp(kk1)=mid(artemp(kk1),len(artemp(kk1))-14+1,2)+right(artemp(kk1),4)
									if kk1 = 4 then 
										g_docObj.SetCellText tableIndex+iii,startRow,startCol+1,Round(artemp(kk1),3),true,false:startCol = startCol+2
									elseif kk1 = 5 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol,Round(artemp(kk1),3),true,false:startCol = startCol+1
									elseif kk1 = 7 then 
										g_docObj.SetCellText tableIndex+iii,startRow,startCol,artemp(kk1),true,false
									else
										g_docObj.SetCellText tableIndex+iii,startRow,startCol,artemp(kk1),true,false:startCol = startCol+1
									end if
								next
								'管线属性
								artempline = split(arGXlineRecord(k22),",")
								strPressure = "给水、燃气、热力、工业"
								for kk22 = 0 to ubound(artempline)
									if kk22 =0 and len(artempline(kk22))>6 then artempline(kk22)=mid(artempline(kk22),len(artempline(kk22))-14+1,2)+right(artempline(kk22),4)
									linelayername = SSProcess.GetObjectAttr( artempline(13), "SSObj_LayerName")
									if kk22 = 0 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk22),true,false:startCol1 = startCol1+4
									elseif kk22 = 1 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk22),true,false:startCol1 = startCol1+2
									elseif kk22=6 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk22)&"/"&artempline(kk22+1),true,false:startCol1 = startCol1+1
									elseif kk22 = 7 then
										startCol1 = startCol1
									elseif kk22=8 then
										if replace(strPressure,linelayername,"")<>strPressure then
											g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(8),true,false:startCol1 = startCol1+1
										else
											startCol1 = startCol1
										end if
									elseif kk22 = 9   then
										if  linelayername = "电力" then
											g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(9),true,false:startCol1 = startCol1+1
										else 
											startCol1 = startCol1
										end if
									elseif kk22<13 then
										g_docObj.SetCellText tableIndex+iii,startRow,startCol1,artempline(kk22),true,false:startCol1 = startCol1+1
									end if
								next
								g_docObj.CloneTableRow tableIndex+iii,startRow,1,1,false
								'终点的行数
								startRow = startRow+1
							next
				
						next
					next
					allCount=g_docObj.GetTableRowCount (tableIndex,false)
					g_docObj.DeleteRow tableIndex,allCount-2,false
				next

			next
			if tableclonecount>0 then g_docObj.DeleteTable	initableIndex,false
			SSProcess.FreeMapFrame
			SSProcess.CloseAccessMdb mdbName 
		end if
End Function




'//创建图幅缓存
Function CreateMapFrame(mark)
		mark="管线涉及图幅缓存"
		SSProcess.PushUndoMark 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_DataMark", "==",mark
		SSProcess.SelectFilter
		SSProcess.DeleteSelectionObj
		
		'创建当前图幅
		maxID=SSProcess.GetGeoMaxID()
		SSProcess.CreateMapFrame
		frameCount = SSProcess.GetMapFrameCount()
		For i=0 To frameCount-1
			SSProcess.CreateOneMapFrame i, 2
		Next
		SSProcess.FreeMapFrame
		
		
		SSProcess.PushUndoMark 
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_ID", ">",maxID
		SSProcess.SelectFilter
		SSProcess.ChangeSelectionObjAttr "SSObj_DataMark", mark	
End Function


'//获取图幅内的地物列表
Function GetTableList(mark,tableList,tableCount)
		tableCount=0:redim tableList(tableCount)
		layers="给水,排水,电力,通信,热力,燃气,工业,其他,综合管廊（沟）":layersList=split(layers,",")
		
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_DataMark", "==",mark
		SSProcess.SelectFilter
		geoCount = SSProcess.GetSelGeoCount()
		If geoCount > 0 Then
			For i=0 To geoCount-1 
				id = SSProcess.GetSelGeoValue(i, "SSObj_ID") 
				ids=SSProcess.SearchInPolyObjIDs(id, 10, "", 0, 1, 1):idsList=split(ids,",")
				innerids=""
				'筛选图幅内符合条件的地物
				for j= 0 to ubound(idsList)
					innerid=idsList(j)
					innerLayer=SSProcess.GetObjectAttr (innerid,"SSObj_LayerName")
					innerType=SSProcess.GetObjectAttr (innerid,"SSObj_Type")
					if innerType="POINT" then 
						isLayer=false
						for m= 0 to ubound(layersList)
							if innerLayer=layersList(m)   then isLayer=true :Exit For
						next
						if isLayer=true then 
							if innerids="" then innerids=innerid else innerids=innerids&","&innerid
						end if 
					end if 
				next
				if innerids<>"" then 
					redim preserve  tableList(tableCount):tableList(tableCount)=innerids
					tableCount=tableCount+1
				end if 
			Next
		End If
End Function


'//输出 建（构）筑物竣工标高测量一览表
Function OutputTable10(byval tableIndex,byref tablenodecount)
		g_docObj.MoveToTable tableIndex,false 
		writeRowCount=4:copyCount=0
		ydhxTableName="JG_用地红线信息属性表"
		strTableName="JG_建设工程建筑单体信息属性表"
		exCondition=""&strTableName&".id>0"
		cellCount=0:redim cellList(cellCount)
		fields="JianZWMC,SWDPBG,SNDPBG,DCNDPBG,JZZGDBG,BeiZ,ID_ZRZ"		
		listCount=GetProjectTableList (strTableName,strTableName&".id,"&fields&"",exCondition,"","",list,fieldCount)

		for i= 0 to listCount-1
			cellValue="":ID_ZRZ=list(i,7):LD=list(i,1)
			listCount2=GetProjectTableList ("FC_自然幢信息属性表","FWJG,DSCS"," LD='"&LD&"'","SpatialData","2",list2,fieldCount2)
			if listCount2=1 then 
				FWJG=list2(0,0)
				value2=FWJG&"-"&FWJG&list2(0,1)

			end if
			for j= 0 to fieldCount-2
				value=list(i,j)
				if value="" then value="/"
				if j=6 then value=value2
				if j<>0 then 
					if j=1 then cellValue=value  else cellValue=cellValue&"||"&value
				end if 
			next
			listCount1=GetProjectTableList ("JG_建筑物屋顶高度信息表","sjgc,scgc"," WZ='建筑最高点' and ID_ZRZ='"&ID_ZRZ&"'","","",list1,fieldCount1)
			if listCount1=1 then value1=list1(0,0)&"||"&list1(0,1)

			cellValue=cellValue&"||"&value1
			redim preserve cellList(cellCount):cellList(cellCount)=cellCount+1&"||"&cellValue:cellCount=cellCount+1
			if i>0 and i mod writeRowCount=0 then copyCount=copyCount+1
		next

		iniRow=4:iniCol=0
		startRow=iniRow:startCol=iniCol

		'根据数据个数复制表格
		for i=0 to copyCount-1
			g_docObj.CloneTable  tableIndex, 1,0,false
		next

		mapindex=1:mapmark=true
		if cellCount MOD writeRowCount=0 then clonecount=cellCount/writeRowCount else clonecount=int(cellCount/writeRowCount)+1
		for i= 0 to cellCount-1
			if i>0 and i mod writeRowCount=0 then startRow=iniRow :tableIndex=tableIndex+1:mapmark=true
			startCol=iniCol
			cellValue=cellList(i)
			cellValueList=split(cellValue,"||")
			for j= 0to ubound(cellValueList)
				g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
			if mapmark=true then
				if mapindex<clonecount then
					g_docObj.MoveToTablePreviousNode tableIndex+1,false
					g_docObj.InsertBreak 5
					tablenodecount=tablenodecount+1
				end if
				OutputTableLI tableIndex, "立面图"&mapindex,8,0
				mapindex=mapindex+1
				mapmark=false
			end if
		next
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

function  RFFHDYCount()
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9450013"
	SSProcess.SelectFilter
	geocount = SSProcess.GetSelGeoCount()
	RFFHDYCount = geocount
end function


'//输出 主要经济技术指标比对表
Function OutputTable11(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		cellCount=0:redim cellList(cellCount)
		'**************************************************************总用地面积
		ydhxTableName="JG_建设工程规划许可证信息属性表"
		fields="GuiHSPZYDMJ"
		listCount=GetProjectTableList ("JG_建设工程规划许可证信息属性表","GuiHSPZYDMJ","ID>0","","",list,fieldCount)
		if listCount=1 then gh_YongDMJ = list(0,0)
		gh_YongDMJ=GetFormatNumber(gh_YongDMJ,2)'规划-总用地面积
		listCount=GetProjectTableList ("JGSCHZXX","YDMJ","ID>0","","",list,fieldCount)
		if listCount=1 then sc_YongDMJ = list(0,0)
		if sc_YongDMJ<>"" then sc_YongDMJ = GetFormatNumber(sc_YongDMJ,2)
		GetSubArea cellList,cellCount, sc_YongDMJ, gh_YongDMJ,2,1

		'**************************************************************总建筑面积
		zrzCount=GetProjectTableList ("JGSCHZXX","JZMJ","ID>0","","",zrzList,fieldCount)
		if zrzCount=1 then sc_SCJZMJ=zrzList(0,0)
		sc_SCJZMJ=GetFormatNumber(sc_SCJZMJ,2)'实测-总建筑面积
		
		ghxkTableName="JG_建设工程规划许可证信息属性表"
		'exCondition="YDHXGUID In (select YDHXGUID from "&ydhxTableName&"  INNER JOIN GeoLineTB ON "&ydhxTableName&".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0)"

		exCondition="ID>0"
		ghxkCount=GetProjectTableList (ghxkTableName,"sum(GuiHSPZJZMJ)",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then gh_SCJZMJ=ghxkList(0,0)
		gh_SCJZMJ=GetFormatNumber(gh_SCJZMJ,2)'规划-总建筑面积
		GetSubArea cellList,cellCount, sc_SCJZMJ, gh_SCJZMJ,2,1
		iniRow=1:iniCol=1
		startRow=iniRow:startCol=iniCol
		ghgnqTableName="JG_规划功能区属性表"
		'**************************************************************地上建筑面积
		listCount = GetProjectTableList ("JGSCHZXX","DSJZMJ"," ID>0","","",List,fieldCount)
		if listCount=1 then sc_DSJZMJ=List(0,0) 
		listCount = GetProjectTableList ("projectinfo"," value "," key='地上建筑面积(m2)' ","","",List,fieldCount)
		if listCount=1 then gh_DSJZMJ=List(0,0) 
		GetSubArea cellList,cellCount, sc_DSJZMJ, gh_DSJZMJ,2,2
		GetGnqAreaList cellList,cellCount, "地上",copyCount
		'复制地上功能区
		startRow=iniRow+2
		startCol=iniCol+1
		if copyCount>1 then 
			g_docObj.CloneTableRow tableIndex,  startRow+1, 1,copyCount-1, false
			g_docObj.MergeCell tableIndex,  startRow,  1,  startRow+copyCount,  1,false
		end if
		'**************************************************************地下建筑面积
		listCount = GetProjectTableList ("JGSCHZXX","DXJZMJ"," ID>0","","",List,fieldCount)
		if listCount=1 then sc_DSJZMJ=List(0,0) 
		listCount = GetProjectTableList ("projectinfo"," value "," key='地下建筑面积(m2)' ","","",List,fieldCount)
		if listCount=1 then gh_DSJZMJ=List(0,0) 
		GetSubArea cellList,cellCount, sc_DSJZMJ, gh_DSJZMJ,2,2
		
		startRow=startRow+copyCount+1
		GetGnqAreaList cellList,cellCount, "地下",copyCount1
		'复制地下功能区
		if copyCount1>1 then 
			g_docObj.CloneTableRow tableIndex,  startRow+1, 1,copyCount1-1, false
			g_docObj.MergeCell tableIndex,  startRow,  1,  startRow+copyCount1,  1,false
		end if
		if  (copyCount1+ copyCount)>0 then   g_docObj.MergeCell tableIndex,  iniRow+3,  0,  iniRow+3 + copyCount1+ copyCount,  0,false
		'**************************************************************建筑基底面积
		jdCount=GetProjectTableList ("JGSCHZXX","JZJDMJ"," ID>0 ","","",jdList,fieldCount)
		if jdCount=1 then sc_JDMJ=jdList(0,0)
		sc_JDMJ=GetFormatNumber(sc_JDMJ,2)'实测-建筑基底面积
		ghxkCount=GetProjectTableList (ghxkTableName,"sum(GuiHSPJDMJ),sum(GuiHSPRJL),sum(GuiHSPJZMD),sum(GuiHSPLHL),sum(ZpsJZMJ),sum(ScZZHS),sum(GhZZHS)",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then 
			gh_JDMJ=ghxkList(0,0):gh_GuiHSPRJL=ghxkList(0,1):gh_GuiHSPJZMD=ghxkList(0,2)
			gh_GuiHSPLHL=ghxkList(0,3):gh_ZpsJZMJ=ghxkList(0,4)
			ScZZHS=ghxkList(0,5):GhZZHS=ghxkList(0,6)
		end if 
		gh_JDMJ=GetFormatNumber(gh_JDMJ,2)'规划-建筑基底面积
		GetSubArea cellList,cellCount, sc_JDMJ, gh_JDMJ,2,1

		ldCount=GetProjectTableList ("JGSCHZXX","LDZMJ","ID>0","","",sclhmjList,fieldCount)
		if ldCount = 1 then sc_lhmj=sclhmjList(0,0)
		If gh_YongDMJ<>"" and gh_GuiHSPLHL<>"" then gh_lhmj = GetFormatNumber(gh_GuiHSPLHL*gh_YongDMJ*0.01,2) else gh_lhmj=""
		GetSubArea cellList,cellCount, sc_lhmj, gh_lhmj,2,1'绿地面积
		listCount=GetProjectTableList ("JGSCHZXX","RJV","ID>0","","",list,fieldCount)'容积率
		if listCount=1 then sc_Rjl = list(0,0)
		GetSubArea cellList,cellCount, sc_Rjl, gh_GuiHSPRJL,2,1'容积率
		listCount=GetProjectTableList ("JGSCHZXX","JZMD","ID>0","","",list,fieldCount)'建筑密度
		if listCount=1 then sc_Jzmd = list(0,0)
		GetSubArea cellList,cellCount, sc_Jzmd, gh_GuiHSPJZMD,2,1'建筑密度
		ldCount=GetProjectTableList ("JGSCHZXX","LVL","ID>0","","",sclhYdmjList,fieldCount)
		if ldCount = 1 then sc_lhl=sclhYdmjList(0,0)
		GetSubArea cellList,cellCount, sc_lhl, gh_GuiHSPLHL,2,1'绿化率
		listCount=GetProjectTableList ("JGSCHZXX","ZPSJZMJ","ID>0","","",list,fieldCount)'装配式建筑面积
		if listCount=1 then sc_ZpsJZMJ = list(0,0)
		'if gh_ZpsJZMJ = 0 then gh_ZpsJZMJ=""
		GetSubArea cellList,cellCount, sc_ZpsJZMJ, gh_ZpsJZMJ,2,1'装配式建筑面积
		cwTableName="CWSCXX"
		cwCount=GetProjectTableList ("JGSCHZXX","DSJDCWGS+DXJDCWGS,DSJDCWGS,DXJDCWGS","ID>0","","",cwList,fieldCount)
		if  cwCount=1 then    
			sc_ds_Jdcw=cwList(0,1):sc_dx_Jdcw=cwList(0,2)
			if sc_ds_Jdcw="" then sc_ds_Jdcw=0:if sc_dx_Jdcw="" then sc_dx_Jdcw=0
			sc_Jdcw=int(sc_ds_Jdcw)+int(sc_dx_Jdcw)
		end if

		ghcwTableName="CWGHXX"
		cwCount=GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX<>'非机动车位'","","",ghcwList,fieldCount)
		if  cwCount=1 then    gh_Jdcw=ghcwList(0,0):gh_ds_Jdcw=ghcwList(0,1):gh_dx_Jdcw=ghcwList(0,2)
		GetSubArea cellList,cellCount, sc_Jdcw, gh_Jdcw,0,1'机动车位
		GetSubArea cellList,cellCount, sc_ds_Jdcw, gh_ds_Jdcw,0,2'地上机动车位
		GetSubArea cellList,cellCount, sc_dx_Jdcw, gh_dx_Jdcw,0,2'地下机动车位
		GetSubArea cellList,cellCount, ScZZHS, GhZZHS,0,1'住宅户数

		cwCount=GetProjectTableList ("JGSCHZXX","DSFJDCWGS,DXFJDCWGS","ID>0","","",cwList,fieldCount)
		if  cwCount=1 then    
			DSFJDCWGS=cwList(0,0):DXFJDCWGS=cwList(0,1)
			if DSFJDCWGS="" then DSFJDCWGS=0:if DXFJDCWGS="" then DXFJDCWGS=0
			sc_Fjdcw=int(DSFJDCWGS)+int(DXFJDCWGS)
		end if
		ghcwCount=GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL)","CWLX='非机动车位'","","",ghcwList,fieldCount)
		if  ghcwCount=1 then    gh_Fjdcw=ghcwList(0,0)
		GetSubArea cellList,cellCount, sc_Fjdcw, gh_Fjdcw,0,1'非机动车位
		
		'填充单元信息
		startRow=iniRow
		for i= 0 to cellCount-1
			cellValue=cellList(i)
			cellValueList=split(cellValue,"||")
			if  ubound(cellValueList)=2 then  startCol=iniCol  else  startCol=iniCol+1
			for j= 0 to ubound(cellValueList)
				g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
		next
End Function

'总体指标表
Function OutputTable01(byval tableIndex,strSubYWMC)
		g_docObj.MoveToTable tableIndex,false
		cellCount=0:redim cellList(cellCount)
		'//规划竣工
		'**************************************************************总用地面积
		ydhxTableName="JG_建设工程规划许可证信息属性表"
		fields="GuiHSPZYDMJ"
		listCount=GetProjectTableList (ydhxTableName,"GuiHSPZYDMJ","","","",list,fieldCount)
		if listCount=1 then gh_YongDMJ=list(0,0)
		gh_YongDMJ=GetFormatNumber(gh_YongDMJ,2)'规划-总用地面积
		listCount=GetProjectTableList ("JGSCHZXX","YDMJ","ID>0","","",list,fieldCount)
		if listCount=1 then sc_YongDMJ = list(0,0)
		if sc_YongDMJ<>"" then sc_YongDMJ = GetFormatNumber(sc_YongDMJ,2)
		g_docObj.SetCellText tableIndex,3,1,gh_YongDMJ,true,false:g_docObj.SetCellText tableIndex,3,2,sc_YongDMJ,true,false
		'**************************************************************总建筑面积
		zrzCount=GetProjectTableList ("JGSCHZXX","JZMJ","ID>0","","",zrzList,fieldCount)
		if zrzCount=1 then sc_SCJZMJ=zrzList(0,0)
		sc_SCJZMJ=GetFormatNumber(sc_SCJZMJ,2)'实测-总建筑面积

		ghxkTableName="JG_建设工程规划许可证信息属性表"
		'exCondition="YDHXGUID In (select YDHXGUID from "&ydhxTableName&"  INNER JOIN GeoLineTB ON "&ydhxTableName&".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0)"
		exCondition="ID>0"
		ghxkCount=GetProjectTableList (ghxkTableName,"sum(GuiHSPZJZMJ)",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then gh_SCJZMJ=ghxkList(0,0)
		gh_SCJZMJ=GetFormatNumber(gh_SCJZMJ,2)'规划-总建筑面积
		g_docObj.SetCellText tableIndex,4,1,gh_SCJZMJ,true,false:g_docObj.SetCellText tableIndex,4,2,sc_SCJZMJ,true,false

		'**************************************************************建筑基底面积
		jdCount=GetProjectTableList ("JGSCHZXX","JZJDMJ","ID>0","","",jdList,fieldCount)
		if jdCount=1 then sc_JDMJ=jdList(0,0)
		sc_JDMJ=GetFormatNumber(sc_JDMJ,2)'实测-建筑基底面积
		ghxkCount=GetProjectTableList (ghxkTableName,"sum(GuiHSPJDMJ),sum(GuiHSPRJL),sum(GuiHSPJZMD),sum(GuiHSPLHL),sum(ZpsJZMJ),sum(ScZZHS),sum(GhZZHS)",exCondition,"","",ghxkList,fieldCount)
		if ghxkCount=1 then 
			gh_JDMJ=ghxkList(0,0):gh_GuiHSPRJL=ghxkList(0,1):gh_GuiHSPJZMD=ghxkList(0,2)
			gh_GuiHSPLHL=ghxkList(0,3):gh_ZpsJZMJ=ghxkList(0,4)
			ScZZHS=ghxkList(0,5):GhZZHS=ghxkList(0,6)
		end if 
		gh_JDMJ=GetFormatNumber(gh_JDMJ,2)'规划-建筑基底面积
		g_docObj.SetCellText tableIndex,5,1,gh_JDMJ,true,false:g_docObj.SetCellText tableIndex,5,2,sc_JDMJ,true,false

		listCount=GetProjectTableList ("JGSCHZXX","RJV","ID>0","","",list,fieldCount)'容积率
		if listCount=1 then sc_Rjl = list(0,0)
		g_docObj.SetCellText tableIndex,7,1,gh_GuiHSPRJL,true,false:g_docObj.SetCellText tableIndex,7,2,sc_Rjl,true,false

		listCount=GetProjectTableList ("JGSCHZXX","JZMD","ID>0","","",list,fieldCount)'建筑密度
		if listCount=1 then sc_Jzmd = list(0,0)
		g_docObj.SetCellText tableIndex,8,1,gh_GuiHSPJZMD,true,false:g_docObj.SetCellText tableIndex,8,2,sc_Jzmd,true,false

		listCount=GetProjectTableList ("JGSCHZXX","ZPSJZMJ","ID>0","","",list,fieldCount)'装配式建筑面积
		if listCount=1 then sc_ZpsJZMJ = list(0,0)
		g_docObj.SetCellText tableIndex,9,1,gh_ZpsJZMJ,true,false:g_docObj.SetCellText tableIndex,9,2,sc_ZpsJZMJ,true,false

		cwTableName="CWSCXX"
		cwCount=GetProjectTableList ("JGSCHZXX","DSJDCWGS+DXJDCWGS,DSJDCWGS,DXJDCWGS","ID>0","","",cwList,fieldCount)
		if  cwCount=1 then    
			sc_ds_Jdcw=cwList(0,1):sc_dx_Jdcw=cwList(0,2)
			if sc_ds_Jdcw="" then sc_ds_Jdcw=0:if sc_dx_Jdcw="" then sc_dx_Jdcw=0
			sc_Jdcw=int(sc_ds_Jdcw)+int(sc_dx_Jdcw)
		end if
		if sc_Jdcw="" then sc_Jdcw=0:if sc_ds_Jdcw="" then sc_ds_Jdcw=0:if sc_dx_Jdcw="" then sc_dx_Jdcw=0

		ghcwTableName="CWGHXX"
		cwCount=GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX<>'非机动车位'","","",ghcwList,fieldCount)
		if  cwCount=1 then    gh_Jdcw=ghcwList(0,0):gh_ds_Jdcw=ghcwList(0,1):gh_dx_Jdcw=ghcwList(0,2)
		if gh_Jdcw="" then gh_Jdcw=0:if gh_ds_Jdcw="" then gh_ds_Jdcw=0:if gh_dx_Jdcw="" then gh_dx_Jdcw=0
		g_docObj.SetCellText tableIndex,10,1,gh_Jdcw,true,false:g_docObj.SetCellText tableIndex,10,4,sc_Jdcw,true,false
		g_docObj.SetCellText tableIndex,10,3,gh_ds_Jdcw,true,false:g_docObj.SetCellText tableIndex,10,6,sc_ds_Jdcw,true,false
		g_docObj.SetCellText tableIndex,11,3,gh_dx_Jdcw,true,false:g_docObj.SetCellText tableIndex,11,6,sc_dx_Jdcw,true,false


		cwCount=GetProjectTableList ("JGSCHZXX","DSFJDCWGS,DXFJDCWGS","ID>0","","",cwList,fieldCount)
		if  cwCount=1 then    
			DSFJDCWGS=cwList(0,0):DXFJDCWGS=cwList(0,1)
			if DSFJDCWGS="" then DSFJDCWGS=0:if DXFJDCWGS="" then DXFJDCWGS=0
			sc_Fjdcw=int(DSFJDCWGS)+int(DXFJDCWGS)
		end if
		ghcwCount=GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL)","CWLX='非机动车位'","","",ghcwList,fieldCount)
		if  ghcwCount=1 then    gh_Fjdcw=ghcwList(0,0)
		if sc_Fjdcw="" then sc_Fjdcw=0:if gh_Fjdcw="" then gh_Fjdcw=0
		g_docObj.SetCellText tableIndex,12,1,gh_Fjdcw,true,false:g_docObj.SetCellText tableIndex,12,2,sc_Fjdcw,true,false


		'//绿地
		ldCount=GetProjectTableList ("JGSCHZXX","LDMJ","ID>0","","",sclhmjList,fieldCount)
		if ldCount = 1 then sc_lhmj=sclhmjList(0,0)
		If gh_YongDMJ<>"" and gh_GuiHSPLHL<>"" then gh_lhmj = GetFormatNumber(gh_GuiHSPLHL*gh_YongDMJ*0.01,2) else gh_lhmj=""
		if sc_lhmj<>"" then sc_lhmj=GetFormatNumber(sc_lhmj,2)
		g_docObj.SetCellText tableIndex,6,1,gh_lhmj,true,false:g_docObj.SetCellText tableIndex,6,2,sc_lhmj,true,false
		g_docObj.SetCellText tableIndex,27,1,gh_lhmj,true,false:g_docObj.SetCellText tableIndex,27,2,sc_lhmj,true,false
		'单块绿地
		ldCount=GetProjectTableList ("JGSCHZXX","DKLDMJ","ID>0","","",scdklhmjList,fieldCount)
		if ldCount = 1 then sc_dk_lhmj=scdklhmjList(0,0)
		ldCount=GetProjectTableList ("PROJECTINFO"," value "," key='单块绿地面积' ","","",ghdklhmjList,fieldCount)
		if ldCount = 1 then gh_dk_lhmj=ghdklhmjList(0,0)
		g_docObj.SetCellText tableIndex,28,1,gh_dk_lhmj,true,false:g_docObj.SetCellText tableIndex,28,2,sc_dk_lhmj,true,false
		'集中绿地
		ldCount=GetProjectTableList ("JGSCHZXX","JZLDMJ","ID>0","","",scjzlhmjList,fieldCount)
		if ldCount = 1 then sc_jz_lhmj=scjzlhmjList(0,0)
		ldCount=GetProjectTableList ("PROJECTINFO"," value "," key='集中绿地面积' ","","",ghjzlhmjList,fieldCount)
		if ldCount = 1 then gh_jz_lhmj=ghjzlhmjList(0,0)
		g_docObj.SetCellText tableIndex,29,1,gh_jz_lhmj,true,false:g_docObj.SetCellText tableIndex,29,2,sc_jz_lhmj,true,false
		
		
		'管线
		gxrow=16
		gxCount=GetProjectTableList ("GXSCHZXX"," distinct GXLB "," ID>0 ","","",gxList,fieldCount)
		for i=0 to gxCount-1
			GXLB=gxList(i,0)
			gxCount1=GetProjectTableList ("GXSCHZXX","GXZL,CGCLCD,TCCD,ZCD"," GXLB='"&GXLB&"' ","","",gxList1,fieldCount)		
			for i1=0 to gxCount1-1
				gxcol=1
				for i2=0 to fieldCount-1
					g_docObj.SetCellText tableIndex,gxrow,gxcol,gxList1(i1,i2),true,false
					gxcol=gxcol+1
				next
				gxrow=gxrow+1
			next
		next

		'//人防
		rfCount=GetProjectTableList ("JGSCHZXX","RFZMJ","ID>0","","",scrfjzmjList,fieldCount)
		if rfCount=1 then sc_rfjzmj= scrfjzmjList(0,0)
		if sc_rfjzmj<>"" then sc_rfjzmj=GetFormatNumber(sc_rfjzmj,2)
		g_docObj.SetCellText tableIndex,33,1,sc_rfjzmj,true,false
		
		'//消防
		xfCount=GetProjectTableList ("JGSCHZXX","DGCDGS","ID>0","","",scxfjzmjList,fieldCount)
		if xfCount=1 then sc_rfjzmj= scrfjzmjList(0,0)
		if sc_xfjzmj<>"" then sc_xfjzmj=GetFormatNumber(sc_xfjzmj,2)
		g_docObj.SetCellText tableIndex,36,1,sc_xfjzmj,true,false

		'根据业务删除行
		strSubYWMC = replace(strSubYWMC,"竣工多测合一:","")
		SubYWMCList = split(strSubYWMC,",")
		strDeleteRowTWMC="规划核实测量,地下管线测量,绿地核实测量,人防核实测量,消防核实测量"
		strDeleteRowTWMCList=""
		for i=0 to ubound(SubYWMCList)
			if replace(strDeleteRowTWMC,SubYWMCList(i),"")<>strDeleteRowTWMC then
					if strDeleteRowTWMCList="" then
						strDeleteRowTWMCList = SubYWMCList(i)
					else
						strDeleteRowTWMCList = strDeleteRowTWMCList&","&SubYWMCList(i)
					end if
			end if
		next
		
	strDeleteYwmc = split(strDeleteRowTWMCList,",")
	str="规划核实测量:（一）规划核实指标,地下管线测量:（二）地下管线规划竣工指标,绿地核实测量:（三）绿地核实指标,人防核实测量:（四）人防核实测量,消防核实测量:（五）消防核实测量"
	arstr = split(str,",")
	for i =0 to ubound(strDeleteYwmc)
		strDeleteRowTWMC= replace(strDeleteRowTWMC,strDeleteYwmc(i),"")
		strDeleteTWMC = strDeleteRowTWMC
		'for i1=0 to ubound(arstr)
			'aar = split(arstr(i1),":")
			'if replace(arstr(i1),strDeleteYwmc(i),"")<>arstr(i1) then  g_docObj.ReplaceDocText tableIndex,StartRow,false
		'next
	next


	ghStartRow=1:gxStartRow=14:ldStartRow=25:rfStartRow=31:xfStartRow=34
	ghRow=13:gxRow=11:ldRow=6:rfRow=3:xfRow=3
	deleteRow=0

	strDeleteTWMCList=split(strDeleteTWMC,",")
	for i=0 to ubound(strDeleteTWMCList)
		if strDeleteTWMCList(i)<>"" then
			if strDeleteTWMCList(i)="规划核实测量" then
				StartRow=ghStartRow
				for i1=0 to ghRow-1
					g_docObj.DeleteRow tableIndex,StartRow,false
				next
				deleteRow=deleteRow+ghRow
			elseif strDeleteTWMCList(i)="地下管线测量" then
				StartRow=gxStartRow-deleteRow
				for i1=0 to gxRow-1
					g_docObj.DeleteRow tableIndex,StartRow,false
				next
				deleteRow=deleteRow+gxRow
			elseif strDeleteTWMCList(i)="绿地核实测量" then
				StartRow=ldStartRow-deleteRow
				for i1=0 to ldRow-1
					g_docObj.DeleteRow tableIndex,StartRow,false
				next
				deleteRow=deleteRow+ldRow
			elseif strDeleteTWMCList(i)="人防核实测量" then
				StartRow=rfStartRow-deleteRow
				for i1=0 to rfRow-1
					g_docObj.DeleteRow tableIndex,StartRow,false
				next
				deleteRow=deleteRow+rfRow
			elseif strDeleteTWMCList(i)="消防核实测量" then
				StartRow=xfStartRow-deleteRow
				for i1=0 to rfRow-1
					g_docObj.DeleteRow tableIndex,StartRow,false
				next
				deleteRow=deleteRow+xfRow
			end if
		end if
	next
End Function


function OutputTable02(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		startRow=3
		listCount=GetProjectTableList ("KZDZBCGXXB","DH,Y,X,GC,GXSJ,BZ"," ID>0 ","","",List,fieldCount)
		if listCount>1 then g_docObj.CloneTableRow tableIndex,3,listCount-1,1,false
		for i=0 to listCount-1
			startCol=0
			for i1=0 to fieldCount-1
				g_docObj.SetCellText tableIndex,startRow,startCol,List(i,i1),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
		next
end function


'//输出 建筑物建筑面积汇总表
Function OutputTable14(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		ZRZTableName = "FC_自然幢信息属性表"
		JDTableName = "JG_建筑物基底面属性表"
		GNQTableName = "JG_规划功能区属性表"
		mdbName = SSProcess.GetProjectFileName  
		SSProcess.OpenAccessMdb mdbName 
		'增加列，画表格
		sql = "select distinct YT from SCLDMJHZXX where  ID>0 and YT<>'*' and YT is not NULL"
		GetSQLRecordAll mdbName, sql, arRecordGHYT, RecordGHYTCount
		for i = 0 to RecordGHYTCount-1
			if strGHYT ="" then
				strGHYT = arRecordGHYT(i)
			else
				strGHYT = strGHYT&","&arRecordGHYT(i)
			end if
		next

		'strGHYT = replace(strGHYT,"住宅","")
		'strGHYT = replace(strGHYT,"商业","")
		'strGHYT = replace(strGHYT,"物业经营","")
		'strGHYT = replace(strGHYT,"物业管理","")
		'strGHYT = replace(strGHYT,"社区用房","")
		arGHYT = split(strGHYT,",")
		col = 2
		'strGHYT1 = "住宅,商业,物业经营,物业管理,社区用房"
		for i = 0 to ubound(arGHYT)
			if arGHYT(i)<>"" then
				g_docObj.InsertTableColumn	tableIndex,col+i,false
				'填值
				'g_docObj.SetCellText tableIndex,0,col+i,arGHYT(i),true,false
				'居中填值
				g_docObj.SetCellValueByBuilder  0 ,  0 ,  col+i ,  arGHYT(i) ,   -1, 1
				'strGHYT1 = strGHYT1&","&arGHYT(i)
			end if
		next
		'strGHYT=strGHYT1
		'增加行,填值
		sql = "select distinct LD,ID_ZRZ from "&ZRZTableName&" where ID>0"
		GetSQLRecordAll mdbName, sql, arRecordCheck, RecordCheckCount
		if RecordCheckCount>0 then  
			startRow  = 1
			rowCount = RecordCheckCount-1
			g_docObj.CloneTableRow tableIndex,startRow,rowCount,1,false
			for i = 0 to RecordCheckCount-1
				arZRZTemp = split(arRecordCheck(i),",")
				g_docObj.SetCellText tableIndex,i+1,0,arZRZTemp(0),true,false
				sql1 = "select sum(JDMJ) from JG_建筑物基底面属性表 where ID_ZRZ = '"&arZRZTemp(1)&"'"
				GetSQLRecordAll mdbName, sql1, arRecordJDMJ, RecordJDMJCount
				if RecordJDMJCount>0 then
					if arRecordJDMJ(0)<>"" then arRecordJDMJ(0)=GetFormatNumber(arRecordJDMJ(0),2)
					g_docObj.SetCellText tableIndex,i+1,1,arRecordJDMJ(0),true,false

					sql2 = "select distinct YT from SCLDMJHZXX where LD = '"&arZRZTemp(0)&"'"
					GetSQLRecordAll mdbName, sql2, arRecordGHYT1, RecordGHYTCount
					for i1 = 0 to RecordGHYTCount-1
						arGHYTTemp = split(arRecordGHYT1(i1),",")
						arTemp = split(strGHYT,",")
						for j= 0 to ubound(arTemp)
							if arGHYTTemp(0) = arTemp(j) then 	
								sql3="select JZMJ from SCLDMJHZXX where LD = '"&arZRZTemp(0)&"' and YT='"&arTemp(j)&"'"
								GetSQLRecordAll mdbName, sql3, arRecordGNQJZMJ, RecordGNQJZMJCount
								if RecordGNQJZMJCount>0 then sc_GNQMJ=arRecordGNQJZMJ(0)
								sc_GNQMJ=GetFormatNumber(sc_GNQMJ,2)
								'g_docObj.SetCellText tableIndex,i+1,j+2,sc_GNQMJ,true,false
								'居中填值
								g_docObj.SetCellValueByBuilder  0 ,  i+1,j+2,sc_GNQMJ,   -1, 1
								g_docObj.SetCellTextFontFormat	tableIndex,i+1,j+2,"宋体",10,0,false
							end if
						next
					next
				end if			
			next
		end if
		'合计
		RowCunt = g_docObj.GetTableRowCount( tableIndex,false)-3
		ColCount =g_docObj.GetTableColCount( tableIndex,0,false)-1
		for i =1 to RowCunt-1
			sumValue = 0
			for i1 =1 to ColCount
				value= g_docObj.GetCellText(tableIndex,i,i1,false)
				value = replace(value,"","")
				if	IsNumeric(value) = true then	sumValue = cdbl(sumValue)+cdbl(value)
			next
			if sumValue = 0 then sumValue =""
			if sumValue<>"" then sumValue=GetFormatNumber(sumValue,2)
			g_docObj.SetCellText tableIndex,i,ColCount,sumValue,true,false
		next
		for i =1 to ColCount
			sumValue = 0
			for i1 =1 to RowCunt-1
				value= g_docObj.GetCellText(tableIndex,i1,i,false)
				value = replace(value,"","")
				if	IsNumeric(value) = true then	sumValue = cdbl(sumValue)+cdbl(value)
			next
			if sumValue = 0 then sumValue =""
			if sumValue<>"" then sumValue=GetFormatNumber(sumValue,2)
			'g_docObj.SetCellText tableIndex,RowCunt,i,sumValue,true,false
			'居中填值
			g_docObj.SetCellValueByBuilder  0 ,  RowCunt,i,sumValue,   -1, 1
			g_docObj.SetCellTextFontFormat	tableIndex,RowCunt,i,"宋体",10,0,false
		next

		'合并
		g_docObj.DeleteCol tableIndex,  ColCount-1,false
		g_docObj.MergeCell tableIndex,  RowCunt+1,  0,  RowCunt+1,  ColCount-1,false
		g_docObj.MergeCell tableIndex,  RowCunt+2,  0,  RowCunt+2,  ColCount-1,false
		
		

		SSProcess.CloseAccessMdb mdbName 
End Function

'//输出 建筑面积变更规划核实表
Function OutputTable15(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		JZMJBGTableName = "JZMJBG"
		mdbName = SSProcess.GetProjectFileName  
		SSProcess.OpenAccessMdb mdbName 
		sql = "select GHXH,GHBGBW,GHJTQK,GHBH,GHSFBA,GHFW,GHMJ,GHBZ from JZMJBG where ID>0 order by GHXH "
		GetSQLRecordAll mdbName, sql, arRecordJZMJBG, RecordJZMJBGCount
		for i=0 to RecordJZMJBGCount-1
			startRow  = 2
			if i>0 then g_docObj.CloneTableRow tableIndex,startRow,1,1,false
		next
		for i =0 to RecordJZMJBGCount-1
			RowCunt = g_docObj.GetTableRowCount( tableIndex,false)
			ColCount =g_docObj.GetTableColCount( tableIndex,0,false)
			arJZMJBGTemp = split(arRecordJZMJBG(i),",")
			for i2 = 0 to ubound(arJZMJBGTemp)
				g_docObj.SetCellText tableIndex,i+2,i2,arJZMJBGTemp(i2),true,false
			next
		next
		SSProcess.CloseAccessMdb mdbName 
End Function

'//输出 特殊部位计算说明规划核实表
Function OutputTable16(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		TSBWJSSMBTableName = "TSBWJSSMB"
		mdbName = SSProcess.GetProjectFileName  
		SSProcess.OpenAccessMdb mdbName 
		sql = "select GHXH,GHJTBW,GHJSGZJFF,GHBZ from TSBWJSSMB where ID>0 order by GHXH "
		GetSQLRecordAll mdbName, sql, arRecordTSBWJSSMB, RecordTSBWJSSMBCount
		for i=0 to RecordTSBWJSSMBCount-1
			startRow  = 1
			if i>0 then g_docObj.CloneTableRow tableIndex,startRow,1,1,false
		next
		for i =0 to RecordTSBWJSSMBCount-1
			RowCunt = g_docObj.GetTableRowCount( tableIndex,false)
			ColCount =g_docObj.GetTableColCount( tableIndex,0,false)
			arTSBWJSSMBTemp = split(arRecordTSBWJSSMB(i),",")
			for i2 = 0 to ubound(arTSBWJSSMBTemp)
				g_docObj.SetCellText tableIndex,i+1,i2,arTSBWJSSMBTemp(i2),true,false
			next
		next
		SSProcess.CloseAccessMdb mdbName 
End Function

Function GetSubArea(cellList,cellCount,byval scArea,byval ghArea,byval numberDigit,byval startCol)
		if isnumeric(numberDigit)=false then numberDigit=2
		if scArea<>"" then scArea1=GetFormatNumber(scArea,numberDigit)	else scArea1=0:scArea="/"
		if ghArea<>"" then ghArea1=GetFormatNumber(ghArea,numberDigit)	else ghArea1=0:ghArea="/"
		subNum=cdbl(scArea1)-cdbl(ghArea1)
		if subNum<>"" then subNum=GetFormatNumber(subNum,numberDigit)'差值-建筑基底面积
		if scArea="/" and ghArea="/" then subNum="/"
		'if scArea = "0.00"  then scArea = "0.00":if scArea = "0"  then scArea = "0"
		'if ghArea = "0.00"  then ghArea = "0.00":if ghArea = "0"  then ghArea = "0"
		'if subNum = "0.00"  then subNum = "0.00":if subNum = "0"  then subNum = "0"

		if startCol=2 then  	cellValue=scArea&"||"&ghArea&"||"&subNum &"||"&""  else 	cellValue=scArea&"||"&ghArea&"||"&subNum
		redim preserve cellList(cellCount): cellList(cellCount)=cellValue:cellCount=cellCount+1
End Function

function SumGnqJzmj(byval lc,byval jzmj,byref sumjzmj)
	if instr(lc,"～")>0 then
		list=split(lc,"～")
		cs=list(1)-list(0)+1
	else
		cs=1
	end if
	sumjzmj=sumjzmj+jzmj*cs
	SumGnqJzmj=sumjzmj
end function

'//获取功能区分类面积
Function GetGnqAreaList(cellList,cellCount,byval KJWZ,byref copyCount)
		copyCount=0
		'**************************************************************建筑面积-各功能区面积
		'ghgnqCount=GetProjectTableList (ghgnqTableName,"SUM(JZMJ),YT",""&strConditon&" group by YT","SpatialData","2",ghgnqList,fieldCount)
		
		ghgnqCount=GetProjectTableList ("ZYJJZBMJHZB","distinct LXMC"," KJWZ='"&KJWZ&"' ","","",ghgnqList,fieldCount)
		sc_GNQMJ=0
		if ghgnqCount>0 then
			'**************************************************************建筑面积
			for i=0 to ghgnqCount-1
				gnqName=ghgnqList(i,0)				
				ghgnqCount1=GetProjectTableList ("ZYJJZBMJHZB","SCJZMJ"," LXMC='"&gnqName&"'","","",ghgnqList1,fieldCount)
				sumjzmj=0
				if ghgnqCount1=1 then sc_gnq_GNQMJ=ghgnqList1(0,0)

				if sc_gnq_GNQMJ<>"" then sc_gnq_GNQMJ=GetFormatNumber(sc_gnq_GNQMJ,2)
				
				ghldxxCount=GetProjectTableList ("ZYJJZBMJHZB","GHJZMJ"," LXMC='"&gnqName&"'","","",ghldxxList,fieldCount)
				IF ghldxxCount=1 THEN ghldxx_gnqmj=ghldxxList(0,0)
				if ghldxx_gnqmj<>"" then ghldxx_gnqmj=GetFormatNumber(ghldxx_gnqmj,2)
				if sc_gnq_GNQMJ="" then sc_gnq_GNQMJ=0
				if ghldxx_gnqmj="" then ghldxx_gnqmj=0
				
				change_gnqmj = GetFormatNumber(sc_gnq_GNQMJ-ghldxx_gnqmj,2)
				'if sc_gnq_GNQMJ = "0.00" or sc_gnq_GNQMJ = "0" then sc_gnq_GNQMJ = ""
				'if ghldxx_gnqmj = "0.00" or ghldxx_gnqmj = "0" then ghldxx_gnqmj = ""
				'if change_gnqmj = "0.00" or change_gnqmj = "0" then change_gnqmj = ""
				cellValue=gnqName&"||"&sc_gnq_GNQMJ&"||"&ghldxx_gnqmj&"||"&change_gnqmj
				redim preserve cellList(cellCount): cellList(cellCount)=cellValue:cellCount=cellCount+1
				copyCount=copyCount+1
			next
		else
				cellValue=gnqName&"||"&""&"||"&""&"||"&""
				redim preserve cellList(cellCount): cellList(cellCount)=cellValue:cellCount=cellCount+1
		end if		
End Function


'//输出 规划测量分幢与规划许可比对结果表
Function OutputTable17(byval tableIndex,byref tablenode)
		cellCount=0 :redim cellList(cellCount)
		
		iniRow=3:iniCol=2
		startRow=iniRow:startCol=iniCol
		g_docObj.MoveToTable tableIndex,false
		
		zrzCount=GetProjectTableList ("JGGHLDXX","ID_ZRZ,LD"," ID>0 order by LD asc","","",zrzList,fieldCount)
		for i=0 to zrzCount-1
			if i>0 then g_docObj.CloneTable  tableIndex, 1,0,false 
		next

		for i= 0 to zrzCount-1			
			ID_ZRZ=zrzList(i,0):ZRZH=zrzList(i,1)
			g_docObj.SetCellText tableIndex,iniRow-1,0,ZRZH,true,false
			'*************************************************************建筑角点坐标 属性获取
			scdCount=GetProjectTableList ("JG_实测点属性表","dh,x,y,sj_x,sj_y","ID_ZRZ='"&ID_ZRZ&"' order by dh asc","SpatialData","0",scdList,fieldCount)
			if scdCount>1 then 	g_docObj.CloneTableRow tableIndex,  iniRow, 2,scdCount-1, false
			for j= 0 to scdCount-1 
				dh=scdList(j,0):x=GetFormatNumber(scdList(j,1),3):y=GetFormatNumber(scdList(j,2),3)
				sj_x=GetFormatNumber(scdList(j,3),3):sj_y=GetFormatNumber(scdList(j,4),3)
				sqr_dist=sqr( ( cdbl(x)-cdbl(sj_x))*( cdbl(x)-cdbl(sj_x)) +( cdbl(y)-cdbl(sj_y))*( cdbl(y)-cdbl(sj_y))  )
				sqr_dist=GetFormatNumber(sqr_dist,3):limit=GetFormatNumber(0.10,2)
				
				OutputTable17_SetCellList cellList,cellCount ,x,  sj_x, 3, tableIndex&"||"&dh,0,sqr_dist,tableIndex
				OutputTable17_SetCellList cellList,cellCount ,y,  sj_y, 3, tableIndex&"||"&dh,0,sqr_dist,tableIndex
			next
			OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			if scdCount=0 then 
				OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
				OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			end if 
			
			'*************************************************************四至间距 属性获取
			if scdCount<2 then  startRow=iniRow+3 else  startRow=iniRow+ 2*scdCount+1
			zbgxCount=GetProjectTableList ("JG_周边关系校核标注属性表","fx,zdtcmc,scjl,ghjl","qdtcmc='"&ZRZH&"' order by fx asc","SpatialData","1",zbgxList,fieldCount)
			zbgxCount1=GetProjectTableList ("JG_周边关系校核标注属性表","fx,qdtcmc,scjl,ghjl","zdtcmc='"&ZRZH&"' order by fx asc","SpatialData","1",zbgxList1,fieldCount)
			if (zbgxCount+zbgxCount1)>1 then 	g_docObj.CloneTableRow tableIndex,  startRow, 1,(zbgxCount+zbgxCount1)-1, false
			for j= 0 to zbgxCount-1
				fx=zbgxList(j,0):zdtcmc=zbgxList(j,1):scjl=GetFormatNumber(zbgxList(j,2),2):ghjl=GetFormatNumber(zbgxList(j,3),2):limit=GetFormatNumber(0.10,2)
				
				OutputTable17_SetCellList cellList,cellCount ,scjl,  ghjl, 2, tableIndex&"||"&fx&"测距"&zdtcmc,0,"",tableIndex
			next			
			for j= 0 to zbgxCount1-1
				fx=zbgxList1(j,0):qdtcmc=zbgxList1(j,1):scjl=GetFormatNumber(zbgxList1(j,2),2):ghjl=GetFormatNumber(zbgxList1(j,3),2):limit=GetFormatNumber(0.10,2)
				if fx="东" then 
					fx="西"
				elseif fx="西" then 
					fx="东"
				elseif fx="北" then
					fx="南"
				elseif fx="南" then 
					fx="北"
				end if
				OutputTable17_SetCellList cellList,cellCount ,scjl,  ghjl, 2, tableIndex&"||"&fx&"测距"&qdtcmc,0,"",tableIndex
			next
			OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			if (zbgxCount+zbgxCount1)=0 then 	OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			
			'*************************************************************建筑高度 属性获取
			if (zbgxCount+zbgxCount1)<2 then  
				startRow=startRow+2 
			else  
				startRow=startRow+2 + (zbgxCount+zbgxCount1) -1
			end if
			jzgdCount=GetProjectTableList ("JG_建筑物屋顶高度信息表","wz,sjgc,scgc","ID_ZRZ='"&ID_ZRZ&"' order by wz asc","","",jzgdList,fieldCount)
			if jzgdCount>1 then 	g_docObj.CloneTableRow tableIndex,  startRow, 1,jzgdCount-1, false
			for j= 0 to jzgdCount-1
				wz=jzgdList(j,0):sjgc=GetFormatNumber(jzgdList(j,1),2):scgc=GetFormatNumber(jzgdList(j,2),2)
				if wz="±0标高" then limit=0.09 else limit=GetFormatNumber(0.10,2)
				OutputTable17_SetCellList cellList,cellCount ,scgc,  sjgc, 2, tableIndex&"||"&wz,0,limit,tableIndex
			next
			OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			if jzgdCount=0 then 	OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			
			'*************************************************************各层层高 属性获取
			if jzgdCount<2 then  startRow=startRow+2  else  startRow=startRow+2 + jzgdCount -1
			cgCount=GetProjectTableList ("JG_立面图标注属性表","bzmc,scgd,sjgd","ID_ZRZ='"&ID_ZRZ&"' order by xh asc","SpatialData","1",cgList,fieldCount)
			if cgCount>1 then 	g_docObj.CloneTableRow tableIndex,  startRow, 1,cgCount-1, false
			for j= 0 to cgCount-1
				bzmc=cgList(j,0):scgd=GetFormatNumber(cgList(j,1),2):sjgd=GetFormatNumber(cgList(j,2),2):
				if scgd>10 then limit=GetFormatNumber(round((0.028+0.0014*scgd),2),2) else limit=0.04
				if bzmc="标准层" then
					zrzCount1=GetProjectTableList ("FC_自然幢信息属性表","dscs","LD='"&ZRZH&"'","SpatialData","2",zrzList1,fieldCount)
					if zrzCount1=1 then dscs=zrzList1(0,0)
					bzmc="3-"&(dscs-1)&"(标准层)"
				end if
				if bzmc="机房层" then bzmc="顶面层"
				OutputTable17_SetCellList cellList,cellCount ,scgd,  sjgd, 2, tableIndex&"||"&bzmc,0,limit,tableIndex
			next
			OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			if cgCount=0 then 	OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			
			'*************************************************************各层外轮廓 属性获取
			if cgCount<2 then  startRow=startRow+2 else  startRow=startRow+2 + cgCount -1			
			wlkCount=GetProjectTableList ("JG_距离标注属性表","wz,sjjl,scjl","ID_ZRZ='"&ID_ZRZ&"'","SpatialData","1",wlkList,fieldCount)
			if wlkCount>1 then 	g_docObj.CloneTableRow tableIndex,  startRow, 1,wlkCount-1, false
			for j= 0 to wlkCount-1
				wz=wlkList(j,0):sjjl=wlkList(j,1):scjl=wlkList(j,2)
				if scjl>10 then limit=GetFormatNumber(round((0.028+0.0014*scjl),2),2) else limit=0.04
				if instr(ZRZH,"#")>0 then
					OutputTable17_SetCellList cellList,cellCount ,scjl,  sjjl, 2, tableIndex&"||标准层||"&wz,3,limit,tableIndex
				else
					zrzCount1=GetProjectTableList ("FC_自然幢信息属性表","dscs","LD='"&ZRZH&"'","SpatialData","2",zrzList1,fieldCount)
					if zrzCount1=1 then dscs=zrzList1(0,0)
					if dscs=1 then 
						strcc="1层" 
					elseif dscs>1 then
						strcc="1-"&dscs&"层"
					else
						strcc="标准层"
					end if
					OutputTable17_SetCellList cellList,cellCount ,scjl,  sjjl, 2, tableIndex&"||"&strcc&"||"&wz,3,limit,tableIndex
				end if
			next
			OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			if wlkCount=0 then 	OutputTable17_SetCellList cellList,cellCount ,"",  "", "", "",1,"",tableIndex
			
			'*************************************************************建筑面积 属性获取
			if wlkCount<2 then  startRow=startRow+2 else  startRow=startRow+2 + wlkCount -1
			ghgnqCount=GetProjectTableList ("CLFZMJHZB","SCDSJZMJ,SCDXJZMJ"," LD='"&ZRZH&"'","","",ghgnqList,fieldCount)
			if ghgnqCount=1 then SCDSJZMJ=ghgnqList(0,0)
			if ghgnqCount=1 then SCDXJZMJ=ghgnqList(0,1)

			if SCDXJZMJ<>"" then SCDXJZMJ=GetFormatNumber(SCDXJZMJ,2)
			if SCDSJZMJ<>"" then SCDSJZMJ=GetFormatNumber(SCDSJZMJ,2)
			limit1=GetFormatNumber(round(0.04*sqr(SCDXJZMJ)+0.002*SCDXJZMJ,2),2)
			limit2=GetFormatNumber(round(0.04*sqr(SCDSJZMJ)+0.002*SCDSJZMJ,2),2)
			OutputTable17_SetCellList cellList,cellCount ,SCDXJZMJ,  "/", 2, tableIndex&"||地下建筑面积",0,limit1,tableIndex
			OutputTable17_SetCellList cellList,cellCount ,SCDSJZMJ,  "/", 2, tableIndex&"||地上建筑面积",0,limit2,tableIndex


			'*************************************************************按数组信息填充单元格属性
			allTableIndex=""
			startRow=iniRow:startCol=iniCol
			for i1= 0 to cellCount-1
				startCol=iniCol
				cellValue=cellList(i1)
				cellValueList=split(cellValue,"||")
				tableIndex=cellValueList(0)
				'每换一张表，初始化表格行
				if instr(allTableIndex,"'"&tableIndex&"'")=0  then 
					startRow=iniRow
					allTableIndex=allTableIndex&","&"'"&tableIndex&"'"
					allNum=""
				end if 
				if ubound(cellValueList)>0 then 
					for j= 0 to ubound(cellValueList)
						if j<>0 then 
							g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
							startCol=startCol+1
						end if 
					next
				end if 
				startRow=startRow+1
			next
			startRow=iniRow:startCol=iniCol
			'合并 建筑角点坐标 点号、差值
			for i1=0 to scdCount-1
				g_docObj.MergeCell tableIndex,  startRow,  2,  startRow+1, 2,false
				g_docObj.MergeCell tableIndex,  startRow,  6,  startRow+1, 6,false
				g_docObj.MergeCell tableIndex,  startRow,  7,  startRow+1, 7,false
				startRow=startRow+2
			next

			if i>0 then
				g_docObj.MoveToTablePreviousNode tableIndex,false
				g_docObj.InsertBreak 5
				tablenode=tablenode+1
			end if
			tableIndex=tableIndex+1


 		next
		
End Function


Function OutputTable17_SetCellList(cellList,cellCount,byval  scNum,byval  ghNum,byval numberDigit,byval exValues,byval index,byval endValues,byval tableIndex)
		scNum=GetFormatNumber(scNum,numberDigit):	ghNum=GetFormatNumber(ghNum,numberDigit)
		sub_Num=GetFormatNumber(cdbl(scNum)-cdbl(ghNum),numberDigit)
		if index=0 then 
			if endValues<>"" then 
				cellValue=exValues&"||"&scNum&"||"&ghNum&"||"&sub_Num&"||"&endValues
			else
				cellValue=exValues&"||"&scNum&"||"&ghNum&"||"&sub_Num
			end if 
		elseif index=2 then 
			cellValue=exValues&"||"&sub_Num&"||"&endValues
		elseif index=3 then 
			cellValue=exValues&"||"&"Δd="&sub_Num&"||"&endValues
		else
			cellValue=tableIndex
		end if  
		redim preserve cellList(cellCount):cellList(cellCount)=cellValue:cellCount=cellCount+1
End Function

'人防基本信息表
function OutputTable18(byval tableIndex)
	g_docObj.MoveToTable tableIndex,false
	'获取人防info表信息
	strTableName="RFPROJECTINFO"
	strField="VALUE"
	strCondition="KEY='建筑结构'"
	listCount=GetProjectTableList (strTableName,strField,strCondition,"","",list,rfFieldCount)
	if listCount=1 then strJZJG=list(0,0)
	strCondition="KEY='地上层数'"
	listCount=GetProjectTableList (strTableName,strField,strCondition,"","",list,rfFieldCount)
	if listCount=1 then strDSCS=list(0,0)
	strRFDYTableName="RFFHDYXX"
	lcCount=GetProjectTableList (strRFDYTableName,"distinct(szcs)"," ID>0 ","","",lcList,fieldCount)
	JZMJCount=GetProjectTableList (strTableName,strField," KEY='人防建筑面积'","","",JZMJList,fieldCount)
	YBMJCount=GetProjectTableList (strTableName,strField," KEY='掩蔽区面积' ","","",YBMJList,fieldCount)
	if lcCount=1 then strDXCS=lcList(0,0)
	if JZMJCount=1 then strJZMJ=JZMJList(0,0)
	if YBMJCount=1 then strYBMJ=YBMJList(0,0)
	'获取projectinfo表信息
	strTableName="PROJECTINFO"
	strField="VALUE"
	strCondition="KEY='测绘单位'"
	listCount=GetProjectTableList (strTableName,strField,strCondition,"","",list,rfFieldCount)
	if listCount=1 then strCHDW=list(0,0) 
	listCount=GetProjectTableList (strTableName,strField," KEY='防护单元个数' ","","",list,rfFieldCount)
	if listCount=1 then FHDYCount=list(0,0) 
	'填单元格
	g_docObj.Replace "{结构类型}",strJZJG,0
	g_docObj.Replace "{人防位于地下层数}",strDXCS,0
	g_docObj.Replace "{防护单元个数}",FHDYCount,0
	g_docObj.Replace "{人防建筑面积}",strJZMJ,0
	g_docObj.Replace "{人防掩蔽面积}",strYBMJ,0
	g_docObj.Replace "{人防地面层数}",strDSCS,0
	g_docObj.Replace "{人防测绘单位}",strCHDW,0


end function


'输出 人防测量成果表
Function OutputTable12(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		'获取 人防测量成果表
		cellCount=0:redim cellList(cellCount)
		strTableName="RFFHDYXX"
		strField="MC,BH,ID_FHDY,JZMJ,YBMJ,GYMJ,ZSGN,PROTECTIONLEVEL,FHDJ,KBDYS,KBSL,SZCS,PSGN,TCWSL,FJDCSL,BZ"
		listCount=GetProjectTableList (strTableName,strField,"MC<>'*' and MC is not null","","",list,rfFieldCount)
		for i=0 to listCount-1
			cellValue=""
			dybh=list(i,0)&list(i,1)
			ID_FHDY=list(i,2)
			'获取注记上的建筑面积、掩蔽面积
			jzmj=0:ybmj=0
			rfNoteCount=GetProjectTableList ("RF_人防防护单元注记属性表","RF_人防防护单元注记属性表.id,ID_FHDY","ID_FHDY='"&ID_FHDY&"'","SpatialData","3",rfNotelist,fieldCount)
			if rfNoteCount=1  then 
				rfid=rfNotelist(0,0)
				fontString=SSProcess.GetObjectAttr (rfid,"SSObj_FontString")
				fontStringList=split(fontString,"\")
				if ubound(fontStringList)=2 then 
					jzmjStr=fontStringList(1):jzmjStrList=split(jzmjStr,"：")
					ybmjStr=fontStringList(2):ybmjStrList=split(ybmjStr,"：")
					
					if ubound(ybmjStrList)=1 then jzmj=jzmjStrList(1)
					if ubound(ybmjStrList)=1 then ybmj=ybmjStrList(1)
					jzmj=replace(jzmj,"平方米",""):ybmj=replace(ybmj,"平方米","")
				end if 
				
				fontStringList1=split(fontString,"：")
				if ubound(fontStringList1)=2 then 
					jzmjStr=fontStringList1(1):jzmjStrList=split(jzmjStr,"平方米")
					ybmjStr=fontStringList1(2):ybmjStrList=split(ybmjStr,"平方米")

					if ubound(ybmjStrList)=1 then jzmj=jzmjStrList(0)
					if ubound(ybmjStrList)=1 then ybmj=ybmjStrList(0)
					jzmj=replace(jzmj,"平方米",""):ybmj=replace(ybmj,"平方米","")
				end if 
			end if 
			jzmj=GetFormatNumber(jzmj,2):		ybmj=GetFormatNumber(ybmj,2)
			
			rfkbCount=0
			strTableName="RF_人防功能区属性表"
			rfgnqCount=GetProjectTableList (strTableName,strTableName&".id","ID_FHDY='"&ID_FHDY&"'","SpatialData","2",rfgnqList,fieldCount)
			for j= 0 to rfgnqCount-1
				rfgnqID=rfgnqList(j,0)
				rfCode=SSProcess.GetObjectAttr (rfgnqID,"SSObj_Code")
				if rfCode="9450053" then rfkbCount=rfkbCount+1  
			next
			
			exValues=dybh&"||"&jzmj&"||"&ybmj
			for j= 5 to rfFieldCount-1
				value=list(i,j)
				if j=10 then value=rfkbCount
				if j=5 then  cellValue=value else cellValue=cellValue&"||"&value
			next
			cellValue=exValues&"||"&cellValue
			redim preserve cellList(cellCount):cellList(cellCount)=cellValue:cellCount=cellCount+1
		next
		if cellCount=0 then Exit Function 
		colList=split(cellList(0),"||")
		
		cellValue=""
		for m= 0 to ubound(colList)
			sumValue=0
			if m=1  or  m=2  or  m=3 or m=7 or m=8 or m= 11 or m=12 then 
				for j=0 to cellCount-1
					cellStrList=split(cellList(j),"||")
					if IsNumeric(cellStrList(m))=true then sumValue=sumValue+cdbl(cellStrList(m))
				next
				if   m=1  or  m=2 then  sumValue=GetFormatNumber(sumValue,2)
			else
				sumValue=""
			end if 
			if m=0 then cellValue="合计" else cellValue=cellValue&"||"&sumValue
		next
		redim preserve cellList(cellCount):cellList(cellCount)=cellValue:cellCount=cellCount+1
		
		'填充 人防测量成果表 单元格
		iniRow=5:iniCol=1
		startRow=iniRow:startCol=iniCol
		if cellCount>1 then   g_docObj.CloneTableRow tableIndex, iniRow, 1,cellCount-1, false
		for i= 0 to cellCount-1
			startCol=iniCol
			cellValue=cellList(i)
			cellValueList=split(cellValue,"||")
			for j= 0 to ubound(cellValueList)
				g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
				startCol=startCol+1
			next
			startRow=startRow+1
		next
		if cellCount>1  then  g_docObj.MergeCell tableIndex,  iniRow,  0,  iniRow+cellCount-1,  0,false
End Function



'//输出 人防面积测绘计算表
Function OutputTable13(byval tableIndex)
		g_docObj.MoveToTable tableIndex,false
		strRFDYTableName="RFFHDYXX"
		strRFDYField="ID_FHDY,MC,BH"
		rfdyCount=GetProjectTableList (strRFDYTableName,strRFDYField,"MC<>'*' and MC is not null group by ID_FHDY,MC,BH order by mc,bh","","",rfdyList,rfdyFieldCount)
		for ii= 1 to rfdyCount-1
			g_docObj.CloneTable  tableIndex, 2,0,false
		next
		for ii= 0 to rfdyCount-1
			cellCount1=0:redim cellList1(cellCount1):erase cellList1
			sumJzmj=0
			ID_FHDY=rfdyList(ii,0):rfdyMC=rfdyList(ii,1):rfdyBH=rfdyList(ii,2)
			strLcTableName="FC_楼层信息属性表"
			strTableName="RF_人防功能区属性表"
			strField=strTableName&".id,id_lc,bh,jzmj,mc"
			strCondition=" ID_FHDY='"&ID_FHDY&"' order by bh asc"
			listCount=GetProjectTableList (strTableName,strField,strCondition,"SpatialData","2",list,fieldCount)
			for i= 0 to listCount-1
				objid=list(i,0):id_lc=list(i,1):bh=list(i,2):jzmj=list(i,3):mc=list(i,4)
				jzmj=GetFormatNumber(jzmj,2)
				sumJzmj=sumJzmj+cdbl(jzmj)
				code=SSProcess.GetObjectAttr(objid,"SSObj_Code") 
				if code="9450033" then 
					beiz="掩蔽"
				elseif code="9450043" then 
					beiz="辅助房间"
				else
					beiz="口部"
				end if 
				'获取人防所关联的楼层信息
				'lcCount=GetProjectTableList (strLcTableName,"szcc,lc","id_lc='"&id_lc&"'","SpatialData","2",lcList,fieldCount)
				lcCount=GetProjectTableList (strRFDYTableName,"szcs,szcs","ID_FHDY='"&ID_FHDY&"'","","",lcList,fieldCount)
				szcc="":lc=""
				if lcCount=1 then 
					szcc=lcList(0,0):lc=lcList(0,1)
				end if 
				GetLCXX	szcc,strText
				cellValue=strText&"||"&lc&"-"&bh&"||"&jzmj&"||"&mc&"||"&beiz
				redim preserve cellList1(cellCount1):cellList1(cellCount1)=cellValue:cellCount1=cellCount1+1
			next

			'数组按备注指定顺序重排
			cellCount=0:redim cellList(cellCount):erase cellList
			beizList=split("掩蔽,口部,辅助房间",",")
			for i= 0 to ubound(beizList)
				for j= 0 to cellCount1-1
					valueList=split(cellList1(j),"||")
					beiz=valueList(4)
					if beiz=beizList(i) then 
						redim preserve cellList(cellCount):cellList(cellCount)=cellList1(j):cellCount=cellCount+1
					end if 
				next
			next
			
			writeRowCount=23
			copyCount=0
			for i= 0 to cellCount-1
				if i>0 and i mod writeRowCount=0 then copyCount=copyCount+1
			next
			
			'根据数据个数复制表格
			for i=0 to copyCount-1
				if  i>0 then   g_docObj.CloneTable  tableIndex, 1,1,false else  g_docObj.CloneTable  tableIndex, 1,0,false
			next
			
			'填充 人防测量成果表 单元格
			tableCount=0:redim  tableList(tableCount):erase tableList
			redim preserve tableList(tableCount):tableList(tableCount)=tableIndex:tableCount=tableCount+1
			iniRow=2:iniCol=0
			startRow=iniRow:startCol=iniCol
			for i= 0 to cellCount-1
				if i>0 and i mod writeRowCount=0 then 
					startRow=iniRow:tableIndex=tableIndex+1
					redim preserve tableList(tableCount):tableList(tableCount)=tableIndex:tableCount=tableCount+1
				end if 
				startCol=iniCol
				cellValue=cellList(i)

				cellValueList=split(cellValue,"||")
				g_docObj.SetCellText tableIndex,0,1,rfdyMC&rfdyBH,true,false
				for j= 0 to ubound(cellValueList)
					g_docObj.SetCellText tableIndex,startRow,startCol,cellValueList(j),true,false
					startCol=startCol+1
				next
				startRow=startRow+1
			next
			MergeSameRowValue tableList, tableCount,0,iniRow,writeRowCount
			MergeSameRowValue tableList, tableCount,4,iniRow,writeRowCount
			sumJzmj=GetFormatNumber(sumJzmj,2)
			g_docObj.SetCellText tableIndex,25,1,sumJzmj,true,false
			tableIndex=tableIndex+1
		next
End Function


'//合并属性一直行
Function MergeSameRowValue(byval tableList,byval tableCount,byval colIndex,byval iniRow,byval writeRowCount)
		mergeRowCount=0:redim mergeRowList(mergeRowCount):erase mergeRowList
		for i= 0 to tableCount-1
			tableIndex= tableList(i)
			allRowValue=""
			for j= iniRow to iniRow+writeRowCount -1
				rowValue=g_docObj.GetCellText( tableIndex, j, colIndex,false)
				rowValue=replace(rowValue,"","")
				if instr(allRowValue,"'"&rowValue&"'")=0  and  rowValue<>""  then
					redim preserve mergeRowList(mergeRowCount):mergeRowList(mergeRowCount)=tableIndex&"||"&rowValue&"||"&j:mergeRowCount=mergeRowCount+1
					allRowValue=allRowValue&","&"'"&rowValue&"'"
				end if 
			next
		next
		
		for i= 0 to mergeRowCount-1
			value=mergeRowList(i):valueList=split(value,"||")
			tableIndex=valueList(0):rowValue=valueList(1):valueStartRow=valueList(2)
			addCount=0
			for j= iniRow to iniRow+writeRowCount -1
				rowValue_=g_docObj.GetCellText( tableIndex, j, colIndex,false)
				rowValue_=replace(rowValue_,"","")
				if  rowValue_=rowValue then   addCount=addCount+1
			next
			if addCount>1 then g_docObj.MergeCell tableIndex,  valueStartRow,  colIndex,  valueStartRow+addCount-1, colIndex,false
		next
End Function



'//数字进位
Function GetFormatNumber(byval number,byval numberDigit)
		if isnumeric(numberDigit)=false then numberDigit=2
		if isnumeric(number)=false then number=0 
		number= formatnumber(round(number+0.00000001,numberDigit),numberDigit,-1,0,0)
		GetFormatNumber=(number)
End Function


'//判断文件是否存在
Function FileExists(byval strSrcFilePath)
		res=false 
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(strSrcFilePath))=true  Then    res=true 
		set fso=nothing
		FileExists=res
End Function


'//获取文件名
Function GetFileName(byval strSrcFilePath)
		GetFileName=""
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(strSrcFilePath))=true  Then     
			set f=fso.getfile(strSrcFilePath)
			GetFileName= fso.GetFileName(f) '获取不含路径的文件名称,这就是输出
		end if
		set f=nothing
		set fso=nothing
End Function


'//获取文件后缀名
Function GetFileExtensionName(byval strSrcFilePath)
		GetFileExtensionName=""
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(strSrcFilePath))=true  Then     
			set f=fso.getfile(strSrcFilePath)
			GetFileExtensionName= fso.GetExtensionName(f)
		end if
		set f=nothing
		set fso=nothing
End Function




'//根据图廓编码打印
Function PrintImage(byval tkCode,byval imageName,byref ShapeHeight,byref ShapeWidth,byref daYZZ)	
		'DeleteAllImage
		outputTitle="成果图打印输出"
		projectFileName=SSProcess.GetProjectFileName
		filePath=SSProcess.GetSysPathName (4)
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		'SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE,AREA" 
		SSProcess.SetSelectCondition "SSObj_Code", "==", tkCode
		if instr(msgInfo ,"地下室")>0 then   		SSProcess.SetSelectCondition "[JianZWMC]", "like", "地下室"    else   SSProcess.SetSelectCondition "[JianZWMC]", "not like", "地下室"
		SSProcess.SelectFilter
		count=SSProcess.GetSelGeoCount
		for i= 0 to count-1
			objID=SSProcess.GetSelGeoValue(i,"SSObj_ID")
			projectName=SSProcess.GetSelGeoValue(i,"[JianZWMC]")
			if projectName="" then 	projectName=SSProcess.GetSelGeoValue(i,"[XiangMMC]")
			scale=SSProcess.GetSelGeoValue(i,"[DaYBL]")
			leftDist=SSProcess.GetSelGeoValue(i,"[ZuoBJ]")
			upDist=SSProcess.GetSelGeoValue(i,"[ShangBJ]")
			daYZZ=SSProcess.GetSelGeoValue(i,"[DaYZZ]")
			if isnumeric(scale)=false then scale=500
			if isnumeric(leftDist)=false then leftDist=0
			if isnumeric(upDist)=false then upDist=0
			if leftDist=0 then leftDist=10:			if upDist=0 then upDist=10
			height=SSProcess.GetSelGeoValue(i,"[ZhiK]")
			width=SSProcess.GetSelGeoValue(i,"[ZhiG]")
			H=0: W=0
			'if isnumeric(width)=false or isnumeric(height)=false then 
				if instr(daYZZ,"A4纵向")>0 then
					BaseHeith=70
					BaseWidth=70
					width=210  	:height=297 
					H=24.9: W=18.8
				elseif instr(daYZZ,"A4横向")>0  then
					BaseHeith=105
					BaseWidth=148.5
					width=297 	:height=210
					H=17.1: W=25.6
					ShapeWidth = 26.345*W
					ShapeHeight = 26.345*H
				elseif instr(daYZZ,"A3纵向") >0 then
					BaseHeith=210
					BaseWidth=148.5
					width=297 	:height=420
					H=37.2: W=26.3
				elseif instr(daYZZ,"A3横向")>0  then
					BaseHeith=148.5
					BaseWidth=210
					width=420 	:height=297
					H=24.9: W=35.2
				elseif instr(daYZZ,"A2纵向")>0  then
					width=420 	:height=594
				elseif instr(daYZZ,"A2横向") >0 then
					width=594 	:height=420
				elseif instr(daYZZ,"A1纵向")>0  then
					width=594 	:height=841
				elseif instr(daYZZ,"A1横向") >0 then
					width=841 	:height=594
				else
					width=297 	:height=210
					H=16.2: W=22.9
				end if 
			'end if
			if H=0 then H=24.9:if W=0 then W=17.6
			ShapeHeight=28.345 *H  : ShapeWidth=28.345 *W
			xDist=1:yDist=0.4
			SSProcess.GetObjectPoint objID,0,x0,y0,z0,ptype0,name0
			SSProcess.GetObjectPoint objID,1,x1,y1,z1,ptype1,name1
			SSProcess.GetObjectPoint objID,2,x2,y2,z2,ptype2,name2

			minX = x0 - 2*Sqr((x0-x1)^2+(y0-y1)^2)/BaseWidth
			minY = y0 - 4*Sqr((x2-x1)^2+(y2-y1)^2)/BaseHeith
			maxX = x2 + 2*Sqr((x0-x1)^2+(y0-y1)^2)/BaseWidth
			maxY = y2 + 4*Sqr((x2-x1)^2+(y2-y1)^2)/BaseHeith
			dpi = 300

			
			if count=1 then 
				imagePath=filePath&projectName&imageName&".bmp"
				SSProcess.WriteEpsIni outputTitle, imageName ,imagePath
			else
				imagePath=filePath&projectName&imageName&i+1&".bmp"
				SSProcess.WriteEpsIni outputTitle, imageName&i+1 ,imagePath
			end if 
			SSFunc.DrawToImage minX-10, minY-5, maxX+10, maxY+10, width&"X"& height, 400, imagePath 

		next
End Function 


'//打印前先删除旧数据
Function DeleteAllImage
		Set fso = CreateObject("Scripting.FileSystemObject")
		filePath=SSProcess.GetSysPathName (4)
		dim filenames(10000)
		GetAllFiles filePath,"bmp",filecount,filenames
		for i= 0 to filecount-1
			projectName=filenames(i)
			if fso.fileExists(projectName)=true then  fso.DeleteFile projectName
		next
		set fso=nothing
End Function 


'//获取所有文件
Function GetAllFiles(ByRef pathname, ByRef fileExt, ByRef filecount, ByRef filenames())
    Dim fso, folder, file, files, subfolder,folder0, fcount
    Set fso = CreateObject("Scripting.FileSystemObject")
	 if  fso.FolderExists(pathname) then 
		 Set folder = fso.GetFolder(pathname)
		 Set files = folder.Files
		 '查找文件
		 For Each file in files
				 extname = fso.GetExtensionName(file.name)
				  If UCase(extname) = UCase(fileExt) Then
					 filenames(filecount) = pathname & file.name
					 filecount = filecount+1
				 End If
		 Next
		 '查找子目录
		 Set subfolder = folder.SubFolders
		 For Each folder0 in subfolder
			  GetAllFiles pathname & folder0.name & "\", fileExt, filecount, filenames
		 Next
	 end if
End Function



'***********************************************************数据库操作函数***********************************************************
'//strTableName 表
'//strFields 字段
'//strAddCondition 条件 
'//strTableType "AttributeData（纯属性表） ,SpatialData（地物属性表）" 
'//strGeoType 地物类型 点、线、面、注记(0点，1线，2面，3注记)
'//rs 表记录二维数组rs(行,列)
'//fieldCount 字段个数
'//返回值 ：sql查询表记录个数
Function GetProjectTableList(byval strTableName,byval strFields,byval strAddCondition,byval strTableType,byval strGeoType,byref rs(),byref fieldCount)
		GetProjectTableList=0:   values="":rsCount = 0:fieldCount=0
		if strTableName="" or strFields="" then Exit function
		if IsTableExits("",strTableName)=false then Exit Function
		'strFields=GetTableAllFields ("", strTableName, strFields)
		if  strFields="" then Exit function
		'设置地物类型
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
			if strAddCondition<>"" then 	 strCondition=" ("&GeoType&".Mark Mod 2)<>0 and "&strAddCondition&""	
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
		'获取当前工程edb表记录
		AccessName=SSProcess.GetProjectFileName
		'判断表是否存在
		'if  IsTableExits(AccessName,strTableName)=false then exit function 
		'set adoConnection=createobject("adodb.connection")
		'strcon="DBQ="& AccessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
		'adoConnection.Open strcon
		Set adoRs=CreateObject("ADODB.recordset")
		count=0
		adoRs.cursorLocation =3 
		adoRs.cursorType =3
		'msgbox sql
		adoRs.open sql  ,adoConnection,3,3
		rcdCount = adoRs.RecordCount
		fieldCount=adoRs.Fields.Count
		redim rs(rcdCount,fieldCount)
		'erase rs
		while adoRs.Eof=false
				nowValues=""
				For i=0 To fieldCount-1
						value=adoRs(i)
						if isnull(value) then value=""
						value=replace(value,",","，")
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

'//判断表是否存在
Function IsTableExits(byval  strMdbName,byval  strTableName_s)
		strMdbName=SSProcess.GetProjectFileName
		IsTableExits=false 
		strTableName_s=ucase(strTableName_s)
		'判断文件DB文件是否存在
		Set fso=CreateObject("Scripting.FileSystemObject")   
		if fso.fileExists(strMdbName)=false then exit function 
		'获取DB文件后缀名
		set f=fso.getfile(strMdbName)
		dbType= fso.GetExtensionName(f)
		set f = nothing 
		set fso = nothing 
		'分DB类型查找
		if dbType="dbf" then 
			strMdbName=Replace(strMdbName,"/","\")
			ipos=InStrRev(strMdbName,"\")
			strMdbName=Left(strMdbName,ipos)
			strcon="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMdbName + ";Extended Properties=dBASE IV;User ID=;Password="
		else
			strcon="DBQ="& strMdbName &";DRIVER={Microsoft Access Driver (*.mdb)};"
		end if 
		'开库
		'''set adoConnection=createobject("adodb.connection")
		'''adoConnection.Open strcon
		Set rsSchema = adoConnection.OpenSchema(20)
		'获取当前DB文件所有表
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
		if IsTableExits=false then addloginfo "【"&strTableName_s&"】表在edb中不存在"
		''adoConnection.Close
		''Set adoConnection = Nothing
End Function	


'///获取当前edb工程表所有字段
'// strTableName 表名
Function GetTableAllFields(byval strAccessName,byval strTableName,byval strInFields)
		strAccessName=SSProcess.GetProjectFileName
		str=""	:strExitFiels="":strUnExitFiels=""
		strInFieldsList=split (strInFields,",")
		'strAccessName=SSProcess.GetProjectFileName  
		SSProcess.OpenAccessMdb strAccessName
      SSProcess.GetAccessFieldInfo strAccessName, strTableName, fieldInfos 
		fieldInfosList=split(fieldInfos,";")
		for j= 0 to ubound(strInFieldsList) 
			fieldExitMark=false
			strInField=ucase(strInFieldsList(j))
			for i = 0 to ubound(fieldInfosList) 
				strs=fieldInfosList(i)
				strsList=split(strs,",")
				strFields=""&UCase(strsList(0))&""
				str=str&","&strFields
				str1=str1&","&strsList(1)
				if strFields=strInField then :fieldExitMark=true 
			next
			if instr(strInField,".ID")=0  and  instr(strInField,"SUM")=0 then 
				if  fieldExitMark=true then   if strExitFiels="" then strExitFiels=strInField else strExitFiels=strExitFiels&","&strInField
				if  fieldExitMark=false then   if strUnExitFiels="" then strUnExitFiels=strInField else strUnExitFiels=strUnExitFiels&","&strInField
			end if 
		next
		GetTableAllFields=strExitFiels
		SSProcess.CloseAccessMdb strAccessName 
		if strUnExitFiels<>"" then    addloginfo "【"&strTableName&"】中的字段【"&strUnExitFiels&"】不存在"
End Function



'//开库
dim  adoConnection
Function InitDB() 
		accessName=SSProcess.GetProjectFileName
		set adoConnection=createobject("adodb.connection")
		strcon="DBQ="& accessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
		adoConnection.Open strcon
End Function

'//关库
Function ReleaseDB()
		adoConnection.Close
		Set adoConnection = Nothing
End Function

'//封面 
Function ReplaceValueFM
		values="项目编号,规划许可证编号,项目名称,项目地址,委托单位,测绘单位,测绘资质证书编号,测量开始时间,测量完成时间,作业依据,已有资料情况,编制人员,检查人员,审核人员,建设单位,不动产权证编号"
		valuesList=split(values,",")
		for i= 0 to ubound(valuesList)
			strFieldValue=""
			strField=valuesList(i)
			listCount=GetProjectTableList ("projectinfo","value","key='"&strField&"'","","",list,fieldCount)
			if listCount=1 then strFieldValue=list(0,0)
			if strField="作业依据" or  strField= "已有资料情况" then
				chrlist=split(strFieldValue,chr(10))
				str=""
				for i1=0 to ubound(chrlist)
					if chrlist(i1)<>"" then
						if str="" then
							str=chrlist(i1)
						else
							str=str&chr(10)&chrlist(i1)
						end if
					end if
				next
				g_docObj.MoveToBookmark strField
				g_docObj.Write(str)
			else
				g_docObj.Replace "{"&strField&"}",strFieldValue,0
			end if
		next
		strFieldValue=""

		strField="仪器名称"
		listCount=GetProjectTableList ("INFO_YQSB",strField,"","","",list,fieldCount)
		for i= 0 to listCount-1
			name=list(i,0)
			if name<>"" then if strFieldValue="" then strFieldValue=name else strFieldValue=strFieldValue&","&name
		next
		g_docObj.Replace "{"&strField&"}",strFieldValue,0
		
		g_docObj.Replace "{年月日}",year(now)&"年"&month(now)&"月"&day(now)&"日",0
End Function

'层次中文转换
function GetLCXX(LC,strText)
		if Len(LC) = 1 then
			for j = 0 to Len(LC)-1
				IntLC = mid(LC,j+1,1)
				NumberChange IntLC,BigNumber
				strText = strText&BigNumber&"层"
			next
		elseif Len(LC)>1 and instr(LC,"-")=0 then 
			for j = 0 to Len(LC)-1
				IntLC = mid(LC,j+1,1)
				NumberChange IntLC,BigNumber
				if strText = "" then
					strText = BigNumber
				else
					strText = strText&"十"&BigNumber&"层"
				end if
			next
			if  mid(LC,1,1) = 1 then
				strText = mid(strText,2,len(strText))
			end if
		elseif Len(LC)>1 and instr(LC,"-")=1 then 
			for j=1 to Len(LC)-1
				IntLC = mid(LC,j+1,1)
				NumberChange IntLC,BigNumber
				strText = "地下"&BigNumber&"层"
			next
		end if

end function 

function NumberChange(Number,BigNumber)
		strNumer = "1,2,3,4,5,6,7,8,9,0"
		strBigNumber = "一,二,三,四,五,六,七,八,九,十"
		artempNumber = split(strNumer,",")
		artempBigNumber = split(strBigNumber,",")
		for i = 0 to 9
			if  artempNumber(i) = Number  then
				BigNumber = artempBigNumber(i)
			end if
		next
end function

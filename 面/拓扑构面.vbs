Sub OnClick()
'添加代码
	geoMaxID = SSProcess.GetGeoMaxID  
	SSProcess.ClearFunctionParameter 
	'悬挂点处理限距
	SSProcess.AddFunctionParameter "limitdist=0.001"
	'拓扑弧段编码
	SSProcess.AddFunctionParameter "SrcArcCodes=9420023,9420021,9420024"
	'删除源弧段
	SSProcess.AddFunctionParameter "DelSrcArc=0"
	'删除上次生成的重叠弧段
	SSProcess.AddFunctionParameter "DelNewArc=0"
	'删除上次生成的原拓扑面
	SSProcess.AddFunctionParameter "DelOldTopArea=0"
	'数据处理后是否存盘
	SSProcess.AddFunctionParameter "SaveDB=1"
	'是否生成拓扑面
	SSProcess.AddFunctionParameter "CreateTopArea=1"
	'拓扑面编码设置  属性点编码1,面编码1,图层名称1/属性点编码2,面编码2,图层名称2
	SSProcess.AddFunctionParameter "NewObject=-1,9420023,竣工测量面积块信息"
	'判断属性点重复的关键字
	SSProcess.AddFunctionParameter "LabelKeyFields="
	'生成拓扑弧段选择：
	'0 不生成弧段
	'1 生成统一编码弧段，编码由UniqueArcCode指定
	'2 生成弧段, 当有多种线状地物重叠时，按 ReserveArcOrder设置的编码顺序优先从前选取
	'3 自动生成与其他弧段重叠的新弧段, 按CreateOverlayArc设置
	SSProcess.AddFunctionParameter "CreateTopArc=0"
	SSProcess.TopProcess "面积块拓扑构面"
	
	SSProcess.ClearSelection  
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition  "SSObj_Type","=","AREA"
	SSProcess.SetSelectCondition  "SSObj_ID",">",geoMaxID
	SSProcess.SetSelectCondition  "SSObj_LayerName","=","竣工测量面积块信息"
	SSProcess.SelectFilter          
	geoCount =  SSProcess.GetSelGeoCount 
	
	SSProcess.ClearSelection  
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition  "SSObj_Type","=","AREA"
	SSProcess.SetSelectCondition  "SSObj_ID",">",geoMaxID
	SSProcess.SetSelectCondition  "SSObj_LayerName","=","竣工测量面积块信息"
	SSProcess.SelectFilter          
	geoCount =   SSProcess.GetSelGeoCount  
	innerObjGetPointMode = 1 '判断焦点
	Dim idList(1000), idCount
	OpenBar "正在匹配面积块属性", geoCount
	for i=0 to geoCount-1
			 If (i mod 10) = 0 Then  RollBar "正在匹配面积块属性", CStr(geoCount-i)
			 geoID = SSProcess.GetSelGeoValue (i, "SSObj_ID") 
			 ids = SSProcess.SearchInnerObjIDs (geoID, 2, "9420023,9420021,9420024", innerObjGetPointMode)
			 If ids<>"" Then
					ScanString ids, ",", idList, idCount
					'同步面标志点位
					posXY = SSRETools.GetAreaLabelPos (idList(0))
					SSRETools.SetAreaLabelPos geoID,posXY 
					'同步属性
					SSProcess.CopyObjectAttr idList(0), geoID, 0, 1

					'删除原面
					SSProcess.DeleteObject idList(0) 
					'获取属性
					MJKMC=SSProcess.GetSelGeoValue (i, "[MianJKMC]") 
					MJKID=SSProcess.GetSelGeoValue (i, "SSObj_ID") 
					SYGN=SSProcess.GetSelGeoValue (i, "[YT]") 
					
					if  SYGN = "住宅"  Then col = RGB(255,0,0)
					if  SYGN = "工业交通仓储"  Then col = RGB(255,255,0)
					if  SYGN = "商业金融信息"  Then col = RGB(0,255,0)
					if  SYGN = "公厕"  Then col = RGB(255,255,0)
					if  SYGN = "教育医疗卫生科研"  Then col = RGB(0,255,255)
					if  SYGN = "文化娱乐体育"  Then col = RGB(0,0,255)
					if  SYGN = "办公"  Then col = RGB(255,0,255)
					if  SYGN = "军事"  Then col = RGB(128,128,128)
					if  SYGN = "未定义面积块"  Then col = RGB(255,255,255)
					if  SYGN = "其他"  Then col = RGB(192,192,192)
					SSProcess.SetObjectAttr MJKID,"SSObj_Color",col
					
			  
			End If
		Next
		CloseBar
End Sub


'创建系统进程条
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

'分解字符串
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

'设置图层是否透明和颜色模式
Function SetLayerMode(byval layernamestr)
    Dim strs1(3000),scount1
	mapHandle  = SSProject.GetActiveMap  
	datasourceHandle = SSProject.GetActiveDatasource (mapHandle )
	lycount = SSProcess.GetLayerCount
	For  i = 0 to lycount-1
		strLayerName =  SSProcess.GetLayerName(i)
		layerHandle = SSProject.GetDataSourceLayerByname (datasourceHandle, strLayerName)
		If strLayerName=layernamestr  Then
			SSProject.SetLayerInfo layerHandle, "DrawAreaMode" , "1"  '0 透明 1 不透明 2 半透明
         SSProject.SetLayerInfo layerHandle, "ColorMode" , "5"  '随符号
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


'这块需要修改提取的属性值
			SCCG = SSProcess.GetObjectAttr (ids, "[ShiCCG]")
			CC = SSProcess.GetObjectAttr (ids, "[项目名称]")
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
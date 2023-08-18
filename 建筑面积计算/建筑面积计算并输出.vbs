
'Excel对象
Dim xlApp,xlFile,xlsheet,yts,sbpts,pcs
Dim YTarr(100000,2),PCarr(100000,2),SBPTarr(100000,2),JZarr(100000,2)
yts = 0
sbpts = 0
pcs = 0
Sub OnClick()
	
	PC "9210413","飘窗"
	PC "9210413","阳台"
	SBPT "9210413","设备平台"
	ExportExcel "9210413"
End Sub

'获取所有独立的室户（包含住宅单元、飘窗、阳台等）
Function PC(ID,lx)
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=",ID
	SSProcess.SelectFilter
	Count=SSProcess.GetSelGeoCount
	'MsgBox Count
	'得到所有的户室部位的名称
	Redim SHBWArray_A(count)
		For i = 0 To count-1
			SHBW = SSProcess.GetSelGeoValue(i, "[SHBW]")
			SHBWArray_A(i) = SHBW
		Next
		s = Join(SHBWArray_A,"@")
		SHBWs = quchongfu(s,True,"@")	
		SHBWArray = split(SHBWs,"@")
		AreaMain = 0.0
		For i=0 to UBound(SHBWArray)
			hsbh = SHBWArray(i)
			hsbh = transform(hsbh)
			If hsbh > 0 Then 
			SSProcess.ClearSelection
			SSProcess.ClearSelectCondition 
			SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
			SSProcess.SetSelectCondition "SSObj_Code", "=",ID
			SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
			SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
			SSProcess.SelectFilter	
			Count=SSProcess.GetSelGeoCount
			'MsgBox Count
			AreaMain = SSProcess.GetSelGeoValue(0, "[KZMJ]") '套内的面积采用勘丈面积
			AreaMain = transform(AreaMain)
			JZArea = SSProcess.GetSelGeoValue(0, "[JZMJ]")
			JZArea = transform(JZArea)
			TotalArea = 0.0
			SubArea = 0.0
			'MsgBox AreaMain
		'DB33／T 1152-2018 《建筑工程建筑面积计算和竣工综合测量技术规程》P20页
		'面积计算规则：1、套内70m²以下，不计面积的飘窗面积大于3m²
		'              2、套内70m²及以上，不计面积的飘窗面积大于5m²
		'超过部分按水平投影的1/2计算
			SSProcess.ClearSelection
			SSProcess.ClearSelectCondition 
			SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
			SSProcess.SetSelectCondition "SSObj_Code", "=",ID
			SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
			SSProcess.SetSelectCondition "[MJKMC]","=",lx
			SSProcess.SelectFilter	
			Count=SSProcess.GetSelGeoCount
			'MsgBox Count
			If Count > 0 Then 
				For l = 0 To Count-1
					SingleArea = SSProcess.GetSelGeoValue(l, "[KZMJ]")
					SingleArea = transform(SingleArea)
					If TotalArea = 0.0 Then 
						TotalArea = SingleArea
					Else 
						TotalArea = TotalArea + SingleArea
					End If 
				Next
			End If 
			If lx = "阳台" Then
				If AreaMain < 70 And TotalArea > 3 Then
					SubArea = (TotalArea - 3) * 0.5
					YTarr(yts,1) = SubArea
					YTarr(yts,0) = SHBWArray(i)
					yts = yts + 1
					JZArea = JZArea + SubArea
					SSProcess.ClearSelection
					SSProcess.ClearSelectCondition 
					SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
					SSProcess.SetSelectCondition "SSObj_Code", "=",ID
					SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
					SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
					SSProcess.SelectFilter
					mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
					SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
					'MsgBox JZArea & "," & SHBWArray(i)
				End If 
			
				If AreaMain > 70 And TotalArea > 5 Then
					SubArea = (TotalArea - 5) * 0.5
					YTarr(yts,1) = SubArea
					YTarr(yts,0) = SHBWArray(i)
					yts = yts + 1
					JZArea = JZArea + SubArea
					SSProcess.ClearSelection
					SSProcess.ClearSelectCondition 
					SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
					SSProcess.SetSelectCondition "SSObj_Code", "=",ID
					SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
					SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
					SSProcess.SelectFilter
					mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
					SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
					'MsgBox JZArea & "," & SHBWArray(i)
				End If 
			
				If AreaMain = 70 And TotalArea > 5 Then
					SubArea = (TotalArea - 5) * 0.5
					YTarr(yts,1) = SubArea
					YTarr(yts,0) = SHBWArray(i)
					yts = yts + 1
					JZArea = JZArea + SubArea
					SSProcess.ClearSelection
					SSProcess.ClearSelectCondition 
					SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
					SSProcess.SetSelectCondition "SSObj_Code", "=",ID
					SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
					SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
					SSProcess.SelectFilter
					mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
					SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
					'MsgBox JZArea & "," & SHBWArray(i)
				End If 
			End If 
		If lx = "飘窗" Then
			If AreaMain < 70 And TotalArea > 3 Then
				SubArea = (TotalArea - 3) * 0.5
				PCarr(pcs,1) = SubArea
				PCarr(pcs,0) = SHBWArray(i)
				pcs = pcs + 1
				JZArea = JZArea + SubArea
				SSProcess.ClearSelection
				SSProcess.ClearSelectCondition 
				SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
				SSProcess.SetSelectCondition "SSObj_Code", "=",ID
				SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
				SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
				SSProcess.SelectFilter
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain > 70 And TotalArea > 5 Then
				SubArea = (TotalArea - 5) * 0.5
				PCarr(pcs,1) = SubArea
				PCarr(pcs,0) = SHBWArray(i)
				pcs = pcs + 1
				JZArea = JZArea + SubArea
				SSProcess.ClearSelection
				SSProcess.ClearSelectCondition 
				SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
				SSProcess.SetSelectCondition "SSObj_Code", "=",ID
				SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
				SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
				SSProcess.SelectFilter
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain = 70 And TotalArea > 5 Then
				SubArea = (TotalArea - 5) * 0.5
				PCarr(pcs,1) = SubArea
				PCarr(pcs,0) = SHBWArray(i)
				pcs = pcs + 1
				JZArea = JZArea + SubArea
				SSProcess.ClearSelection
				SSProcess.ClearSelectCondition 
				SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
				SSProcess.SetSelectCondition "SSObj_Code", "=",ID
				SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
				SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
				SSProcess.SelectFilter
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
		End If 
		End If 
		Next
		msgbox pcs & "," & yts
End Function

Function SBPT(ID,lx)
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=",ID
	SSProcess.SelectFilter
	Count=SSProcess.GetSelGeoCount
	'MsgBox Count
	'得到所有的户室部位的名称
	Redim SHBWArray_A(count)
		For i = 0 To count-1
			SHBW = SSProcess.GetSelGeoValue(i, "[SHBW]")
			SHBWArray_A(i) = SHBW
		Next
		s = Join(SHBWArray_A,"@")
		SHBWs = quchongfu(s,True,"@")	
		SHBWArray = split(SHBWs,"@")
		AreaMain = 0.0
		For i=0 to UBound(SHBWArray)
			sbpthh = SHBWArray(i)
			sbpthh = transform(sbpthh)
			If sbpthh > 0 Then
			SSProcess.ClearSelection
			SSProcess.ClearSelectCondition 
			SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
			SSProcess.SetSelectCondition "SSObj_Code", "=",ID
			SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
			SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
			SSProcess.SelectFilter	
			Count=SSProcess.GetSelGeoCount
			'MsgBox Count
			AreaMain = SSProcess.GetSelGeoValue(0, "[KZMJ]") '套内的面积采用勘丈面积
			AreaMain = transform(AreaMain)
			JZArea = SSProcess.GetSelGeoValue(0, "[JZMJ]")
			JZArea = transform(JZArea)
			TotalArea = 0.0
			SubArea = 0.0
			'MsgBox AreaMain
		'DB33／T 1152-2018 《建筑工程建筑面积计算和竣工综合测量技术规程》P21页
		'面积计算规则：1、套内70m²以下，面积大于3m²
		'              2、套内70m²及以上，面积大于5m²
		'超过部分按水平投影全面积计算
			SSProcess.ClearSelection
			SSProcess.ClearSelectCondition 
			SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
			SSProcess.SetSelectCondition "SSObj_Code", "=",ID
			SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
			SSProcess.SetSelectCondition "[MJKMC]","=",lx
			SSProcess.SelectFilter	
			Count=SSProcess.GetSelGeoCount
			'MsgBox Count
			If Count > 0 Then 
				For l = 0 To Count-1
					SingleArea = SSProcess.GetSelGeoValue(l, "[KZMJ]")
					SingleArea = transform(SingleArea)
					If TotalArea = 0.0 Then 
						TotalArea = SingleArea
					Else 
						TotalArea = TotalArea + SingleArea
					End If 
				Next
			End If 
			
			If AreaMain < 70 And TotalArea > 3 Then
				SubArea = TotalArea - 3
				SBPTarr(sbpts,1) = SubArea
				SBPTarr(sbpts,0) = SHBWArray(i)
				sbpts = sbpts + 1 
				JZArea = JZArea + SubArea
				SSProcess.ClearSelection
				SSProcess.ClearSelectCondition 
				SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
				SSProcess.SetSelectCondition "SSObj_Code", "=",ID
				SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
				SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
				SSProcess.SelectFilter
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain > 70 And TotalArea > 5 Then
				SubArea = TotalArea - 5
				SBPTarr(sbpts,1) = SubArea
				SBPTarr(sbpts,0) = SHBWArray(i)
				sbpts = sbpts + 1 
				JZArea = JZArea + SubArea
				SSProcess.ClearSelection
				SSProcess.ClearSelectCondition 
				SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
				SSProcess.SetSelectCondition "SSObj_Code", "=",ID
				SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
				SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
				SSProcess.SelectFilter
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain = 70 And TotalArea > 5 Then
				SubArea = TotalArea - 5
				SBPTarr(sbpts,1) = SubArea
				SBPTarr(sbpts,0) = SHBWArray(i)
				sbpts = sbpts + 1
				JZArea = JZArea + SubArea
				SSProcess.ClearSelection
				SSProcess.ClearSelectCondition 
				SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
				SSProcess.SetSelectCondition "SSObj_Code", "=",ID
				SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
				SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
				SSProcess.SelectFilter
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
		End If 
		Next
		MsgBox sbpts
End Function

Function ExportExcel(ID)
dim table1:table1 = 2
dim table2:table2 = 2
dim table3:table3 = 2
dim table4:table4 = 2
Dim jzs:jzs = 0
	Filename = SSProcess.GetSysPathName(7)
	ExcelName = "建筑面积检核表.xls"
	ExcelFile = Filename & ExcelName
	Set xlApp=CreateObject("Excel.Application")
	Set xlFile=xlApp.Workbooks.Open(ExcelFile)
	Set xlsheet = xlFile.Worksheets("建筑总面积检核表")
		xlsheet.Activate
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=",ID
	SSProcess.SelectFilter
	Count=SSProcess.GetSelGeoCount
	'MsgBox Count
	'得到所有的户室部位的名称
	Redim SHBWArray_A(count)
		For i = 0 To count-1
			SHBW = SSProcess.GetSelGeoValue(i, "[SHBW]")
			SHBWArray_A(i) = SHBW
		Next
		s = Join(SHBWArray_A,"@")
		SHBWs = quchongfu(s,True,"@")	
		SHBWArray = split(SHBWs,"@")
		For i=0 To UBound(SHBWArray)
			ISHBW = SHBWArray(i)
			ISHBW = transform(ISHBW)
			If ISHBW > 0 Then
			SSProcess.ClearSelection
			SSProcess.ClearSelectCondition 
			SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
			SSProcess.SetSelectCondition "SSObj_Code", "=",ID
			SSProcess.SetSelectCondition "[SHBW]","=",SHBWArray(i)
			SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
			SSProcess.SelectFilter	
			Count=SSProcess.GetSelGeoCount
			
			'msgbox ISHBW
			If Count > 0 Then
				For k = 0 To Count-1
					TNMJ = SSProcess.GetSelGeoValue(k, "[KZMJ]")
					TNMJ = transform(TNMJ)
					ZSJZMJ = SSProcess.GetSelGeoValue(k, "[JZMJ]")
					ZSJZMJ = transform(ZSJZMJ)
					ZSMJ = ZSJZMJ - TNMJ	
					'msgbox TNMJ & "," & ZSJZMJ & "," & ZSMJ
					JZarr(jzs,1) = ZSMJ
					JZarr(jzs,0) = ISHBW
					jzs = jzs + 1 
				Next
			End If
		End If 
		Next
		For n = 0 To jzs - 1
		If  xlApp.Cells(table1,1) = ""  Then 
			xlApp.Cells(table1,1) =  JZarr(n,0)
			xlApp.Cells(table1,2) =  JZarr(n,1)
			table1=table1+1
		End If 
		Next
		
		Set xlsheet = xlFile.Worksheets("建筑阳台面积检核表")
		xlsheet.Activate
		
		For n = 0 To yts - 1
			If  xlApp.Cells(table2,1) = ""  Then
				xlApp.Cells(table2,1) =  YTarr(n,0)
				xlApp.Cells(table2,2) =  YTarr(n,1)
				table2=table2+1
			End If
		Next
		
		Set xlsheet = xlFile.Worksheets("建筑飘窗面积检核表")
		xlsheet.Activate
		
		For n = 0 To pcs - 1
			If  xlApp.Cells(table3,1) = ""  Then
				xlApp.Cells(table3,1) =  PCarr(n,0)
				xlApp.Cells(table3,2) =  PCarr(n,1)
				table3=table3+1
			End If
		Next
		
		Set xlsheet = xlFile.Worksheets("建筑设备平台面积检核表")
		xlsheet.Activate
		
		For n = 0 To sbpts - 1
			If  xlApp.Cells(table4,1) = ""  Then
				xlApp.Cells(table4,1) =  SBPTarr(n,0)
				xlApp.Cells(table4,2) =  SBPTarr(n,1)
				table4=table4+1
			End If
		Next
		xlFile.Save
		xlApp.quit
End Function 

Function quchongfu(srcstr,ig,sp)
 Dim objDict,x,y
 srcarr=split(Trim(srcstr),sp)
 Set objDict=createobject("Scripting.Dictionary")
 For Each x In srcArr
  If x<>"" Then
   If ig=True  Then
    y=LCase(x)
   Else
    y=x
   End If
   If Not objDict.Exists(y) Then objDict.Add x,y
  End If
 Next
 x=Join(objDict.Items,sp)
 If Right(x,1)=sp Then
  quchongfu=left(x,Len(x)-1)
 Else
  quchongfu=x
 End If
 Set objDict=Nothing
End Function

'数据类型转换
Function transform(content)
	If content <> "" Then content = CDbl(content)
	transform = content
End Function
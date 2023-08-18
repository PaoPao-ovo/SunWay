
Sub OnClick()
	PC "9210413","飘窗"
	PC "9210413","阳台"
	SBPT "9210413","设备平台"
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
			
			If AreaMain < 70 And TotalArea > 3 Then
				SubArea = (TotalArea - 3) * 0.5
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
				MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain > 70 And TotalArea > 5 Then
				SubArea = (TotalArea - 5) * 0.5
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
				MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain = 70 And TotalArea > 5 Then
				SubArea = (TotalArea - 5) * 0.5
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
				MsgBox JZArea & "," & SHBWArray(i)
			End If 
		Next
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
		'面积计算规则：1、套内70m²以下，不计面积的飘窗面积大于3m²
		'              2、套内70m²及以上，不计面积的飘窗面积大于5m²
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
				MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain > 70 And TotalArea > 5 Then
				SubArea = TotalArea - 5
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
				MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain = 70 And TotalArea > 5 Then
				SubArea = TotalArea - 5
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
				MsgBox JZArea & "," & SHBWArray(i)
			End If 
		Next
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
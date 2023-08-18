'Excel 所需对象
Dim xlApp,xlFile,xlsheet
'不算面积的阳台数
Dim yts
'不算面积的飘窗数
Dim pcs
'设备平台数
Dim sbpts
'不算面积阳台的户室部位和面积
Dim YTarr(100000,2)
'不算面积飘窗的户室部位和面积
Dim PCarr(100000,2)
'半算阳台飘窗的户室部位和面积
Dim PCYTarr(100000,2)
'设备平台的户室部位和面积
Dim SBPTarr(100000,2)
'户的户室部位和面积
Dim JZarr(100000,2)
yts = 0
sbpts = 0
pcs = 0

'入口函数
Sub OnClick()
    PCYT "9210413"
	NonArea "9210413","飘窗"
	NonArea "9210413","阳台"
	SBPT "9210413","设备平台"
End Sub ' OnClick

'飘窗阳台半算面积的计算规则
Function PCYT(Code) 'Code 编码
        SHBWArray = GetHsArr("9210413")
		'MsgBox SHBWArray(0)
		AreaMain = 0.0
        SubArea = 0.0 
		For i=0 to UBound(SHBWArray)
			hsbh = SHBWArray(i)
			'MsgBox hsbh
			hsbh = transform(hsbh)
            If hsbh > 0 Then
                YTarea = GetYTArea(Code,SHBWArray(i))
                PCarea = GetPCArea(Code,SHBWArray(i))
				'MsgBox PCarea & "---" & SHBWArray(i)
               	SelectNorma Code,SHBWArray(i)
			    Count=SSProcess.GetSelGeoCount
                mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
                AreaMain = SSProcess.GetSelGeoValue(0, "[KZMJ]") '套内的面积采用勘丈面积
			    AreaMain = transform(AreaMain)
                JZArea = SSProcess.GetSelGeoValue(0, "[JZMJ]")
			    JZArea = transform(JZArea)
                'DB33／T 1152-2018 《建筑工程建筑面积计算和竣工综合测量技术规程》P21页
                '按1/2计算后的单套面积住宅阳台和按1/2计算的飘窗面积之和占该套住宅套内建筑面积比值超过 7% 的，超过部分按全面积计算。
                TotalArea = YTarea + PCarea '计算面积的阳台和飘窗的总面积
                SubArea = AreaMain * 0.07
                If TotalArea > SubArea Then
                    OverArea = TotalArea - SubArea '超过部分的面积
                    OverArea = OverArea * 2
                    TotalArea = TotalArea + OverArea
                ElseIf TotalArea < SubArea  Then
                    TotalArea = TotalArea
				Else 
					TotalArea = TotalArea
                End If
            End If
            JZArea = JZArea + TotalArea
            SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
            PCYTarr(i,0) = SHBWArray(i)
            PCYTarr(i,1) = TotalArea
        Next
End Function ' PCYT

'飘窗阳台不算面积计算规则
Function NonArea(Code,mc) 'Code 编码 mc 面积块名称
	    '户室部位数组
		SHBWArray = GetHsArr("9210413")
		AreaMain = 0.0
		For i=0 to UBound(SHBWArray)
			hsbh = SHBWArray(i)
			hsbh = transform(hsbh)
			If hsbh > 0 Then 
			SelectNorma Code,SHBWArray(i) 
			mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")	
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
			TotalArea = GetNonArea(Code,SHBWArray(i),mc)
			If lx = "阳台" Then
				If AreaMain < 70 And TotalArea > 3 Then
					SubArea = (TotalArea - 3) * 0.5
					YTarr(yts,1) = SubArea
					YTarr(yts,0) = SHBWArray(i)
					yts = yts + 1
					JZArea = JZArea + SubArea
					SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
					'MsgBox JZArea & "," & SHBWArray(i)
				End If 
				If AreaMain > 70 And TotalArea > 5 Then
					SubArea = (TotalArea - 5) * 0.5
					YTarr(yts,1) = SubArea
					YTarr(yts,0) = SHBWArray(i)
					yts = yts + 1
					JZArea = JZArea + SubArea
					SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
					'MsgBox JZArea & "," & SHBWArray(i)
				End If 
			
				If AreaMain = 70 And TotalArea > 5 Then
					SubArea = (TotalArea - 5) * 0.5
					YTarr(yts,1) = SubArea
					YTarr(yts,0) = SHBWArray(i)
					yts = yts + 1
					JZArea = JZArea + SubArea
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
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain > 70 And TotalArea > 5 Then
				SubArea = (TotalArea - 5) * 0.5
				PCarr(pcs,1) = SubArea
				PCarr(pcs,0) = SHBWArray(i)
				pcs = pcs + 1
				JZArea = JZArea + SubArea
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
			
			If AreaMain = 70 And TotalArea > 5 Then
				SubArea = (TotalArea - 5) * 0.5
				PCarr(pcs,1) = SubArea
				PCarr(pcs,0) = SHBWArray(i)
				pcs = pcs + 1
				JZArea = JZArea + SubArea
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
		End If 
		End If 
		Next
		'msgbox pcs & "," & yts
End Function

'设备平台计算规则
Function SBPT(Code,mc)'Code 编码 mc 面积块名称
	    SHBWArray = GetHsArr("9210413")
		'MsgBox UBound(SHBWArray)
		AreaMain = 0.0
		For i=0 to UBound(SHBWArray)
			sbpthh = SHBWArray(i)
			sbpthh = transform(sbpthh)
			If sbpthh > 0 Then
			SelectNorma Code,SHBWArray(i)
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
			
			TotalArea = GetSbPtArea(Code,SHBWArray(i),mc)
			If AreaMain < 70 And TotalArea > 3 Then
				SubArea = TotalArea - 3
				SBPTarr(sbpts,1) = SubArea
				SBPTarr(sbpts,0) = SHBWArray(i)
				sbpts = sbpts + 1 
				JZArea = JZArea + SubArea
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
				mainid = SSProcess.GetSelGeoValue(0, "SSObj_ID")
				SSProcess.SetObjectAttr mainid,"[JZMJ]",JZArea
				'MsgBox JZArea & "," & SHBWArray(i)
			End If 
		End If 
		Next
		'MsgBox sbpts
End Function

'选择当前幢
Function SelectNorma(Code,shbw) 'Code 编码 shbw 室户部位
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition 
	SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=",Code
	SSProcess.SetSelectCondition "[SHBW]","=",shbw
	SSProcess.SetSelectCondition "[MJKMC]","=","住宅单元"
	SSProcess.SelectFilter
End Function ' SelectNorma

'获取当前幢设备平台的面积
Function GetSbPtArea(Code,shbw,mc) 'Code 编码 shbw 室户部位 mc 面积块名称
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition 
	SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=",Code
	SSProcess.SetSelectCondition "[SHBW]","=",shbw
	SSProcess.SetSelectCondition "[MJKMC]","=",mc
	SSProcess.SelectFilter	
	Count=SSProcess.GetSelGeoCount
	'MsgBox Count
	If Count > 0 Then 
		For l = 0 To Count-1
			SingleArea = SSProcess.GetSelGeoValue(l, "[JZMJ]")
			SingleArea = transform(SingleArea)
			If TotalArea = 0.0 Then 
				TotalArea = SingleArea
			Else 
				TotalArea = TotalArea + SingleArea
			End If 
		Next
	End If 
End Function ' GetSbPtArea

'获取当前幢不算的面积(飘窗或阳台)
Function GetNonArea(Code,shbw,mc)'Code 编码 shbw 室户部位 mc 面积块名称
	SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=",Code
    SSProcess.SetSelectCondition "[MJKMC]", "=",mc
    SSProcess.SetSelectCondition "[SHBW]","=",shbw
    SSProcess.SelectFilter
    Count=SSProcess.GetSelGeoCount
    JSArea = 0.0
	 If Count > 0 Then   
        For i=0 To Count -1
            mjxs = SSProcess.GetSelGeoValue(i, "[MJXS]")
			mjxs = transform(mjxs)
            If mjxs = 0 Then
                temp =  SSProcess.GetSelGeoValue(i, "[KZMJ]")
                temp = transform(temp)
                If JSArea = 0.0 Then
                  JSArea = temp
                Else JSArea = JSArea + temp
                End If
            End If 
        Next 'i
    End If
	GetNonArea = JSArea
End Function ' GetNonArea

'获取当前幢半算的阳台面积
Function GetYTArea(Code,shbw) ' Code 编码 shbw 室户部位
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=",Code
    SSProcess.SetSelectCondition "[MJKMC]", "=","阳台"
    SSProcess.SetSelectCondition "[SHBW]","=",shbw
    SSProcess.SelectFilter
    Count=SSProcess.GetSelGeoCount
    JSArea = 0.0
    If Count > 0 Then   
        For i=0 To Count -1
            mjxs = SSProcess.GetSelGeoValue(i, "[MJXS]")
			mjxs = transform(mjxs)
            If mjxs = 0.5 Then
                temp =  SSProcess.GetSelGeoValue(i, "[JZMJ]")
                temp = transform(temp)
                If JSArea = 0.0 Then
                  JSArea = temp
                Else JSArea = JSArea + temp
                End If
            End If 
        Next 'i
    End If
	GetYTArea = JSArea
End Function ' GetYTArea

'获取当前幢半算的飘窗面积
Function GetPCArea(Code,shbw) ' Code 编码 shbw 室户部位
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=",Code
    SSProcess.SetSelectCondition "[MJKMC]", "=","飘窗"
    SSProcess.SetSelectCondition "[SHBW]","=",shbw
    SSProcess.SelectFilter
    Count=SSProcess.GetSelGeoCount
    JSArea = 0.0
    If Count > 0 Then   
        For i=0 To Count -1
            mjxs = SSProcess.GetSelGeoValue(i, "[MJXS]")
			mjxs = transform(mjxs)
            If mjxs = 0.5 Then
                temp =  SSProcess.GetSelGeoValue(i, "[JZMJ]")
                temp = transform(temp)
                If JSArea = 0.0 Then
                  JSArea = temp
                Else JSArea = JSArea + temp
                End If
            End If 
        Next 'i
    End If
	GetPCArea = JSArea
End Function ' GetPCArea

'返回所有的户室部位值
Function GetHsArr(Code) 'Code 编码
    SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
	SSProcess.SetSelectCondition "SSObj_Code", "=",Code
	SSProcess.SelectFilter
	Count=SSProcess.GetSelGeoCount
    '得到所有的户室部位的名称
	Redim SHBWArray_A(Count)
		For i = 0 To Count-1
			SHBW = SSProcess.GetSelGeoValue(i, "[SHBW]")
			SHBWArray_A(i) = SHBW
		Next
		s = Join(SHBWArray_A,"@")
		SHBWs = quchongfu(s,True,"@")	
		SHBWArray = split(SHBWs,"@")
        GetHsArr = SHBWArray
End Function ' GetHsArr

'去重复值
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


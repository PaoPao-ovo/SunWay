Dim MJKID,LCID,arrmjkmc(1000),mjkcount,arrmjksygn(1000),arrmjkmjxs(1000),arrmjkmjjrxs(1000),arrsfjr(1000)
Sub OnInitScript()
	mode = 0 '=0 无参数对话框 =1 有参数对话框
	title="选择面积块"
	SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title,"DlgWidth" , "340" 
	SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title, "DlgHeight" ,"430" 
	SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title, "ColumnWidth" ,"160" 
	SSProcess.ShowScriptDlg mode,title 
	SSProcess.SetCursorStatus 6

End Sub

Sub OnOK()
			SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
			SSProcess.ExecuteSDLFunction "$SDL.SSWorkSpace.Display.RedrawExtend", 0
			SSProcess.RefreshView
End Sub

Sub OnExitScript()

End Sub

Sub OnCancel()
		SSProcess.SetCursorStatus 6

End Sub



Function GetAllLC(byref LCIDS(),byref LCCOUNT)
		GetInfo dh,gcbh
      LCCOUNT=0
		mdbName = SSProcess.GetProjectFileName
		SSProcess.OpenAccessMdb mdbName
		sql= "SELECT JG_建筑面积分层图信息属性表.ID FROM JG_建筑面积分层图信息属性表 INNER JOIN GeoLineTB ON JG_建筑面积分层图信息属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark mod 2<>0 )"
		SSProcess.OpenAccessRecordset mdbName, sql 
		While SSProcess.AccessIsEOF(mdbName, sql )=false
				SSProcess.GetAccessRecord mdbName, sql, fields, values
            LCIDS(LCCOUNT)=values
            LCCOUNT=LCCOUNT+1
				SSProcess.AccessMoveNext mdbName, sql 
		wend
		SSProcess.CloseAccessRecordset mdbName, sql
		SSProcess.CloseAccessMdb mdbName
End Function

Function GetInfo(dh,gcbh)
Dim ArrValue(3)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName 
	sql = "SELECT JG_建筑面积分层图信息属性表.JianZWMC,GuiHXKZBH FROM JG_建筑面积分层图信息属性表 INNER JOIN GeoLineTB ON JG_建筑面积分层图信息属性表.ID = GeoLineTB.ID WHERE [GeoLineTB].[Mark] Mod 2 <>0 "
  SSProcess.OpenAccessRecordset projectName, sql
	rscount = SSProcess.GetAccessRecordCount (projectName, sql )
	'If rscount = 1 Then
		SSProcess.AccessMoveFirst projectName, sql
		while (SSProcess.AccessIsEOF (projectName, sql ) = False)
			SSProcess.GetAccessRecord projectName, sql, fields, values
			SSFunc.Scanstring values,",",ArrValue,ArrValueCount
			DH = ArrValue(0)
			gcbh = ArrValue(1)
			SSProcess.AccessMoveNext projectName, sql 
		Wend
	'End If
	SSProcess.CloseAccessRecordset projectName, sql 
	SSProcess.CloseAccessMdb projectName 
End Function 


Function OnLButtonDown(x, y, spx, spy, flags)
      OnLButtonDown=1
      Dim LCIDS(800),LCCOUNT
		GetAllLC LCIDS,LCCOUNT
      LCID=""
      For i=0 To LCCOUNT-1
				POA=SSProcess.IsPtInPoly  (spx, spy, LCIDS(i), 0.001) 
            If POA<>0 and  POA<>2 Then
                  LCID=LCIDS(i)  :  Exit For
            End If
      Next
		If LCID<>"" then
			MJKIDS=SSProcess.SearchInnerObjIDs (LCID, "2", "9400403", 0) 
			Dim strs(2000),scount
			SSFunc.ScanString MJKIDS,",",strs,scount
			MJKID=""
			For i=0 To scount-1
				POA=SSProcess.IsPtInPoly  (spx, spy, strs(i), 0.001) 
				If POA<>0 And POA<>2 Then
						MJKID=strs(i)  :  Exit For
				End If
			Next
			If MJKID<>"" Then '
					arrmjkmc(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[MianJKMC]")'面积块名称
					arrmjksygn(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[GongNYT]") '功能用途
					arrmjkmjxs(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[MianJXS]") '面积系数
					arrmjkmjjrxs(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[JiRMJXS]") '计容面积系数
					arrsfjr(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[ShiFJR]") '是否计计容
					mjkcount=mjkcount+1
					MJKCG=SSProcess.GetObjectAttr(MJKID,"[CengG]") 
		  'GetInfo dh,gcbh

			xkzxx= SSProcess.GetObjectAttr (MJKID, "[GuiHXKZBH]")'规划许可证名称
			lzhxx=SSProcess.GetObjectAttr (MJKID, "[JianZWMC]")'建筑物名称
			getGNYTXX lzhxx,xkzxx,GNYTXX,GNMCXX
			End If
			mode = 1 '=0 无参数对话框 =1 有参数对话框
			title="面积块信息录入"
			SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title,"DlgWidth" , "240" 
			SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title, "DlgHeight" ,"250" 
			SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title, "ColumnWidth" ,"110" 
			'SSProcess.ClearInputParameter
			for c=0 to mjkcount-1
				If arrmjksygn(c) <>"" Then 
					SSProcess.AddInputParameter "使用功能名称", arrmjksygn(c),3,"SYS_DROPDOWNLIST,"&GNYTXX, ""

					'SSProcess.AddInputParameter "面积块名称", arrmjkmc(c),0, "住宅,商业,办公,工业交通仓储,教育医疗卫生科研,文化娱乐体育,军事,建筑物主体,地下室,半地下室,地下室出入口,架空层,坡屋顶,场馆看台下的建筑,无围护的场观看台,门厅,大厅,悬挑看台,架空走廊,无围护的架空走廊,立体书库,立体仓库,立体车库,控制室,落地橱窗,凸(飘)窗,室外走廊（挑廊）,檐廊,门斗,门廊,有柱雨篷,无柱雨篷,楼梯间、水箱间、电梯机房,室内楼梯,电梯井,提物井,通风排气竖井,烟道,有顶盖采光井,室外楼梯,阳台,半阳台,车棚,货棚,站台,加油站,收费站,变形缝,设备层,管道层,避难层,不计容构筑物,未定义面积块,其他", ""
					SSProcess.AddInputParameter "面积块名称", arrmjkmc(c),0, GNMCXX, ""

					SSProcess.AddInputParameter "面积系数" ,arrmjkmjxs(c), 0, "0,1,0.5", ""
					'SSProcess.AddInputParameter "权属性质", "",0, "私有,共有", ""
					'SSProcess.AddInputParameter "序号自动累加", "是",3, "SYS_DROPDOWNLIST,是,否", ""
					SSProcess.AddInputParameter "是否计容", arrsfjr(c),3, "SYS_DROPDOWNLIST,是,否", ""
					SSProcess.AddInputParameter "计容面积系数", arrmjkmjjrxs(c),0, "1,0.5", ""
				Else
					MJKlx1 = "住宅,商业,办公,工业交通仓储,教育医疗卫生科研,文化娱乐体育,军事,建筑物主体,地下室,半地下室,地下室出入口,架空层,坡屋顶,场馆看台下的建筑,无围护的场观看台,门厅,大厅,悬挑看台,架空走廊,无围护的架空走廊,立体书库,立体仓库,立体车库,控制室,落地橱窗,凸(飘)窗,室外走廊（挑廊）,檐廊,门斗,门廊,有柱雨篷,无柱雨篷,楼梯间、水箱间、电梯机房,室内楼梯,电梯井,提物井,通风排气竖井,烟道,有顶盖采光井,室外楼梯,阳台,半阳台,车棚,货棚,站台,加油站,收费站,变形缝,设备层,管道层,避难层,不计容构筑物,未定义面积块,其他"


					MJKGN1=SSProcess.ReadEpsIni("MJKGNXX", "MJKSYGN" ,"")
					MJKGN2=SSProcess.ReadEpsIni("MJKGNXX", "MJKMCZ" ,"")
					MJKGN3=SSProcess.ReadEpsIni("MJKGNXX", "SFJRZ" ,"")
					MJKGN4=SSProcess.ReadEpsIni("MJKGNXX", "MianJXSZ" ,"")
					MJKGN5=SSProcess.ReadEpsIni("MJKGNXX", "JiRMJXSZ" ,"")
'&,,
					SSProcess.AddInputParameter "使用功能名称", MJKGN1,3,"SYS_DROPDOWNLIST,"&GNYTXX, ""
					SSProcess.AddInputParameter "面积块名称", MJKGN2,0,GNMCXX&MJKlx1, ""
					SSProcess.AddInputParameter "面积系数" ,MJKGN4, 0, "0,1,0.5", ""
					SSProcess.AddInputParameter "是否计容",MJKGN3,3, "SYS_DROPDOWNLIST,是,否", ""
					SSProcess.AddInputParameter "计容面积系数", MJKGN5,0, "1,0.5", ""
				End If
			next 
			SSProcess.ShowScriptDlg mode,title
			SSProcess.SetCursorStatus 0
		End If
End Function

dim MJKCG
'属性值发生改变
Function OnPropertyChanged( strName, strValue)
If isnumeric(mjkcg) = false Then mjkcg = 0
		SSProcess.UpdateScriptDlgParameter 1
			if  strName="使用功能名称"  Then
				if strValue <> "公建"  then 
					SSProcess.AddInputParameter "面积块名称", arrmjkmc(c),0, "住宅,商业,办公,工业交通仓储,教育医疗卫生科研,文化娱乐体育,军事,建筑物主体,地下室,半地下室,地下室出入口,架空层,坡屋顶,场馆看台下的建筑,无围护的场观看台,门厅,大厅,悬挑看台,架空走廊,无围护的架空走廊,立体书库,立体仓库,立体车库,控制室,落地橱窗,凸(飘)窗,室外走廊（挑廊）,檐廊,门斗,门廊,有柱雨篷,无柱雨篷,楼梯间、水箱间、电梯机房,室内楼梯,电梯井,提物井,通风排气竖井,烟道,有顶盖采光井,室外楼梯,阳台,半阳台,车棚,货棚,站台,加油站,收费站,变形缝,设备层,管道层,避难层,不计容构筑物,未定义面积块,其他", ""
				SSProcess.AddInputParameter "是否计容", "是",3, "SYS_DROPDOWNLIST,是,否", ""
				else
					xkzxx= SSProcess.GetObjectAttr (MJKID, "[GuiHXKZBH]")'规划许可证名称
					lzhxx=SSProcess.GetObjectAttr (MJKID, "[JianZWMC]")'建筑物名称
					getGNMCXX lzhxx,xkzxx,GNMCZB
					SSProcess.AddInputParameter "面积块名称", arrmjkmc(c),0,GNMCZB, ""
				SSProcess.AddInputParameter "是否计容", "是",3, "SYS_DROPDOWNLIST,是,否", ""
				end if
 
		elseif  strName="是否计容"  Then
				if strValue = "是"  then 
				dyjmjkmc = SSProcess.GetInputParameter ("面积块名称")'第一级面积块名称
				dyjgnyt = SSProcess.GetInputParameter ("使用功能名称")'第一级面积块名称
				If dyjgnyt="住宅"   then
					if dyjmjkmc = "门厅" or dyjmjkmc = "大厅" then 
					jrmjxs = 1
					else
						IF MJKCG <= 3.6 then 
							jrmjxs = 1
						ELSe
							a = fix((mjkcg-3.6)/2.2)
							jrmjxs = a + 2
						End if 
					end if
				Elseif dyjgnyt = "行政办公"  Then
					if dyjmjkmc = "建筑物主体"  then 
							IF MJKCG <= 4.5 then 
								jrmjxs = 1
							ELSe
								a = fix((mjkcg-4.5)/2.2)
								jrmjxs = a + 2
							End if 
					else
						jrmjxs = 1
					end if
				Elseif dyjgnyt = "商业"  Then
					if dyjmjkmc = "建筑物主体"  then 
						IF MJKCG <= 5.1 then 
							jrmjxs = 1
						ELSe
							a = fix((mjkcg-5.1)/2.2)
							jrmjxs = a + 2
						End if 
					else
						jrmjxs = 1
					end if
				Else 
					jrmjxs = 1
				End if
					SSProcess.SetInputParameter "计容面积系数", jrmjxs
				else 
					SSProcess.SetInputParameter "计容面积系数", "0"
				end if
		elseif  strName="面积块名称"  Then
				If strValue="未定义面积块"   then
					SSProcess.SetInputParameter "使用功能名称", strValue
					SSProcess.SetInputParameter "面积系数", "1"
				ElseIf strValue="避难层" or strValue="管道层" or strValue="设备层" or strValue="楼梯间、水箱间、电梯机房"  or strValue="门斗"  or strValue="落地橱窗"  or strValue="控制室"  or strValue="立体车库"  or  strValue="立体仓库"  or  strValue="立体书库"  or strValue="大厅"  or strValue="门厅"  or strValue="架空层"  or strValue="建筑物主体"  or   strValue="地下室" or strValue="半地下室"  Then
					'SSProcess.SetInputParameter "使用功能名称", "住宅"
					IF MJKCG < 2.2 THEN 
					SSProcess.SetInputParameter "面积系数", "0.5"
					ELSE 
					SSProcess.SetInputParameter "面积系数", "1"
					end if 
				ElseIf  strValue="坡屋顶"  or strValue="场馆看台下的建筑"  Then
					'SSProcess.SetInputParameter "使用功能名称", "住宅"
					IF MJKCG < 1.2 THEN 
					SSProcess.SetInputParameter "面积系数", "0"
					ELSEif MJKCG < 2.1 THEN 
					SSProcess.SetInputParameter "面积系数", "0.5"
					ELSE
					SSProcess.SetInputParameter "面积系数", "1"
					END IF
				ElseIf replace(strValue,"夹层","") <>  strValue  Then
					'SSProcess.SetInputParameter "使用功能名称", "住宅"
					IF MJKCG < 2 THEN 
					SSProcess.SetInputParameter "面积系数", "0"
					END IF
				ElseIf   strValue="其他" or strValue="变形缝" or  strValue="保温层" or  strValue="阳台" or strValue="烟道" or strValue="通风排气竖井" or strValue="管道井" or strValue="提物井" or strValue="电梯井" or strValue="室内楼梯" or strValue="架空走廊"  Then
					'SSProcess.SetInputParameter "使用功能", "住宅"
					SSProcess.SetInputParameter "面积系数", "1"
				ElseIf strValue="收费站" or strValue="加油站" or strValue="站台" or strValue="货棚" or strValue="车棚" or strValue="半阳台" or strValue="室外楼梯" or strValue="有柱雨篷" or strValue="门廊" or strValue="檐廊" or strValue="室外走廊（挑廊）" or strValue="地下室出入口" or strValue="无围护的架空走廊"  Then
					'SSProcess.SetInputParameter "使用功能", "住宅"
					SSProcess.SetInputParameter "面积系数", "0.5"
				ElseIf  strValue="不计容构筑物"  Then
					'SSProcess.SetInputParameter "使用功能", "住宅"
					SSProcess.SetInputParameter "面积系数", "0"
				ElseIf   strValue="无柱雨篷" or strValue="凸（飘）窗"  Then
					IF MJKCG >= 2.1 THEN 
					'SSProcess.SetInputParameter "使用功能", "住宅"
					SSProcess.SetInputParameter "面积系数", "0.5"
					else
					SSProcess.SetInputParameter "面积系数", "0"
				End If
				ElseIf   strValue="有顶盖采光井"   Then
					IF MJKCG >= 2.1 THEN 
					'SSProcess.SetInputParameter "使用功能", "住宅"
					SSProcess.SetInputParameter "面积系数", "1"
					else
					SSProcess.SetInputParameter "面积系数", "0.5"
					End If
				ELSE
					SSProcess.SetInputParameter "面积系数", "1"
					'SSProcess.SetInputParameter "计容面积系数", "1"
				End If

				SYGNSXZ = SSProcess.GetInputParameter ("使用功能名称")'第一级使用功能
				If SYGNSXZ="住宅"   then
				if MJKCG="" then MJKCG=1
					if strValue = "门厅" or strValue = "大厅" then 
						jrmjxs = 1
					else
						IF MJKCG <= 3.6 then 
							jrmjxs = 1
						ELSe
							a = fix((mjkcg-3.6)/2.2)
							jrmjxs = a + 2
						End if 
					end if
				Elseif SYGNSXZ = "行政办公"  Then
					if strValue = "建筑物主体"  then 
							IF MJKCG <= 4.5 then 
								jrmjxs = 1
							ELSe
								a = fix((mjkcg-4.5)/2.2)
								jrmjxs = a + 2
							End if 
					else
						jrmjxs = 1
					end if
				Elseif SYGNSXZ = "商业"  Then
					if strValue = "建筑物主体"  then 
						IF MJKCG <= 5.1 then 
							jrmjxs = 1
						ELSe
							a = fix((mjkcg-5.1)/2.2)
							jrmjxs = a + 2
						End if 
					else
						jrmjxs = 1
					end if
				Else 
					jrmjxs = 1
				End if
					SSProcess.SetInputParameter "计容面积系数", jrmjxs
		end IF
		OnPropertyChanged = 0
End Function


Function OnRButtonDown(x, y, spx, spy, flags)
      OnRButtonDown=1
      If MJKID<>"" Then 
            LCGUID=SSProcess.GetObjectAttr(MJKID,"[LCGUID]")
            MYC=SSProcess.GetObjectAttr(MJKID,"[MYC]")
           	JHMJ=SSProcess.GetObjectAttr(MJKID,"SSObj_Area")
				JHMJ = FormatNumber(JHMJ, 2) 
				SSProcess.UpdateScriptDlgParameter 1
				'dyxh=SSProcess.GetInputParameter ("单元序号")
            If dyxh<>"" Then
                  If isnumeric(dyxh)=false Then msgbox "单元序号应为数值XX"  :  Exit Function
            End If
				mjkmc=SSProcess.GetInputParameter ("面积块名称")
				sygn=SSProcess.GetInputParameter ("使用功能名称")
				mjxs=SSProcess.GetInputParameter ("面积系数")
            JR=SSProcess.GetInputParameter ("是否计容")
            JRMJXSZ=SSProcess.GetInputParameter ("计容面积系数")
				if mjxs ="" then mjxs=0
				if JRMJXSZ ="" then JRMJXSZ=0
				if JR="否" then JRMJXSZ=0
				if JR="" then JRMJXSZ=0
				JZMJ = cdbl(JHMJ) * cdbl(mjxs)
				JRMJ = cdbl(JZMJ) * cdbl(JRMJXSZ)
				JZMJ = FormatNumber(JZMJ, 2) 
				JRMJ = FormatNumber(JRMJ, 2) 
				if JR="否" then
					BJRMJ = JZMJ
				else
					BJRMJ = 0
				end if 
				if  SYGN = "住宅"  Then col = RGB(255,0,0)
				if  SYGN = "工业交通仓储"  Then col = RGB(255,255,0)
				if  SYGN = "商业"  Then col = RGB(0,255,0)
				if  SYGN = "教育医疗卫生科研"  Then col = RGB(0,255,255)
				if  SYGN = "文化娱乐体育"  Then col = RGB(0,0,255)
				if  SYGN = "办公"  Then col = RGB(255,0,255)
				if  SYGN = "军事"  Then col = RGB(128,128,128)
				if  SYGN = "未定义面积块"  Then col = RGB(255,255,255)
				if  SYGN = "其他"  Then col = RGB(192,192,192)
				'qsxz=SSProcess.GetInputParameter ("权属性质")
				SSProcess.SetObjectAttr MJKID,"SSObj_Color",col
				SSProcess.SetObjectAttr MJKID,"[MianJKMC]",MJKMC
				SSProcess.SetObjectAttr MJKID,"[GongNYT]",sygn
				SSProcess.SetObjectAttr MJKID,"[TouYMJ]",cdbl(JHMJ)
				SSProcess.SetObjectAttr MJKID,"[MianJXS]",mjxs
				SSProcess.SetObjectAttr MJKID,"[JianZMJ]",cdbl(JZMJ)
				SSProcess.SetObjectAttr MJKID,"[ShiFJR]",JR
				SSProcess.SetObjectAttr MJKID,"[JiRMJ]",cdbl(JRMJ)
				SSProcess.SetObjectAttr MJKID,"[JiRMJXS]",JRMJXSZ
				SSProcess.SetObjectAttr MJKID,"[BuJRMJ]",cdbl(BJRMJ)
           ' SSProcess.SetObjectAttr MJKID,"[MianJKMC],[GongNYT],[MianJXS],[JianZMJ],[ShiFJR],[JiRMJ]",MJKMC & "," & sygn & "," & mjxs& "," & JZMJ &"," & JR&"," & JRMJ
				SSProcess.SaveBufferObjToDatabase
				SSProcess.ExecuteSDLFunction "ssproject,display.redrawextend", Reason 	 
            'LJ=SSProcess.GetInputParameter ("序号自动累加")
				'SSProcess.AddNewObjToSaveObjList


            'If LJ<>"是" Then Exit Function
           ' XH=LJXH(dyxh)
				'mode = 1 '=0 无参数对话框 =1 有参数对话框
				'title="面积块信息录入"
				'SSProcess.ClearInputParameter
				'SSProcess.AddInputParameter "面积块名称", "",0, "住宅,商业,办公,工业交通仓储,教育医疗卫生科研,文化娱乐体育,军事,建筑物主体,地下室,半地下室,地下室出入口,架空层,坡屋顶,场馆看台下的建筑,无围护的场观看台,门厅,大厅,悬挑看台,架空走廊,无围护的架空走廊,立体书库,立体仓库,立体车库,控制室,落地橱窗,凸(飘)窗,室外走廊（挑廊）,檐廊,门斗,门廊,有柱雨篷,无柱雨篷,楼梯间、水箱间、电梯机房,室内楼梯,电梯井,提物井,通风排气竖井,烟道,有顶盖采光井,室外楼梯,阳台,半阳台,车棚,货棚,站台,加油站,收费站,变形缝,设备层,管道层,避难层,不计容构筑物,未定义面积块,其他", ""
				'SSProcess.AddInputParameter "使用功能名称", sygn,0, "住宅,工业交通仓储,商业,教育医疗卫生科研,文化娱乐体育,办公,军事,未定义面积块,其他", ""
				'SSProcess.AddInputParameter "面积块名称", mjkmc,0, "住宅,商业,办公,工业交通仓储,教育医疗卫生科研,文化娱乐体育,军事,地下室,半地下室,地下室出入口,架空层,坡屋顶,场馆看台,无围护的场观看台,门厅,大厅,悬挑看台,架空走廊,无围护的架空走廊,立体书库,立体仓库,立体车库,控制室,落地橱窗,凸(飘)窗,室外走廊（挑廊）,檐廊,门斗,门廊,有柱雨篷,无柱雨篷,楼梯间、水箱间、电梯机房,室内楼梯、电梯井、提物井、通风排气竖井、烟道,有顶盖采光井,室外楼梯,阳台,半阳台,车棚、货棚、站台、加油站、收费站,变形缝,设备层、管道层、避难层,不计容构筑物,未定义面积块,其他", ""
				'SSProcess.AddInputParameter "面积块名称", "",0, "住宅,商业,办公,工业交通仓储,教育医疗卫生科研,文化娱乐体育,军事,地下室,半地下室,地下室出入口,架空层,坡屋顶,场馆看台,无围护的场观看台,门厅,大厅,悬挑看台,架空走廊,无围护的架空走廊,立体书库,立体仓库,立体车库,控制室,落地橱窗,凸(飘)窗,室外走廊（挑廊）,檐廊,门斗,门廊,有柱雨篷,无柱雨篷,楼梯间、水箱间、电梯机房,室内楼梯、电梯井、提物井、通风排气竖井、烟道,有顶盖采光井,室外楼梯,阳台,半阳台,车棚、货棚、站台、加油站、收费站,变形缝,设备层、管道层、避难层,不计容构筑物,未定义面积块,其他", ""	
				'SSProcess.AddInputParameter "使用功能名称", sygn,0, "住宅,工业交通仓储,商业,教育医疗卫生科研,文化娱乐体育,办公,军事,其他", ""
				'SSProcess.AddInputParameter "面积系数", mjxs,0, "", ""
			'	SSProcess.AddInputParameter "面积系数" ,mjxs, 3, "SYS_DROPDOWNLIST", ""
				'SSProcess.AddInputParameter "权属性质", qsxz,0, "私有,共有", ""
				'SSProcess.AddInputParameter "序号自动累加", LJ,0, "是,否", ""
				'SSProcess.AddInputParameter "是否计容", JR,0, "是,否", ""
				'SSProcess.ShowScriptDlg mode,title

				SSProcess.WriteEpsIni "MJKGNXX", "MJKSYGN" ,sygn
				SSProcess.WriteEpsIni "MJKGNXX", "MJKMCZ" ,MJKMC
				SSProcess.WriteEpsIni "MJKGNXX", "SFJRZ" ,JR
				SSProcess.WriteEpsIni "MJKGNXX", "MianJXSZ" ,mjxs
				SSProcess.WriteEpsIni "MJKGNXX", "JiRMJXSZ" ,JRMJXSZ

				End If
End Function

Function LJXH(byval dyxh)
      LJXH=""
      If dyxh="" Then Exit Function
      C=LEN(dyxh)
      dyxh=CInt(dyxh) + 1
      C1=LEN(dyxh)
      LJXH=dyxh
      For i=0 To (C-C1-1)
            LJXH = "0" & LJXH
      Next
End Function
DIM JRMJ
'获取功能用途信息
Function getGNYTXX(DH,XKZBH,GNYTXX,GNMCXX)
	Dim Fvalues(2)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName
	sql = "SELECT JG_建筑物单体建筑面积指标核实信息属性表.GongNLX,GongNMC FROM JG_建筑物单体建筑面积指标核实信息属性表 WHERE (JG_建筑物单体建筑面积指标核实信息属性表.[JianZWMC]) = '"&DH&"' AND (JG_建筑物单体建筑面积指标核实信息属性表.[GuiHXKZBH]) = '"&XKZBH&"';"
	SSProcess.OpenAccessRecordset projectName, sql
	rscount = SSProcess.GetAccessRecordCount (projectName, sql )
	If rscount > 0 Then
		SSProcess.AccessMoveFirst projectName, sql
		while (SSProcess.AccessIsEOF (projectName, sql ) = False)
			SSProcess.GetAccessRecord projectName, sql, fields, values
			SSFunc.ScanString values, ",", Fvalues, FvaluesCount
			if GNYTXX="" then
				GNYTXX=Fvalues(0)
			else
				GNYTXX=GNYTXX&","&Fvalues(0)
			end if
			if GNMCXX="" then
				GNMCXX=Fvalues(1)
			else
				GNMCXX=GNMCXX&","&Fvalues(1)
			end if
			SSProcess.AccessMoveNext projectName, sql 
		Wend
	End If
	SSProcess.CloseAccessRecordset projectName, sql 
	SSProcess.CloseAccessMdb projectName 
End Function

'获取功能名称信息
Function getGNMCXX(DH,XKZBH,GNMCZB)
	Dim Fvalues(2)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName
	sql = "SELECT JG_建筑物单体建筑面积指标核实信息属性表.GongNMC FROM JG_建筑物单体建筑面积指标核实信息属性表 WHERE (JG_建筑物单体建筑面积指标核实信息属性表.[JianZWMC]) = '"&DH&"' AND (JG_建筑物单体建筑面积指标核实信息属性表.[GuiHXKZBH]) = '"&XKZBH&"' And (JG_建筑物单体建筑面积指标核实信息属性表.[GongNLX]) = '公建';"
	SSProcess.OpenAccessRecordset projectName, sql
	rscount = SSProcess.GetAccessRecordCount (projectName, sql )
	If rscount > 0 Then
		SSProcess.AccessMoveFirst projectName, sql
		while (SSProcess.AccessIsEOF (projectName, sql ) = False)
			SSProcess.GetAccessRecord projectName, sql, fields, values
			SSFunc.ScanString values, ",", Fvalues, FvaluesCount
			if GNMCZB="" then
				GNMCZB=Fvalues(0)
			else
				GNMCZB=GNMCZB&","&Fvalues(0)
			end if
			SSProcess.AccessMoveNext projectName, sql 
		Wend
	End If
	SSProcess.CloseAccessRecordset projectName, sql 
	SSProcess.CloseAccessMdb projectName 
End Function




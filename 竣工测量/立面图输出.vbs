'输出立面图对话框
Sub OnInitScript()
	mode = 1 '=0 无参数对话框 =1 有参数对话框
	title="输出立面图....."
	SSProcess.ClearInputParameter 
	SSProcess.AddInputParameter "立面图类型", "平面屋面",  0, "平面屋面,斜面屋面,不同高度屋面", ""
	SSProcess.AddInputParameter "立面位置", "东视",  3,"东视,西视,北视,南视", ""
	SSProcess.UpdateScriptDlgParameter 1  '更新脚本运行对话框参数(不更新到内存)
	SSProcess.ShowScriptDlg mode,title
	SSProcess.WriteEpsIni "CRunScriptDlg2_" & LMtitle,"DlgWidth" , "250" 
	SSProcess.WriteEpsIni "CRunScriptDlg2_" & LMtitle, "DlgHeight" ,"200" 
	SSProcess.WriteEpsIni "CRunScriptDlg2_" & LMtitle, "ColumnWidth" ,"80" 
	SSProcess.RefreshView 
	'LMtitle 和 WriteEpsIni
End Sub
Sub OnExitScript()
	'添加代码
End Sub


Sub OnOK()
		SSProcess.UpdateScriptDlgParameter 1
		LMLX = SSProcess.GetInputParameter ("立面图类型" )
		LMfx = SSProcess.GetInputParameter ("立面位置" )
		SSProcess.UpdateSysSelection 0 '系统选择集内容更新到脚本选择集
		geoCount = SSProcess.GetSelGeoCount()
		if geoCount<>1 then msgbox "请选择竣工测量立面图图廓！":exit sub
			GeoCode= SSProcess.GetSelGeoValue( 0, "SSObj_Code" )
		if GeoCode<>"9400604"  then  msgbox "请选择竣工测量立面图图廓！":exit sub
		If geoCount=1 then 
			LMTTKID = SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
			JZWMC =SSProcess.GetObjectAttr( LMTTKID, "[JianZWMC]") '名称
			JZGHXKZH =SSProcess.GetObjectAttr( LMTTKID, "[GuiHXKZBH]")'编号
			SSProcess.SetObjectAttr LMTTKID, "[LiMTFX]",LMfx 
			SSObj_GroupID=SSProcess.GetGeoMaxID  
		End If

		GetJZWFWXXX JZGHXKZH,JZWMC,jzwzgd,jzwjbgd,neqgd,neqgc,zgdgc,zflbg,dxsgd,dxsdbgc
		SZZLCL dxsgd
		SZZLCL jzwzgd
		SZZLCL jzwjbgd
		SZZLCL neqgd
		SZZLCL zflbg
		SZZLCL zgdgc
		SZZLCL neqgc

		notaise=1


		wjgd=0
		'斜面高度
		ZGzgdgd=CDBL(zgdgc)-CDBL(zflbg)-CDBL(jzwzgd)
		
		If  LMLX = "平面屋面" Then 
			if wjgd <>"" then wjgd=0
			if jzwjbgd <> 0 then jzwjbgd=0
			if neqgc = ""  then neqgc=0
			if jzwjbgd = ""  then jzwjbgd =0
			If neqgd=0 then
				If zgdgc <> 0 and neqgc <> 0 Then
					neqgd = cdbl(zgdgc) - cdbl(neqgc) 
				End If
			End If
			'空、ID、索引、地下室高度、总高度、女儿墙高度、wjgd、局部高度、正负零标高
			Createlimiantu scsndpg,LMTTKID,geoid,dxsgd,jzwzgd,neqgd,wjgd,jzwjbgd,zflbg

		Elseif  LMLX = "斜面屋面" Then 
			if jzwjbgd <> 0 then jzwjbgd=0
			if neqgd <> 0 then neqgd=0
			If wjgd=0 then
				wjgd = cdbl(zgdgc) - cdbl(zflbg)  - cdbl(jzwzgd) 
				wjgd=formatnumber(wjgd,2,-1,0,0) '格式化为两位小数的数字
			End If
			Createlimiantu scsndpg,LMTTKID,geoid,dxsgd,jzwzgd,neqgd,wjgd,jzwjbgd,zflbg
		Elseif  LMLX = "不同高度屋面" Then 
			if wjgd <>"" then wjgd=0
			Createlimiantu scsndpg,LMTTKID,geoid,dxsgd,jzwzgd,neqgd,wjgd,jzwjbgd,zflbg
		End If

		SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
		SSProcess.ExecuteSDLFunction "$SDL.SSWorkSpace.Display.RedrawExtend", 0
		SSProcess.RefreshView 
End Sub


function Createlimiantu(scsndpg,TKid,geoid,Dxsgd,Zgd,Neqgd,Wjgd,Deztgd,zflbg)
		SSProcess.GetObjectPoint TKid, 0, x0,  y0,  z,  ptype,  name 
		SSProcess.GetObjectPoint TKid, 1, x1,  y1,  z,  ptype,  name 
		SSProcess.GetObjectPoint TKid, 2, x2,  y2,  z,  ptype,  name 
		TKheigh = cdbl(y2)-cdbl(y0)
		TKwidgh =cdbl(x1)-cdbl(x0)

		jdY0=cdbl(y0)+TKheigh*1/3:jdY2=cdbl(y2)-TKheigh*1/10:jdX0=x0+TKwidgh*25/92:jdX1=x1-TKwidgh*25/92
		kuandu=abs(jdX1-jdX0):gaodu=abs(jdY2-jdY0)
		jdYY=jdY0
		if Dxsgd="" then Dxsgd=0
		if Zgd="" then Zgd=0
		if Neqgd="" then Neqgd=0
		if Wjgd="" then Wjgd=0
		if Deztgd="" then Deztgd=0
		Lzg=cdbl(Dxsgd)+cdbl(Zgd)+cdbl(Neqgd)+cdbl(Wjgd)
		Ztgdc=cdbl(Zgd)-cdbl(Deztgd)
		IF Lzg<> 0 THEN blxs=gaodu/Lzg
		TMDXSGD=Dxsgd*blxs  '图面地下室高度
		TMZGD=Zgd*blxs  '图面总高度
		TMNEQGD=Neqgd*blxs  '图面女儿墙高度
		TMZLG=Lzg*blxs '图面总楼高
		TMWJG=Wjgd*blxs '图面屋脊高度
		TMDEZTGD=Deztgd*blxs '图面第二主体高度
		TMZTGDC=Ztgdc*blxs '图面主体高度差
		If Deztgd<> 0 then
	'****************绘制地下室线*********************
			If Dxsgd <> 0 then
				addline jdX0,jdY0,jdX1, jdY0,"9400506",geoid'绘制地下室第一根线
				addline jdX0, jdY0+TMDXSGD,jdX0,jdY0,"9400506",geoid'绘制左竖线
				addline jdX1,jdY0,jdX1, jdY0+TMDXSGD,"9400507",geoid'绘制右竖线
				addnote "地下室底板",jdX0-5.5,jdY0-1,250*notaise,250*notaise,"黑体",geoid

	'左边标注
				addline jdX0-2.5,jdY0,jdX0-2.5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注线
				addline jdX0,jdY0,jdX0-5, jdY0,"940050602",geoid'绘制地下高度距离标注横线
				addline jdX0,jdY0+TMDXSGD,jdX0-5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注横线
				addPoint "9400504",jdX0-2.5,jdY0+((cdbl(TMDXSGD))/2),"",Dxsgd,geoid
				addline jdX0-2.5-1,jdY0-0.5,jdX0-2.5+1, jdY0+0.5,"940050602",geoid'绘制总高度标注短线
				addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线

	'右边标注
				addline jdX1+2.5,jdY0,jdX1+2.5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注线
				addline jdX1,jdY0,jdX1+5,jdY0,"940050602",geoid'绘制地下高度距离标注横线
				addline jdX1,jdY0+TMDXSGD,jdX1-5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注横线
				addPoint "9400504",jdX1+5,jdY0+((cdbl(TMDXSGD))/2),"地下总高度",Dxsgd,geoid
				addline jdX1+2.5+1,jdY0-0.5,jdX1+2.5-1, jdY0+0.5,"940050602",geoid'绘制总高度标注短线
				addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线

			End If
	'****************绘制总高度线*********************

			If Dxsgd<> 0 then 
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"940050602",geoid'绘制±0标高线
			Else
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"9400506",geoid'绘制±0标高线
			End If
			addPoint "9400502",jdX0-5,jdY0+TMDXSGD,zflbg, "",geoid '正负零点
			addPoint "9400502",jdX1+5,jdY0+TMDXSGD, zflbg,"",geoid'正负零点


			addline jdX0,jdY0,jdX0, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'绘制总高度左线
			addline jdX1,jdY0,jdX1, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制总高度右线


			addline jdX0,jdY0+TMDXSGD+TMDEZTGD,jdX0+2/5*kuandu, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'绘制第二高度横线
			addline jdX0+2/5*kuandu,jdY0+TMDXSGD+TMDEZTGD,jdX0+2/5*kuandu, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制高度竖直线
			addline jdX0+2/5*kuandu,jdY0+TMDXSGD+TMZGD,jdX1,jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制高度竖直线


			addline jdX0-2.5,jdY0+TMDXSGD,jdX0-2.5, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'绘制左侧总高度距离标注线
			addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMDEZTGD))/2),"",Deztgd,geoid
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))-0.5,jdX0+1-2.5, jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))+0.5,"940050602",geoid'绘制总高度标注短线
			addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0+1-2.5, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线
			addline jdX0-5,jdY0+TMDXSGD,jdX0, jdY0+TMDXSGD,"940050602",geoid'绘制总高度高度标注横线


			addline jdX0-2.5,jdY0+TMDXSGD+TMDEZTGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制左侧高度差距离标注线
			addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMDEZTGD)+(cdbl(TMZTGDC))/2),"",Ztgdc,geoid
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))-0.5,jdX0-2.5+1, jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))+0.5,"940050602",geoid'绘制总高度标注短线
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX0-2.5+1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'绘制总高度高度标注短线

			addline jdX0,jdY0+TMDXSGD+TMZGD,jdX0-5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制第一高度檐口线
			addnote "楼顶",jdX0-5.5,jdY0+TMDXSGD+TMZGD+1,250*notaise,250*notaise,"黑体",geoid
			addline jdX0,jdY0+TMDXSGD+TMDEZTGD,jdX0-5, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'绘制第二高度檐口线
			addnote "楼顶",jdX0-5.5,jdY0+TMDXSGD+TMDEZTGD+1,250*notaise,250*notaise,"黑体",geoid


			addline jdX1+2.5,jdY0+TMDXSGD,jdX1+2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制总高度距离标注线
			addline jdX1,jdY0+TMDXSGD+TMZGD,jdX1+5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制女儿墙高度右线
			addPoint "9400504",jdX1+2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMZGD))/2),Zgd, "地上总高度",geoid
			addline jdX1+2.5+1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX1+2.5-1, jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'绘制总高度标注短线
			addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线



	'****************绘制女儿墙*********************
			If Neqgd<> 0 then 
				addline jdX0+2/5*kuandu,jdY0+TMDXSGD+TMZGD,jdX0+2/5*kuandu, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线
				addline jdX1,jdY0+TMDXSGD+TMZGD,jdX1,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线

				addline jdX0-2.5,jdY0+TMDXSGD+TMZGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度标注线
				addline jdX0, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线
				addline jdX0-2.5-1, jdY0+TMDXSGD+TMZGD+TMNEQGD-0.5,jdX0-2.5+1,  jdY0+TMDXSGD+TMZGD+TMNEQGD+0.5,"940050602",geoid'绘制女儿墙高度标注短线
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+0.5,"940050602",geoid'绘制女儿墙高度标注短线
				addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMDEZTGD)+cdbl(TMZTGDC)+(cdbl(TMNEQGD)/2)),"",Neqgd,geoid

				addline jdX0, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙左边顶部横线
				addnote "女儿墙顶",jdX0-5.5, jdY0+TMDXSGD+TMZGD+TMNEQGD+1,250*notaise,250*notaise,"黑体",geoid



				addline jdX0+2/5*kuandu, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0+2/5*kuandu+5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线横线
				addline jdX1, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX1-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线横线
				addline jdX0+2/5*kuandu+5,jdY0+TMDXSGD+TMZGD,jdX0+2/5*kuandu+5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线竖线
				addline jdX1-5,jdY0+TMDXSGD+TMZGD,jdX1-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线竖线




			'绘制第二高度女儿墙

				addline jdX0,jdY0+TMDXSGD+TMDEZTGD,jdX0, jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线
				addline jdX0,jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,jdX0+5, jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线横线
				addline jdX0+5,jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,jdX0+5, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'绘制女儿墙高度左线竖线
			End If

		Else
	'****************绘制地下室线*********************
			If Dxsgd <> 0 then
				addline jdX0,jdY0,jdX1, jdY0,"9400506",geoid'绘制地下室第一根线
				addline jdX0, jdY0+TMDXSGD,jdX0,jdY0,"9400506",geoid'绘制左竖线
				addline jdX1,jdY0,jdX1, jdY0+TMDXSGD,"9400507",geoid'绘制右竖线
				addnote "地下室底板",jdX0-5.5,jdY0-1,250*notaise,250*notaise,"黑体",geoid

	'左边标注
				addline jdX0-2.5,jdY0,jdX0-2.5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注线
				addline jdX0,jdY0,jdX0-5, jdY0,"940050602",geoid'绘制地下高度距离标注横线
				addline jdX0,jdY0+TMDXSGD,jdX0-5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注横线
				addPoint "9400504",jdX0-2.5,jdY0+((cdbl(TMDXSGD))/2),"",Dxsgd,geoid
				addline jdX0-2.5-1,jdY0-0.5,jdX0-2.5+1, jdY0+0.5,"940050602",geoid'绘制总高度标注短线
				addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线

	'右边标注
				addline jdX1+2.5,jdY0,jdX1+2.5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注线
				addline jdX1,jdY0,jdX1+5,jdY0,"940050602",geoid'绘制地下高度距离标注横线
				addline jdX1,jdY0+TMDXSGD,jdX1-5, jdY0+TMDXSGD,"940050602",geoid'绘制地下高度距离标注横线
				addPoint "9400504",jdX1+5,jdY0+((cdbl(TMDXSGD))/2),"地下总高度",Dxsgd,geoid
				addline jdX1+2.5+1,jdY0-0.5,jdX1+2.5-1, jdY0+0.5,"940050602",geoid'绘制总高度标注短线
				addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线

			End If
	'****************绘制总高度线*********************

			If Dxsgd<> 0 then 
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"940050602",geoid'绘制±0标高线
			Else
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"9400506",geoid'绘制±0标高线
			End If
			addPoint "9400502",jdX0-5,jdY0+TMDXSGD, zflbg,"",geoid '正负零点
			addPoint "9400502",jdX1+5,jdY0+TMDXSGD, zflbg, "",geoid'正负零点


			addline jdX0,jdY0,jdX0, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制总高度左线
			addline jdX1,jdY0,jdX1, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制总高度右线
			addline jdX0-5,jdY0+TMDXSGD+TMZGD,jdX1+5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制檐口线
			addnote "檐顶",jdX0-5.5,jdY0+TMDXSGD+TMZGD-1,250*notaise,250*notaise,"黑体",geoid

 '******************绘制总高度右标注*********************
			addline jdX1+2.5,jdY0+TMDXSGD,jdX1+2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制总高度距离标注线
			addPoint "9400504",jdX1+2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMZGD))/2),Zgd, "地上总高度",geoid
			addline jdX1+2.5+1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX1+2.5-1, jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'绘制总高度标注短线
			addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线

 '******************绘制总高度左标注*********************
			addline jdX0-2.5,jdY0+TMDXSGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'绘制总高度距离标注线
			addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMZGD))/2), "",Zgd,geoid
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX0-2.5+1, jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'绘制总高度标注短线
			addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+0.5,"940050602",geoid'绘制总高度高度标注短线




	'****************绘制女儿墙*********************
			If Neqgd<> 0 then 
				addline jdX0,jdY0+TMDXSGD+TMZGD,jdX0, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线
				addline jdX1,jdY0+TMDXSGD+TMZGD,jdX1, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线
				addline jdX0-2.5,jdY0+TMDXSGD+TMZGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度标注线
				addline jdX0,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD+TMNEQGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+TMNEQGD+0.5,"940050602",geoid'绘制女儿墙高度标注短线
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+0.5,"940050602",geoid'绘制女儿墙高度标注短线
				addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMZGD)+(cdbl(TMNEQGD))/2), "",Neqgd,geoid
				addline jdX0,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙左边顶部横线
				addnote "女儿墙顶",jdX0-5.5,jdY0+TMDXSGD+TMZGD+TMNEQGD+1,250*notaise,250*notaise,"黑体",geoid

				addline jdX0,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0+5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线横线
				addline jdX1,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX1-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线横线

				addline jdX0+5,jdY0+TMDXSGD+TMZGD,jdX0+5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度左线竖线
				addline jdX1-5,jdY0+TMDXSGD+TMZGD,jdX1-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'绘制女儿墙高度右线竖线
			End If

			If Wjgd<> 0 then
				WJZXDXZB=jdX0+CDBL(kuandu/2) '屋脊中心点X坐标
				addline jdX0,jdY0+TMDXSGD+TMZGD,WJZXDXZB, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'绘制屋脊高度左线
				addline jdX1,jdY0+TMDXSGD+TMZGD,WJZXDXZB, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'绘制屋脊高度右线		
				
				addline jdX0,jdY0+TMDXSGD+TMZGD+TMWJG,jdX0-5, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'绘制屋脊高度标注横线
				
				addline jdX0-2.5,jdY0+TMDXSGD+TMZGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'绘制屋脊高度标注线
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD+TMWJG-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+TMWJG+0.5,"940050602",geoid'绘制屋脊墙高度标注短线
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+0.5,"940050602",geoid'绘制屋脊墙高度标注短线
				addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMZGD)+(cdbl(TMWJG))/2),"",Wjgd,geoid

				addnote "屋脊",jdX0-5.5,jdY0+TMDXSGD+TMZGD+TMWJG+1,250*notaise,250*notaise,"黑体",geoid
			End If
		End If

end function 




function addline(x00,y00,x11,y11,code,geoid)
		if code="" then code="1"
		SSProcess.CreateNewObj 1
		SSProcess.SetNewObjValue "SSObj_Code", code
		SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
		SSProcess.SetNewObjValue "SSObj_DataMark", "竣工测量立面图信息"
		SSProcess.SetNewObjValue "SSObj_GroupID ", GeoID
		SSProcess.AddNewObjPoint x00, y00, 0, 0, ""
		SSProcess.AddNewObjPoint x11, y11, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function
function addPoint(Code,x00,y00,sxz1, sxz2,geoid)
		SSProcess.CreateNewObj 0
		SSProcess.SetNewObjValue "SSObj_Code",Code
		SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
		SSProcess.SetNewObjValue "SSObj_DataMark", "竣工测量立面图信息"
		SSProcess.SetNewObjValue "[GuiHSPZ]", sxz1
		SSProcess.SetNewObjValue "[JunGCLZ]", sxz2
		SSProcess.AddNewObjPoint x00, y00, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SetNewObjValue "SSObj_GroupID ",GeoID
		SSProcess.SaveBufferObjToDatabase
		if Code="9400503" or Code="9400504" then
				SSProcess.ExplodeObj SSProcess.GetGeoMaxID  , 0, 1, "" 
		end if
end function
function addPoint1(Code,x00,y00,sxz1,geoid)
		SSProcess.CreateNewObj 0
		SSProcess.SetNewObjValue "SSObj_Code",Code
		SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
		SSProcess.SetNewObjValue "SSObj_DataMark", "竣工测量立面图信息"
		SSProcess.SetNewObjValue "SSObj_GroupID ", GeoID
		SSProcess.SetNewObjValue "[BiaoZNR]", sxz1
		SSProcess.AddNewObjPoint x00, y00, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
		if Code="9400503" or Code="9400504" then
				SSProcess.ExplodeObj SSProcess.GetGeoMaxID  , 0, 1, "" 
		end if
end function
function addnote(nr,xx,yy,zg,zk,zt,geoid)
		SSProcess.CreateNewObj 3
		SSProcess.SetNewObjValue "SSObj_FontClass", "0"
		SSProcess.SetNewObjValue "SSObj_FontString", nr
		SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
		SSProcess.SetNewObjValue "SSObj_DataMark", "竣工测量立面图信息"
		SSProcess.SetNewObjValue "SSObj_GroupID ", GeoID
		SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
		SSProcess.SetNewObjValue "SSObj_FontName", zt
		SSProcess.SetNewObjValue "SSObj_FontWidth", zk
		SSProcess.SetNewObjValue "SSObj_FontHeight", zg
		SSProcess.AddNewObjPoint xx,yy, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 

'获取建筑物信息
Function GetJZWFWXXX(ghxkzbh,jzwmc,jzwzgd,jzwjbgd,neqgd,neqgc,zgdgc,zflbg,dxsgd,dxsdbgc)
	dim Fvalues(10)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName 
	sql = "SELECT JG_新建建筑物范围线信息属性表.JianZWZGD,JianZWJBZGD,NvEQGD,NvEQGC,JianZDBZGCGC,ZhengFLBG,DiXGD,DiXSDBGC FROM JG_新建建筑物范围线信息属性表 INNER JOIN GeoAreaTB ON JG_新建建筑物范围线信息属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 and ([JG_新建建筑物范围线信息属性表].[ID] > 0 And ([JG_新建建筑物范围线信息属性表].[JianZWMC] = '"&jzwmc&"') And ([JG_新建建筑物范围线信息属性表].[GuiHXKZBH] = '"&ghxkzbh&"'));"
	SSProcess.OpenAccessRecordset projectName, sql 
	rscount = SSProcess.GetAccessRecordCount (projectName, sql )
	If rscount >  0 Then
		SSProcess.AccessMoveFirst projectName, sql
		while (SSProcess.AccessIsEOF (projectName, sql ) = False)
			SSProcess.GetAccessRecord projectName, sql, fields, values
			SSFunc.Scanstring values,",",Fvalues,Fvaluescount		
			jzwzgd=Fvalues(0) '总高度
			jzwjbgd=Fvalues(1) '局部高度
			neqgd=Fvalues(2) '女儿墙高度
			neqgc=Fvalues(3) '女儿墙高程
			zgdgc=Fvalues(4) '最高点高程
			zflbg=Fvalues(5) '正负零标高
			dxsgd=Fvalues(6) '地下室高度
			dxsdbgc=Fvalues(7) '地下室底板高程
			SSProcess.AccessMoveNext projectName, sql 
		Wend
	End If
	SSProcess.CloseAccessRecordset projectName, sql 
	SSProcess.CloseAccessMdb projectName 
End Function

'数值置零处理
Function SZZLCL (SJvalue)
		If  SJvalue="" and SJvalue="NULL"   and SJvalue="*"  then SJvalue=0
		If Isnumeric(SJvalue) = False Then SJvalue = 0
End Function






'*************************************************************************************
'Copyright (c) 2018-2019 Kevin Yang. All Rights Reserved.
'Tel.15357565878           E-mail.1402565009@qq.com
'Origin:Kevin Yang,20181024-02:00
'*************************************************************************************
Dim wApp,docWord,projectName
zongneirong=""
Sub OnClick()
	Dim arArray(100),count
	'&&&&&&&&&&==========<创建Word>==========&&&&&&&&&&
	SSProcess.PushUndoMark 
   'pathname=SSProcess.GetSysPathName(5)
	Pathname1=SSProcess.GetSysPathName(5)&"联合测绘"
	set folder=createobject("scripting.filesystemobject")
	If CheckFloderExists(Pathname1) = False Then 
		set fld=folder.createfolder(Pathname1)
	end if
	pathname2 = Pathname1&"\"&"绿地"
		If CheckFloderExists(pathname2) = False Then 
		set fld=folder.createfolder(pathname2)
	end if
	pathname=pathname2&"\"
	If pathname = "" Then Exit Sub
	'If ExportWord_LD_CheckInfo(pathname)  =  false then  msgbox zongneirong&"数据目录结构错误，请创建目录结构，本次取消输出！" : exit sub
  ' msg = msgbox("是否删除输出目录下已有文件？",4,"提示：")
	'If msg = 6 Then
		'	DeleteAllFiles pathname&"\绿地测量\成果数据", "docx"'删除dwg  
	'Else
			'msgbox "取消输出":exit sub
'	End if 
	WordTemplatePath = SSProcess.GetSysPathName(7)&"建筑工程竣工绿地测量成果报告书.docx"
	'WordSavePath = pathname&"\绿地测量\成果数据\"
	'WordSavePath = pathname&"\绿地测量\成果数据\"
	WordSavePath = pathname
	WordSaveName = "建筑工程竣工绿地测量成果报告书.docx"
	CreatWordbyTemplate WordTemplatePath,WordSavePath,WordSaveName

	'&&&&&&&&&&==========<打开Access>==========&&&&&&&&&&
	projectName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb projectName
	'&&&&&&&&&&==========<Tbale4 绿地测量成果表>==========&&&&&&&&&&
	Table4
	'NONEED&&&&&&&&&&==========<Tbale3 测量说明>==========&&&&&&&&&&
	'Table3
	'PASS&&&&&&&&&&==========<Tbale2 测绘责任人>==========&&&&&&&&&&
	Table2
	'PASS&&&&&&&&&&==========<Tbale1  封面&Replace>==========&&&&&&&&&&
	Table1
  	'&&&&&&&&&&==========<打印图片>==========&&&&&&&&&&
    table_5()
	'&&&&&&&&&&==========<关闭Word>==========&&&&&&&&&&
	CloseWord WordSavePath&WordSaveName
	'&&&&&&&&&&==========<关闭Access>==========&&&&&&&&&&
   SSProcess.CloseAccessMdb projectName
End Sub

'&&&&&&&&&&==========<Tbale1 封面&Replace>==========&&&&&&&&&&

'// 输出封面


Function Table1()

  GetYDHXinfo HXYDBH,HXYDMJ,HXXMMC,HXXMDZ,HXSJDW,HXWTDW,HXJSDW,HXCHDW,HXCHSJ,HXBZR,HXJCR,HXSHR
	ReplaceOneStr "{GCBH}",HXYDBH'建设用地项目编号
	ReplaceOneStr "{JSDW}",HXJSDW'建设单位
	ReplaceOneStr "{XMMC}",HXXMMC'项目名称
	ReplaceOneStr "{XMDZ}",HXXMDZ'项目地址
	ReplaceOneStr "{SJDW}",HXSJDW'设计单位
	ReplaceOneStr "{WTDW}",HXWTDW'委托单位
	ReplaceOneStr "{HXH}",HXH
	ReplaceOneStr "{CHDW}",HXCHDW '测绘单位
	ReplaceOneStr "{CHSJ}",HXCHSJ'测绘时间
	ReplaceOneStr "{BZZ}",HXBZR'编制人
	ReplaceOneStr "{JCZ}",HXJCR'检查人
	ReplaceOneStr "{SHZ}",HXSHR'审核人
	ReplaceOneStr "{年}",year(NOW)
	ReplaceOneStr "{月}",Month(NOW)
	ReplaceOneStr "{日}",day(NOW)
End Function 

Function GetYDHXinfo(HXYDBH,HXYDMJ,HXXMMC,HXXMDZ,HXSJDW,HXWTDW,HXJSDW,HXCHDW,HXCHSJ,HXBZR,HXJCR,HXSHR)
               SSProcess.PushUndoMark 
	            SSProcess.ClearSelection 
					SSProcess.ClearSelectCondition 
					SSProcess.SetSelectCondition "SSObj_Code", "==", "9103900"
					SSProcess.SelectFilter
					HXCount = SSProcess.GetSelGeoCount()
              if HXCount=1 then
               for i=0 to HXCount-1 
			        	    HXID= SSProcess.GetSelGeoValue(0,"SSObj_ID")
                      HXYDBH=SSProcess.GetSelGeoValue( 0, "[工程编号]")
                      HXYDMJ=SSProcess.GetSelGeoValue( 0, "[规划用地面积]")
                      HXXMMC=SSProcess.GetSelGeoValue(0, "[项目名称]")
                      HXXMDZ=SSProcess.GetSelGeoValue( 0, "[项目地址]")
                      HXSJDW=SSProcess.GetSelGeoValue(0, "[设计单位]")
                      HXWTDW=SSProcess.GetSelGeoValue( 0, "[委托单位]")
                      HXJSDW=SSProcess.GetSelGeoValue( 0, "[建设单位]")
                      HXCHDW=SSProcess.GetSelGeoValue( 0, "[测绘单位]")
                      HXCHSJ=SSProcess.GetSelGeoValue( 0, "[测绘时间]")
                      HXBZR=SSProcess.GetSelGeoValue( 0, "[编制]")
                      HXJCR=SSProcess.GetSelGeoValue( 0, "[检查]")
                      HXSHR=SSProcess.GetSelGeoValue( 0, "[审核]")
                      next
               else 
                msgbox "数据中存在多条用地红线,请监检查数据！"
               end if 
End Function 

'&&&&&&&&&&==========<Tbale1  封面&Replace>==========&&&&&&&&&&

'&&&&&&&&&&==========<Tbale2 测绘责任人>==========&&&&&&&&&&
Function Table2()
	Dim RYXX(1000),SinRYXX(3)
	TableIndex = 2
	RYCount = 0
	'GetRYXX RYXX,RYCount
	If RYCount > 0 Then 
			For i  = 0 to RYCount - 1
					Erase SinRYXX
					SSFunc.Scanstring RYXX(i),",",SinRYXX,SinRYXXCount
					WriteCell TableIndex,(3 + i),2,SinRYXX(0)
					WriteCell TableIndex,(3 + i),3,SinRYXX(1)
					WriteCell TableIndex,(3 + i),4,SinRYXX(2)
			Next
	End If
End Function 

Function GetRYXX(RYXX,RYCount)
	RYCount = 0
	sql = "SELECT 姓名,上岗证书编号或职业资格证书号,备注 FROM 人员信息属性表 WHERE ([人员信息属性表].[作业业务类型] = '绿地测量');"
	SSProcess.OpenAccessRecordset projectName,sql
	rscount = SSProcess.GetAccessRecordCount(projectName,sql)
	If rscount > 0 Then
			SSProcess.AccessMoveFirst projectName,sql
			While (SSProcess.AccessIsEOF (projectName, sql) = False)
					SSProcess.GetAccessRecord projectName, sql, fields, values
					RYXX(RYCount) = values
					RYCount = RYCount + 1
					SSProcess.AccessMoveNext projectName, sql
			Wend
	End If
	SSProcess.CloseAccessRecordset projectName,sql
End Function 
'&&&&&&&&&&==========<Tbale2 测绘责任人>==========&&&&&&&&&&

'&&&&&&&&&&==========<Tbale4 绿地测量成果表>==========&&&&&&&&&&
Function Table4()
	Dim LDBHJMJ(10000,25),Sum(25),LDvArray(10000),ALLLDIDS(10000),ALLLDBH(10000)
	TableIndex =4:Page = 6
	CopyPage Page
	StartLine = 7

	GetGreenMapID GreenMapID
	LDFWXIDs = SSProcess.SearchInnerObjIDs(GreenMapID, 2, "9104901", 0)
     IF LDFWXIDs<>""Then
     SSFunc.ScanString LDFWXIDs, ",",LDvArray, LDnCount

     if LDnCount>7 then 
       crgs= LDnCount-7
		wApp.ActiveDocument.Tables(TableIndex).Cell(13,28).Select
	   wApp.Selection.MoveRight 1, 1, 1
      Selection.InsertRows crgs
     end if 

     for i=0 to LDnCount
      ALLLDIDS(i)=SSProcess.GetObjectAttr( LDvArray(i), "SSObj_ID")
      ALLLDBH(i)=SSProcess.GetObjectAttr( LDvArray(i), "[绿地编号]")
     next
     end if 
     SSFunc.SortArrayByValue ALLLDIDS,ALLLDBH,LDnCount,1,0 '地块编码从小到大进行排序
     bfh="%": ALLZSMJ="":DKDMLDMJ="":DKDXSMJ="":DKWDMJ="":JZDMLDMJ="":JZDXSMJ=""
     for j=0 to LDnCount -1
      DDLDID=SSProcess.GetObjectAttr( ALLLDIDS(j), "SSObj_ID")
      DDLDAREA=SSProcess.GetObjectAttr( ALLLDIDS(j), "SSObj_Area")
      DDLDBH=SSProcess.GetObjectAttr( ALLLDIDS(j), "[绿地编号]")
      DDLDLB=SSProcess.GetObjectAttr( ALLLDIDS(j), "[绿地类别]")
      DDLDXZ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[绿地性质]")
      DDLDMJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[绿地面积]")'去除景观园林面积
      DDLDMJ=Round(cdbl(DDLDMJ),2)
      DDLDZSHMJ=SSProcess.GetObjectAttr( ALLLDIDS(j),"[折算后绿地面积]")
      DDLDZSHMJ=Round(cdbl(DDLDZSHMJ),2)
      DDLCXQMJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[除休憩后绿地面积]")
      DDLCXQMJ=Round(cdbl(DDLCXQMJ),2)
      DDLDFTHD=SSProcess.GetObjectAttr( ALLLDIDS(j), "[覆土厚度]")
      DDLDZSXJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[折算系数]")
      DDLDJGMJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[景观水体与园路及园林铺装面积]")
      DDLDJGMJ=Round(cdbl(DDLDJGMJ),2)
         if DDLDXZ="单块" then 
              if DDLDLB="地面绿化" then
					 WriteCell TableIndex,(StartLine),1,DDLDBH
					 WriteCell TableIndex,(StartLine),2,DDLDMJ
					 WriteCell TableIndex,(StartLine),3,DDLDJGMJ
					 WriteCell TableIndex,(StartLine),4,(round(cdbl(DDLDJGMJ)/cdbl(DDLDMJ),2))*100&""&bfh
					 WriteCell TableIndex,(StartLine),5,DDLDZSHMJ

					  if DKDMLDMJ="" THEN
					  DKDMLDMJ=DDLDZSHMJ
					  ELSE 
					  DKDMLDMJ=cdbl(DKDMLDMJ)+cdbl(DDLDZSHMJ)
					  end if 

                elseif DDLDLB="地下室及半地下室顶绿化" then 
					 WriteCell TableIndex,(StartLine),1,DDLDBH
                 if DDLDFTHD>1.5 THEN 
					 WriteCell TableIndex,(StartLine),6,DDLDZSHMJ
                elseif 1<DDLDFTHD< 1.5 then
					 WriteCell TableIndex,(StartLine),7,DDLDZSHMJ
                end if 
					 WriteCell TableIndex,(StartLine),8,DDLDJGMJ
					 WriteCell TableIndex,(StartLine),9,(round(cdbl(DDLDJGMJ)/cdbl(DDLDMJ),2))*100&""&bfh
					 WriteCell TableIndex,(StartLine),10,DDLDZSHMJ

					  if DKDXSMJ="" THEN
					  DKDXSMJ=DDLDZSHMJ
					  ELSE 
					  DKDXSMJ=cdbl(DKDXSMJ)+cdbl(DDLDZSHMJ)
					  end if 

              elseif DDLDLB="屋顶绿化" then 
					 WriteCell TableIndex,(StartLine),1,DDLDBH
                  if DDLDFTHD> 1.5 then
					    WriteCell TableIndex,(StartLine),11,DDLDZSHMJ
               elseif 1<DDLDFTHD<1.5  then
					 WriteCell TableIndex,(StartLine),12,DDLDZSHMJ
               elseif 0.<DDLDFTHD<1  then
					 WriteCell TableIndex,(StartLine),13,DDLDZSHMJ
               elseif 0.3<DDLDFTHD<0.5 then
					 WriteCell TableIndex,(StartLine),14,DDLDZSHMJ
                 elseif DDLDFTHD<0.3 then
					 WriteCell TableIndex,(StartLine),15,DDLDZSHMJ
                end if 
					 WriteCell TableIndex,(StartLine),16,DDLDJGMJ
					 WriteCell TableIndex,(StartLine),17,(round(cdbl(DDLDJGMJ)/cdbl(DDLDMJ),2))*100&""&bfh
					 WriteCell TableIndex,(StartLine),18,DDLDZSHMJ

					  if DKWDMJ="" THEN
					  DKWDMJ=DDLDZSHMJ
					  ELSE 
					  DKWDMJ=cdbl(DKWDMJ)+cdbl(DDLDZSHMJ)
					  end if 
              end if 
         end if 
          IF DDLDXZ="集中" then 
              if DDLDLB="地面绿化" then
					 WriteCell TableIndex,(StartLine),1,DDLDBH
					 WriteCell TableIndex,(StartLine),19,DDLDMJ
					 WriteCell TableIndex,(StartLine),20,DDLDJGMJ
					 WriteCell TableIndex,(StartLine),21,(round(cdbl(DDLDJGMJ)/cdbl(DDLDMJ),2))*100&""&bfh
					 WriteCell TableIndex,(StartLine),22,DDLDZSHMJ
						  if JZDMLDMJ=""THEN
						  JZDMLDMJ=DDLDZSHMJ
						  ELSE 
						  JZDMLDMJ=cdbl(JZDMLDMJ)+cdbl(DDLDZSHMJ)
						  end if
						 elseif DDLDLB="地下室及半地下室顶绿化" then 
						 WriteCell TableIndex,(StartLine),1,DDLDBH
						 if DDLDFTHD>1.5 THEN 
						 WriteCell TableIndex,(StartLine),23,DDLDZSHMJ
						 elseif  1<DDLDFTHD<1.5 then
						 WriteCell TableIndex,(StartLine),24,DDLDZSHMJ
						 end if 
						 WriteCell TableIndex,(StartLine),25,DDLDJGMJ
						 WriteCell TableIndex,(StartLine),26,(round(cdbl(DDLDJGMJ)/cdbl(DDLDMJ),2))*100&""&bfh
						 WriteCell TableIndex,(StartLine),27,DDLDZSHMJ
						  if JZDXSMJ="" THEN
						  JZDXSMJ=DDLDZSHMJ
						  ELSE 
						  JZDXSMJ=cdbl(JZDMLDMJ)+cdbl(DDLDZSHMJ)
						  end if
              end if 
         end if 
					 WriteCell TableIndex,(StartLine),28,DDLDZSHMJ
                StartLine=StartLine+1

					  if ALLZSMJ="" THEN
					  ALLZSMJ=DDLDZSHMJ
					  ELSE 
					  ALLZSMJ=cdbl(ALLZSMJ)+cdbl(DDLDZSHMJ)
					  end if 
     next

  GetYDHXinfo HXYDBH,HXYDMJ,HXXMMC,HXXMDZ,HXSJDW,HXWTDW,HXJSDW,HXCHDW,HXCHSJ,HXBZR,HXJCR,HXSHR
   if  LDnCount< 8 then 
   WriteCell TableIndex,14,3,DKDMLDMJ
   WriteCell TableIndex,14,5,DKDXSMJ
   WriteCell TableIndex,14,7,DKWDMJ
   WriteCell TableIndex,14,9,JZDMLDMJ
   WriteCell TableIndex,14,11,JZDXSMJ
   WriteCell TableIndex,15,2,ALLZSMJ
   if HXYDMJ="" then msgbox "请确认用地红线中【规划用地面积】字段是否填写"
   WriteCell TableIndex,16,2,HXYDMJ'??????????????????????????????  从信息录入属性表中取总用地面积
   WriteCell TableIndex,17,2,(round(cdbl(ALLZSMJ)/cdbl(HXYDMJ),4))*100&""&bfh
   else
   WriteCell TableIndex,7+LDnCount,3,DKDMLDMJ
   WriteCell TableIndex,7+LDnCount,5,DKDXSMJ
   WriteCell TableIndex,7+LDnCount,7,DKWDMJ
   WriteCell TableIndex,7+LDnCount,9,JZDMLDMJ
   WriteCell TableIndex,7+LDnCount,11,JZDXSMJ
   WriteCell TableIndex,8+LDnCount,2,ALLZSMJ
   if HXYDMJ="" then msgbox "请确认用地红线中【规划用地面积】字段是否填写"
   WriteCell TableIndex,9+LDnCount,2,HXYDMJ'??????????????????????????????  从信息录入属性表中取总用地面积
   WriteCell TableIndex,10+LDnCount,2,(round(cdbl(ALLZSMJ)/cdbl(HXYDMJ),4))*100&""&bfh
   end if 
End Function

'综合竣工图
Function table_5()
Dim H,W
              tableIndex =5
					SSProcess.ClearSelection 
					SSProcess.ClearSelectCondition 
					SSProcess.SetSelectCondition "SSObj_Code", "==", "9104903"
					SSProcess.SelectFilter
					CountZHTK = SSProcess.GetSelGeoCount()
               IF  CountZHTK=1 THEN
					For k = 0  to CountZHTK-1
								idZH= SSProcess.GetSelGeoValue(k,"SSObj_ID")
					Next
               end if 

					PrintPaper = SSProcess.GetObjectAttr(idZH,"[打印纸张]")
					PrintScale = SSProcess.GetObjectAttr(idZH,"[打印比例]")
					SSProcess.SetMapScale PrintScale 
					wApp.ActiveDocument.Tables(tableIndex).Cell(1,1).Select       '图片
					if PrintPaper = "A4横向" then
						BaseHeith=105
					   BaseWidth=148.5
					   strPaperSize="297X210"
						SetPaper 2,2,1,1,29.7,21
						H=16.2: W=22.9
					elseif PrintPaper = "A4纵向" then
						BaseHeith=148.5
					BaseWidth=105
					strPaperSize="210X297"
						SetPaper  2,1,1,1,21,29.7 
						H=24.9: W=17.6
					elseif PrintPaper = "A3纵向" then
						BaseHeith=210
					BaseWidth=148.5
					strPaperSize="297X420"
						SetPaper  2,1,1,1,29.7,42
						H=37.2: W=26.3
					else
						BaseHeith=148.5
					BaseWidth=210
					strPaperSize="420X297"
					  SetPaper  2,2,1,1,42,29.7
						H=24.9: W=35.2
					end if
					SSProcess.GetObjectPoint idZH,0,x0,y0,z0,ptype0,name0
					SSProcess.GetObjectPoint idZH,1,x1,y1,z1,ptype1,name1
					SSProcess.GetObjectPoint idZH,2,x2,y2,z2,ptype2,name2
					path = SSProcess.GetSysPathName(7)&"Pictures\"
					strBmpFile = path & "RFT"&i&".wmf"
					Scale = SSProcess.GetMapScale
					minX = x0 - 9*Sqr((x0-x1)^2 + (y0-y1)^2)/BaseWidth
					minY = y0 - 32*Sqr((x2-x1)^2 + (y2-y1)^2)/BaseHeith
					maxX = x2 + 9*Sqr((x0-x1)^2 + (y0-y1)^2)/BaseWidth
					maxY = y2 + 32*Sqr((x2-x1)^2 + (y2-y1)^2)/BaseHeith
					dpi = 300
					'TZK = (x2 - x0)* 1000/PrintScale 
					'TZG = (y2 - y0)* 1000/PrintScale 
					'strPaperSize = TZK&"X"&TZG
					SSFunc.DrawToImage minX, minY, maxX, maxY, strPaperSize, dpi, strBmpFile '输出指定范围内的图形到bmp图片
					'wApp.ActiveDocument.Tables(tableIndex).Cell(1,1).Select       '图片
					Set iShape = wApp.Selection.InlineShapes.AddPicture (strBmpFile,False,True) '插入图片
					'wApp.Selection.ParagraphFormat.Alignment =1     '居中
					iShape.Height = 28.345 *H            '设置图片插入后高度
	            iShape.Width = 28.345 *W        '设置图片插入后宽度（锁定纵横比）
					'转换插入图片类型，干掉下划线
					iShape.ConvertToShape
					'deletefiles strBmpFile
					wApp.ActiveDocument.Tables(tableIndex).Cell(1,1).Select  
               wApp.Selection.ParagraphFormat.Alignment =1
End Function

Function 	GetLDBHJMJ(LDFWXIDs,LDBHJMJ,SinLDFWXIDsCount)
	Dim SinLDFWXIDs(1000)
	SSFunc.Scanstring LDFWXIDs,",",SinLDFWXIDs,SinLDFWXIDsCount
	For i = 0 to SinLDFWXIDsCount - 1
			LDBHJMJ(i,0) = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[绿地编号]" )
			LDLX = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[绿地类型]" )
			ZXLX = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[子项类型]" )
			FTHDJSX = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[覆土厚度及属性]" )
			LDMJ = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[绿地面积]" )
			'绿地类型=单块绿地,集中绿地;子项类型=地面绿化,地下室及半地下室顶绿地,地面绿化内园路及园林铺装,地面绿化内景观水体,地下室及半地下室顶园路及园林铺装,地下室及半地下室顶景观水体,屋顶绿化,屋顶绿化园路及园林铺装,园路及园林铺装,景观水体;覆土厚度及属性=大于1.5m,1.0m~1.5m,0.5m~1.0m,小于0.5m,地面绿化,园路及铺装,地面,地下室顶
			If LDLX = "单块绿地" Then
					If ZXLX = "地面绿化" Then
							LDBHJMJ(i,1) = LDMJ
					ElseIf ZXLX = "地下室及半地下室顶绿地" Then
							If FTHDJSX = "大于1.5m" Then
									LDBHJMJ(i,2) = LDMJ
							ElseIf FTHDJSX = "1.0m~1.5m" Then
									LDBHJMJ(i,3) = LDMJ
							ElseIf FTHDJSX = "0.5m~1.0m" Then
									LDBHJMJ(i,4) = LDMJ
							ElseIf FTHDJSX = "小于0.5m" Then
									LDBHJMJ(i,5) = LDMJ
							End If
					ElseIf ZXLX = "地面绿化内园路及园林铺装" Then
							LDBHJMJ(i,6) = LDMJ
					ElseIf ZXLX = "地面绿化内景观水体" Then
							LDBHJMJ(i,7) = LDMJ
					ElseIf ZXLX = "地下室及半地下室顶园路及园林铺装" Then
							LDBHJMJ(i,8) = LDMJ
					ElseIf ZXLX = "地下室及半地下室顶景观水体" Then
							LDBHJMJ(i,9) = LDMJ
					ElseIf ZXLX = "屋顶绿化" Then
							If FTHDJSX = "大于1.5m" Then
									LDBHJMJ(i,10) = LDMJ
							ElseIf FTHDJSX = "1.0m~1.5m" Then
									LDBHJMJ(i,11) = LDMJ
							ElseIf FTHDJSX = "0.5m~1.0m" Then
									LDBHJMJ(i,12) = LDMJ
							ElseIf FTHDJSX = "小于0.5m" Then
									LDBHJMJ(i,13) = LDMJ
							End If
					ElseIf ZXLX = "屋顶绿化园路及园林铺装" Then
							LDBHJMJ(i,14) = LDMJ
					End If
			ElseIf LDLX = "集中绿地" Then
					If ZXLX = "地面绿化" Then
							If FTHDJSX = "地面绿化" Then
									LDBHJMJ(i,15) = LDMJ
							ElseIf FTHDJSX = "园路及铺装" Then
									LDBHJMJ(i,16) = LDMJ
							End If
					ElseIf ZXLX = "地下室及半地下室顶绿地" Then
							If FTHDJSX = "大于1.5m" Then
									LDBHJMJ(i,17) = LDMJ
							ElseIf FTHDJSX = "1.0m~1.5m" Then
									LDBHJMJ(i,18) = LDMJ
							ElseIf FTHDJSX = "0.5m~1.0m" Then
									LDBHJMJ(i,19) = LDMJ
							ElseIf FTHDJSX = "小于0.5m" Then
									LDBHJMJ(i,20) = LDMJ
							ElseIf FTHDJSX = "园路及铺装" Then
									LDBHJMJ(i,21) = LDMJ
							End If
					ElseIf ZXLX = "园路及园林铺装" Then
							If FTHDJSX = "地面" Then
									LDBHJMJ(i,22) = LDMJ
							ElseIf FTHDJSX = "地下室顶" Then
									LDBHJMJ(i,23) = LDMJ
							End If
					ElseIf ZXLX = "景观水体" Then
							If FTHDJSX = "地面" Then
									LDBHJMJ(i,24) = LDMJ
							ElseIf FTHDJSX = "地下室顶" Then
									LDBHJMJ(i,25) = LDMJ
							End If
					End If
			End If
	Next
End Function 



'&&&&&&&&&&==========<>==========&&&&&&&&&&
Function GetGreenMapID(GreenMapID)
		GreenMapID = ""
		sql = "SELECT GH_绿地竣工地形图图廓属性表.ID FROM GH_绿地竣工地形图图廓属性表 INNER JOIN GeoLineTB ON GH_绿地竣工地形图图廓属性表.ID = GeoLineTB.ID WHERE (([GeoLineTB].[Mark] Mod 2<>0));"
		SSProcess.OpenAccessRecordset projectName,sql
		rscount = SSProcess.GetAccessRecordCount(projectName,sql)
		If rscount = 1 Then
				SSProcess.AccessMoveFirst projectName,sql
				While (SSProcess.AccessIsEOF (projectName, sql ) = False)
						SSProcess.GetAccessRecord projectName, sql, fields, values
						GreenMapID = values
						SSProcess.AccessMoveNext projectName, sql
				Wend
		End If
		SSProcess.CloseAccessRecordset projectName,sql
End Function
'&&&&&&&&&&==========<Tbale4 绿地测量成果表>==========&&&&&&&&&&

'&&&&&&&&&&==========<FormatData81 格式化参数>==========&&&&&&&&&&
Function FormatData81(var1,dec)
	FormatData81 = ""
	If var1 <> "" Then 
			If Isnumeric(var1) = True Then
					FormatData81 = Cstr(FormatNumber(Cdbl(var1),dec,-1,0,0))
			End If
	End If
End Function 

'&&&&&&&&&&==========<ExportWord_LD_CheckInfo 检查输出目录结构是否正确>==========&&&&&&&&&&
Function ExportWord_LD_CheckInfo( pathname )  '目录结构初始化
			ExportWord_LD_CheckInfo=false
			Dim fso,arID(100)
			Set fso = CreateObject("Scripting.FileSystemObject") 
			'Epsfilename=SSProcess.GetProjectFileName'获取工程路径
			'SSFunc.ScanString Epsfilename, "\", arID, idCount'通过  \  来分解路径，获取  arID-》分解后的路径  IDCount-》分解后的数目
			'filename=replace(replace(arID(idCount-1),".edb","")," ","")'获取工程名称
		'判断同一目录下是否有同名文件夹
		   if right(filename,1)="\" then filename=left(filename,len(filename)-1)
			'ishavefile  fso,DwgPathName,filename,zongneirong'最外层文件名称   即以工程编号命名的文件名称
			ishavefile  fso,pathname,"绿地测量",zongneirong  '下一层文件名称
			ishavefile  fso,pathname,"绿地测量\原始数据",zongneirong
			ishavefile  fso,pathname,"绿地测量\过程数据",zongneirong
			ishavefile  fso,pathname,"绿地测量\成果数据",zongneirong
			if zongneirong="" then ExportWord_LD_CheckInfo=true
End Function
'&&&&&&&&&&==========<ishavefile 判断路径下文件是否存在>==========&&&&&&&&&&
function ishavefile(fso,PathName,filenames,zongneirong)
	if fso.folderExists(PathName&filenames)=false  then
			 if zongneirong="" then
					zongneirong="【"&filenames&"】 文件夹不存在！"
			  else
					zongneirong=zongneirong&chr(10)&chr(13)&"【"&filenames&"】 文件夹不存在！"
			  end if
	end if
end function
'&&&&&&&&&&==========<DeleteAllFiles 删除指定目录下指定类型文件>==========&&&&&&&&&&
Function DeleteAllFiles(Filespathname,hzhui)
		 if right(Filespathname,1)<>"\" then Filespathname=Filespathname&"\"
		 Set fso = CreateObject("Scripting.FileSystemObject")
		 dim filenames(1000)
	 	 dim filecount
		 filecount = 0
		 GetAllFiles  Filespathname, hzhui,filecount, filenames
		 for i=0 to filecount-1
					Set PDFfile = fso.GetFile(filenames(i))
					PDFfile.Delete
		 next
end function

'&&&&&&&&&&==========<ReplaceOneStr 替换文字内容>==========&&&&&&&&&&
Function ReplaceOneStr(ByVal Field, ByVal Value)
   wApp.Selection.Find.ClearFormatting
   wApp.Selection.Find.Replacement.ClearFormatting
   With wApp.Selection.Find
      .Text =  Field
      .Replacement.Text =  Value
      .Forward = True
      .Wrap = 2
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchByte = True
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
   End With
   wApp.Selection.Find.Execute , , , , , , , , , ,2
End Function 

'&&&&&&&&&&==========<CopyPage 复制一页>==========&&&&&&&&&&
Function CopyPage(page)
   wApp.Selection.Goto 1, , Page 'wNum
   wApp.Selection.Bookmarks("\Page").Range.copy '复制
End Function

'&&&&&&&&&&==========<PastePage 粘贴一页>==========&&&&&&&&&&
Function PastePage(page)
	 wApp.Selection.Goto 1, , Page 'wNum
	 wApp.Selection.Paste'粘贴
End Function

'&&&&&&&&&&==========<WriteCell 写单元格内容>==========&&&&&&&&&&
Function WriteCell(TableIndex,Line,Row,Words)
		wApp.ActiveDocument.Tables(TableIndex).Cell(Line,Row).Select
		wApp.Selection.TypeText cstr(Words)
End Function 

'&&&&&&&&&&==========<CloseWord 关闭Word>==========&&&&&&&&&&
Function CloseWord(WordPathName)
		docWord.SaveAs WordPathName
		docWord.close
		wApp.Quit
End Function 

'&&&&&&&&&&==========<CreatWordbyTemplate 根据指定模板在指定目录创建文件>==========&&&&&&&&&&
Function CreatWordbyTemplate(WordTemplatePath,WordSavePath,WordSaveName)
	Dim Filecount, Filenames(1000)
   Set fso = CreateObject("Scripting.FileSystemObject")
   If fso.fileExists(WordTemplatePath) = False  Then
			Msgbox "未找到成果报告书模板，无法输出！"
			Exit Function
   End If
   Set wApp = CreateObject("Word.Application")
   Set docWord = wApp.Documents.Add(WordTemplatePath)
	GetAllFiles WordSavePath, "docx",Filecount, Filenames
	For i=0 To Filecount-1
			If Replace(Ucase(Filenames(i)),Ucase(WordSavePath&WordSaveName),"") <> Ucase(Filenames(i) ) Then
					WordSaveName = Replace(WordSaveName,".docx","")&"("&Replace(Date,"/","")&" "&Replace(Time,":","：")&")"&".docx"
               Exit For
			End If
	Next
   wApp.ActiveDocument.SaveAs WordSavePath&WordSaveName
   wApp.Application.Visible = True
	wApp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
   wApp.ActiveWindow.View.DisplayPageBoundaries = True
End Function 

Function GetAllFiles(ByRef pathname, ByRef fileExt, ByRef filecount, ByRef filenames())
	Dim fso, folder, file, files, subfolder,folder0, fcount
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(pathname)
	Set files = folder.Files
	'查找文件
	For Each file In files
			Extname = fso.GetExtensionName(file.name)
			If UCase(Extname) = UCase(fileExt) Then
					filenames(filecount) = pathname & file.name
					filecount = filecount+1
			End If
	Next
End Function


'设置纸张打印格式及样式
Function SetPaper(IniTop,IniBottom,IniLeft,IniRight,PageWidth ,PageHeight)
    With wApp.Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = 1
        .TopMargin = 28.345*IniTop
        .BottomMargin =28.345*IniBottom
        .LeftMargin = 28.345*IniLeft
        .RightMargin = 28.345*IniRight
        .Gutter = 0
        .HeaderDistance =28.345*1.5
        .FooterDistance = 28.345*1.75
        .PageWidth =28.345*PageWidth
        .PageHeight = 28.345*PageHeight
        .FirstPageTray = 0
        .OtherPagesTray = 0
        .SectionStart = 2
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = 0
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = 0
        .LayoutMode = 0
    End With
End Function


Function deletefiles(byval path)
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.deleteFile path
      set fso=nothing
End Function

Function CheckFloderExists(flodername)
	Dim fso
	Set fso = CreateObject ("scripting.filesystemobject")
	CheckFloderExists = fso.FolderExists(flodername)
End Function

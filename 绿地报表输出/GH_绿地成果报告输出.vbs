
'*************************************************************************************
'Copyright (c) 2018-2019 Kevin Yang. All Rights Reserved.
'Tel.15357565878           E-mail.1402565009@qq.com
'Origin:Kevin Yang,20181024-02:00
'*************************************************************************************
Dim wApp,docWord,projectName
zongneirong=""
Sub OnClick()
	Dim arArray(100),count
	'&&&&&&&&&&==========<����Word>==========&&&&&&&&&&
	SSProcess.PushUndoMark 
   'pathname=SSProcess.GetSysPathName(5)
	Pathname1=SSProcess.GetSysPathName(5)&"���ϲ��"
	set folder=createobject("scripting.filesystemobject")
	If CheckFloderExists(Pathname1) = False Then 
		set fld=folder.createfolder(Pathname1)
	end if
	pathname2 = Pathname1&"\"&"�̵�"
		If CheckFloderExists(pathname2) = False Then 
		set fld=folder.createfolder(pathname2)
	end if
	pathname=pathname2&"\"
	If pathname = "" Then Exit Sub
	'If ExportWord_LD_CheckInfo(pathname)  =  false then  msgbox zongneirong&"����Ŀ¼�ṹ�����봴��Ŀ¼�ṹ������ȡ�������" : exit sub
  ' msg = msgbox("�Ƿ�ɾ�����Ŀ¼�������ļ���",4,"��ʾ��")
	'If msg = 6 Then
		'	DeleteAllFiles pathname&"\�̵ز���\�ɹ�����", "docx"'ɾ��dwg  
	'Else
			'msgbox "ȡ�����":exit sub
'	End if 
	WordTemplatePath = SSProcess.GetSysPathName(7)&"�������̿����̵ز����ɹ�������.docx"
	'WordSavePath = pathname&"\�̵ز���\�ɹ�����\"
	'WordSavePath = pathname&"\�̵ز���\�ɹ�����\"
	WordSavePath = pathname
	WordSaveName = "�������̿����̵ز����ɹ�������.docx"
	CreatWordbyTemplate WordTemplatePath,WordSavePath,WordSaveName

	'&&&&&&&&&&==========<��Access>==========&&&&&&&&&&
	projectName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb projectName
	'&&&&&&&&&&==========<Tbale4 �̵ز����ɹ���>==========&&&&&&&&&&
	Table4
	'NONEED&&&&&&&&&&==========<Tbale3 ����˵��>==========&&&&&&&&&&
	'Table3
	'PASS&&&&&&&&&&==========<Tbale2 ���������>==========&&&&&&&&&&
	Table2
	'PASS&&&&&&&&&&==========<Tbale1  ����&Replace>==========&&&&&&&&&&
	Table1
  	'&&&&&&&&&&==========<��ӡͼƬ>==========&&&&&&&&&&
    table_5()
	'&&&&&&&&&&==========<�ر�Word>==========&&&&&&&&&&
	CloseWord WordSavePath&WordSaveName
	'&&&&&&&&&&==========<�ر�Access>==========&&&&&&&&&&
   SSProcess.CloseAccessMdb projectName
End Sub

'&&&&&&&&&&==========<Tbale1 ����&Replace>==========&&&&&&&&&&

'// �������


Function Table1()

  GetYDHXinfo HXYDBH,HXYDMJ,HXXMMC,HXXMDZ,HXSJDW,HXWTDW,HXJSDW,HXCHDW,HXCHSJ,HXBZR,HXJCR,HXSHR
	ReplaceOneStr "{GCBH}",HXYDBH'�����õ���Ŀ���
	ReplaceOneStr "{JSDW}",HXJSDW'���赥λ
	ReplaceOneStr "{XMMC}",HXXMMC'��Ŀ����
	ReplaceOneStr "{XMDZ}",HXXMDZ'��Ŀ��ַ
	ReplaceOneStr "{SJDW}",HXSJDW'��Ƶ�λ
	ReplaceOneStr "{WTDW}",HXWTDW'ί�е�λ
	ReplaceOneStr "{HXH}",HXH
	ReplaceOneStr "{CHDW}",HXCHDW '��浥λ
	ReplaceOneStr "{CHSJ}",HXCHSJ'���ʱ��
	ReplaceOneStr "{BZZ}",HXBZR'������
	ReplaceOneStr "{JCZ}",HXJCR'�����
	ReplaceOneStr "{SHZ}",HXSHR'�����
	ReplaceOneStr "{��}",year(NOW)
	ReplaceOneStr "{��}",Month(NOW)
	ReplaceOneStr "{��}",day(NOW)
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
                      HXYDBH=SSProcess.GetSelGeoValue( 0, "[���̱��]")
                      HXYDMJ=SSProcess.GetSelGeoValue( 0, "[�滮�õ����]")
                      HXXMMC=SSProcess.GetSelGeoValue(0, "[��Ŀ����]")
                      HXXMDZ=SSProcess.GetSelGeoValue( 0, "[��Ŀ��ַ]")
                      HXSJDW=SSProcess.GetSelGeoValue(0, "[��Ƶ�λ]")
                      HXWTDW=SSProcess.GetSelGeoValue( 0, "[ί�е�λ]")
                      HXJSDW=SSProcess.GetSelGeoValue( 0, "[���赥λ]")
                      HXCHDW=SSProcess.GetSelGeoValue( 0, "[��浥λ]")
                      HXCHSJ=SSProcess.GetSelGeoValue( 0, "[���ʱ��]")
                      HXBZR=SSProcess.GetSelGeoValue( 0, "[����]")
                      HXJCR=SSProcess.GetSelGeoValue( 0, "[���]")
                      HXSHR=SSProcess.GetSelGeoValue( 0, "[���]")
                      next
               else 
                msgbox "�����д��ڶ����õغ���,��������ݣ�"
               end if 
End Function 

'&&&&&&&&&&==========<Tbale1  ����&Replace>==========&&&&&&&&&&

'&&&&&&&&&&==========<Tbale2 ���������>==========&&&&&&&&&&
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
	sql = "SELECT ����,�ϸ�֤���Ż�ְҵ�ʸ�֤���,��ע FROM ��Ա��Ϣ���Ա� WHERE ([��Ա��Ϣ���Ա�].[��ҵҵ������] = '�̵ز���');"
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
'&&&&&&&&&&==========<Tbale2 ���������>==========&&&&&&&&&&

'&&&&&&&&&&==========<Tbale4 �̵ز����ɹ���>==========&&&&&&&&&&
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
      ALLLDBH(i)=SSProcess.GetObjectAttr( LDvArray(i), "[�̵ر��]")
     next
     end if 
     SSFunc.SortArrayByValue ALLLDIDS,ALLLDBH,LDnCount,1,0 '�ؿ�����С�����������
     bfh="%": ALLZSMJ="":DKDMLDMJ="":DKDXSMJ="":DKWDMJ="":JZDMLDMJ="":JZDXSMJ=""
     for j=0 to LDnCount -1
      DDLDID=SSProcess.GetObjectAttr( ALLLDIDS(j), "SSObj_ID")
      DDLDAREA=SSProcess.GetObjectAttr( ALLLDIDS(j), "SSObj_Area")
      DDLDBH=SSProcess.GetObjectAttr( ALLLDIDS(j), "[�̵ر��]")
      DDLDLB=SSProcess.GetObjectAttr( ALLLDIDS(j), "[�̵����]")
      DDLDXZ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[�̵�����]")
      DDLDMJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[�̵����]")'ȥ������԰�����
      DDLDMJ=Round(cdbl(DDLDMJ),2)
      DDLDZSHMJ=SSProcess.GetObjectAttr( ALLLDIDS(j),"[������̵����]")
      DDLDZSHMJ=Round(cdbl(DDLDZSHMJ),2)
      DDLCXQMJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[�������̵����]")
      DDLCXQMJ=Round(cdbl(DDLCXQMJ),2)
      DDLDFTHD=SSProcess.GetObjectAttr( ALLLDIDS(j), "[�������]")
      DDLDZSXJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[����ϵ��]")
      DDLDJGMJ=SSProcess.GetObjectAttr( ALLLDIDS(j), "[����ˮ����԰·��԰����װ���]")
      DDLDJGMJ=Round(cdbl(DDLDJGMJ),2)
         if DDLDXZ="����" then 
              if DDLDLB="�����̻�" then
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

                elseif DDLDLB="�����Ҽ�������Ҷ��̻�" then 
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

              elseif DDLDLB="�ݶ��̻�" then 
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
          IF DDLDXZ="����" then 
              if DDLDLB="�����̻�" then
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
						 elseif DDLDLB="�����Ҽ�������Ҷ��̻�" then 
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
   if HXYDMJ="" then msgbox "��ȷ���õغ����С��滮�õ�������ֶ��Ƿ���д"
   WriteCell TableIndex,16,2,HXYDMJ'??????????????????????????????  ����Ϣ¼�����Ա���ȡ���õ����
   WriteCell TableIndex,17,2,(round(cdbl(ALLZSMJ)/cdbl(HXYDMJ),4))*100&""&bfh
   else
   WriteCell TableIndex,7+LDnCount,3,DKDMLDMJ
   WriteCell TableIndex,7+LDnCount,5,DKDXSMJ
   WriteCell TableIndex,7+LDnCount,7,DKWDMJ
   WriteCell TableIndex,7+LDnCount,9,JZDMLDMJ
   WriteCell TableIndex,7+LDnCount,11,JZDXSMJ
   WriteCell TableIndex,8+LDnCount,2,ALLZSMJ
   if HXYDMJ="" then msgbox "��ȷ���õغ����С��滮�õ�������ֶ��Ƿ���д"
   WriteCell TableIndex,9+LDnCount,2,HXYDMJ'??????????????????????????????  ����Ϣ¼�����Ա���ȡ���õ����
   WriteCell TableIndex,10+LDnCount,2,(round(cdbl(ALLZSMJ)/cdbl(HXYDMJ),4))*100&""&bfh
   end if 
End Function

'�ۺϿ���ͼ
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

					PrintPaper = SSProcess.GetObjectAttr(idZH,"[��ӡֽ��]")
					PrintScale = SSProcess.GetObjectAttr(idZH,"[��ӡ����]")
					SSProcess.SetMapScale PrintScale 
					wApp.ActiveDocument.Tables(tableIndex).Cell(1,1).Select       'ͼƬ
					if PrintPaper = "A4����" then
						BaseHeith=105
					   BaseWidth=148.5
					   strPaperSize="297X210"
						SetPaper 2,2,1,1,29.7,21
						H=16.2: W=22.9
					elseif PrintPaper = "A4����" then
						BaseHeith=148.5
					BaseWidth=105
					strPaperSize="210X297"
						SetPaper  2,1,1,1,21,29.7 
						H=24.9: W=17.6
					elseif PrintPaper = "A3����" then
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
					SSFunc.DrawToImage minX, minY, maxX, maxY, strPaperSize, dpi, strBmpFile '���ָ����Χ�ڵ�ͼ�ε�bmpͼƬ
					'wApp.ActiveDocument.Tables(tableIndex).Cell(1,1).Select       'ͼƬ
					Set iShape = wApp.Selection.InlineShapes.AddPicture (strBmpFile,False,True) '����ͼƬ
					'wApp.Selection.ParagraphFormat.Alignment =1     '����
					iShape.Height = 28.345 *H            '����ͼƬ�����߶�
	            iShape.Width = 28.345 *W        '����ͼƬ������ȣ������ݺ�ȣ�
					'ת������ͼƬ���ͣ��ɵ��»���
					iShape.ConvertToShape
					'deletefiles strBmpFile
					wApp.ActiveDocument.Tables(tableIndex).Cell(1,1).Select  
               wApp.Selection.ParagraphFormat.Alignment =1
End Function

Function 	GetLDBHJMJ(LDFWXIDs,LDBHJMJ,SinLDFWXIDsCount)
	Dim SinLDFWXIDs(1000)
	SSFunc.Scanstring LDFWXIDs,",",SinLDFWXIDs,SinLDFWXIDsCount
	For i = 0 to SinLDFWXIDsCount - 1
			LDBHJMJ(i,0) = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[�̵ر��]" )
			LDLX = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[�̵�����]" )
			ZXLX = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[��������]" )
			FTHDJSX = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[������ȼ�����]" )
			LDMJ = SSProcess.GetObjectAttr (SinLDFWXIDs(i), "[�̵����]" )
			'�̵�����=�����̵�,�����̵�;��������=�����̻�,�����Ҽ�������Ҷ��̵�,�����̻���԰·��԰����װ,�����̻��ھ���ˮ��,�����Ҽ�������Ҷ�԰·��԰����װ,�����Ҽ�������Ҷ�����ˮ��,�ݶ��̻�,�ݶ��̻�԰·��԰����װ,԰·��԰����װ,����ˮ��;������ȼ�����=����1.5m,1.0m~1.5m,0.5m~1.0m,С��0.5m,�����̻�,԰·����װ,����,�����Ҷ�
			If LDLX = "�����̵�" Then
					If ZXLX = "�����̻�" Then
							LDBHJMJ(i,1) = LDMJ
					ElseIf ZXLX = "�����Ҽ�������Ҷ��̵�" Then
							If FTHDJSX = "����1.5m" Then
									LDBHJMJ(i,2) = LDMJ
							ElseIf FTHDJSX = "1.0m~1.5m" Then
									LDBHJMJ(i,3) = LDMJ
							ElseIf FTHDJSX = "0.5m~1.0m" Then
									LDBHJMJ(i,4) = LDMJ
							ElseIf FTHDJSX = "С��0.5m" Then
									LDBHJMJ(i,5) = LDMJ
							End If
					ElseIf ZXLX = "�����̻���԰·��԰����װ" Then
							LDBHJMJ(i,6) = LDMJ
					ElseIf ZXLX = "�����̻��ھ���ˮ��" Then
							LDBHJMJ(i,7) = LDMJ
					ElseIf ZXLX = "�����Ҽ�������Ҷ�԰·��԰����װ" Then
							LDBHJMJ(i,8) = LDMJ
					ElseIf ZXLX = "�����Ҽ�������Ҷ�����ˮ��" Then
							LDBHJMJ(i,9) = LDMJ
					ElseIf ZXLX = "�ݶ��̻�" Then
							If FTHDJSX = "����1.5m" Then
									LDBHJMJ(i,10) = LDMJ
							ElseIf FTHDJSX = "1.0m~1.5m" Then
									LDBHJMJ(i,11) = LDMJ
							ElseIf FTHDJSX = "0.5m~1.0m" Then
									LDBHJMJ(i,12) = LDMJ
							ElseIf FTHDJSX = "С��0.5m" Then
									LDBHJMJ(i,13) = LDMJ
							End If
					ElseIf ZXLX = "�ݶ��̻�԰·��԰����װ" Then
							LDBHJMJ(i,14) = LDMJ
					End If
			ElseIf LDLX = "�����̵�" Then
					If ZXLX = "�����̻�" Then
							If FTHDJSX = "�����̻�" Then
									LDBHJMJ(i,15) = LDMJ
							ElseIf FTHDJSX = "԰·����װ" Then
									LDBHJMJ(i,16) = LDMJ
							End If
					ElseIf ZXLX = "�����Ҽ�������Ҷ��̵�" Then
							If FTHDJSX = "����1.5m" Then
									LDBHJMJ(i,17) = LDMJ
							ElseIf FTHDJSX = "1.0m~1.5m" Then
									LDBHJMJ(i,18) = LDMJ
							ElseIf FTHDJSX = "0.5m~1.0m" Then
									LDBHJMJ(i,19) = LDMJ
							ElseIf FTHDJSX = "С��0.5m" Then
									LDBHJMJ(i,20) = LDMJ
							ElseIf FTHDJSX = "԰·����װ" Then
									LDBHJMJ(i,21) = LDMJ
							End If
					ElseIf ZXLX = "԰·��԰����װ" Then
							If FTHDJSX = "����" Then
									LDBHJMJ(i,22) = LDMJ
							ElseIf FTHDJSX = "�����Ҷ�" Then
									LDBHJMJ(i,23) = LDMJ
							End If
					ElseIf ZXLX = "����ˮ��" Then
							If FTHDJSX = "����" Then
									LDBHJMJ(i,24) = LDMJ
							ElseIf FTHDJSX = "�����Ҷ�" Then
									LDBHJMJ(i,25) = LDMJ
							End If
					End If
			End If
	Next
End Function 



'&&&&&&&&&&==========<>==========&&&&&&&&&&
Function GetGreenMapID(GreenMapID)
		GreenMapID = ""
		sql = "SELECT GH_�̵ؿ�������ͼͼ�����Ա�.ID FROM GH_�̵ؿ�������ͼͼ�����Ա� INNER JOIN GeoLineTB ON GH_�̵ؿ�������ͼͼ�����Ա�.ID = GeoLineTB.ID WHERE (([GeoLineTB].[Mark] Mod 2<>0));"
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
'&&&&&&&&&&==========<Tbale4 �̵ز����ɹ���>==========&&&&&&&&&&

'&&&&&&&&&&==========<FormatData81 ��ʽ������>==========&&&&&&&&&&
Function FormatData81(var1,dec)
	FormatData81 = ""
	If var1 <> "" Then 
			If Isnumeric(var1) = True Then
					FormatData81 = Cstr(FormatNumber(Cdbl(var1),dec,-1,0,0))
			End If
	End If
End Function 

'&&&&&&&&&&==========<ExportWord_LD_CheckInfo ������Ŀ¼�ṹ�Ƿ���ȷ>==========&&&&&&&&&&
Function ExportWord_LD_CheckInfo( pathname )  'Ŀ¼�ṹ��ʼ��
			ExportWord_LD_CheckInfo=false
			Dim fso,arID(100)
			Set fso = CreateObject("Scripting.FileSystemObject") 
			'Epsfilename=SSProcess.GetProjectFileName'��ȡ����·��
			'SSFunc.ScanString Epsfilename, "\", arID, idCount'ͨ��  \  ���ֽ�·������ȡ  arID-���ֽ���·��  IDCount-���ֽ�����Ŀ
			'filename=replace(replace(arID(idCount-1),".edb","")," ","")'��ȡ��������
		'�ж�ͬһĿ¼���Ƿ���ͬ���ļ���
		   if right(filename,1)="\" then filename=left(filename,len(filename)-1)
			'ishavefile  fso,DwgPathName,filename,zongneirong'������ļ�����   ���Թ��̱���������ļ�����
			ishavefile  fso,pathname,"�̵ز���",zongneirong  '��һ���ļ�����
			ishavefile  fso,pathname,"�̵ز���\ԭʼ����",zongneirong
			ishavefile  fso,pathname,"�̵ز���\��������",zongneirong
			ishavefile  fso,pathname,"�̵ز���\�ɹ�����",zongneirong
			if zongneirong="" then ExportWord_LD_CheckInfo=true
End Function
'&&&&&&&&&&==========<ishavefile �ж�·�����ļ��Ƿ����>==========&&&&&&&&&&
function ishavefile(fso,PathName,filenames,zongneirong)
	if fso.folderExists(PathName&filenames)=false  then
			 if zongneirong="" then
					zongneirong="��"&filenames&"�� �ļ��в����ڣ�"
			  else
					zongneirong=zongneirong&chr(10)&chr(13)&"��"&filenames&"�� �ļ��в����ڣ�"
			  end if
	end if
end function
'&&&&&&&&&&==========<DeleteAllFiles ɾ��ָ��Ŀ¼��ָ�������ļ�>==========&&&&&&&&&&
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

'&&&&&&&&&&==========<ReplaceOneStr �滻��������>==========&&&&&&&&&&
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

'&&&&&&&&&&==========<CopyPage ����һҳ>==========&&&&&&&&&&
Function CopyPage(page)
   wApp.Selection.Goto 1, , Page 'wNum
   wApp.Selection.Bookmarks("\Page").Range.copy '����
End Function

'&&&&&&&&&&==========<PastePage ճ��һҳ>==========&&&&&&&&&&
Function PastePage(page)
	 wApp.Selection.Goto 1, , Page 'wNum
	 wApp.Selection.Paste'ճ��
End Function

'&&&&&&&&&&==========<WriteCell д��Ԫ������>==========&&&&&&&&&&
Function WriteCell(TableIndex,Line,Row,Words)
		wApp.ActiveDocument.Tables(TableIndex).Cell(Line,Row).Select
		wApp.Selection.TypeText cstr(Words)
End Function 

'&&&&&&&&&&==========<CloseWord �ر�Word>==========&&&&&&&&&&
Function CloseWord(WordPathName)
		docWord.SaveAs WordPathName
		docWord.close
		wApp.Quit
End Function 

'&&&&&&&&&&==========<CreatWordbyTemplate ����ָ��ģ����ָ��Ŀ¼�����ļ�>==========&&&&&&&&&&
Function CreatWordbyTemplate(WordTemplatePath,WordSavePath,WordSaveName)
	Dim Filecount, Filenames(1000)
   Set fso = CreateObject("Scripting.FileSystemObject")
   If fso.fileExists(WordTemplatePath) = False  Then
			Msgbox "δ�ҵ��ɹ�������ģ�壬�޷������"
			Exit Function
   End If
   Set wApp = CreateObject("Word.Application")
   Set docWord = wApp.Documents.Add(WordTemplatePath)
	GetAllFiles WordSavePath, "docx",Filecount, Filenames
	For i=0 To Filecount-1
			If Replace(Ucase(Filenames(i)),Ucase(WordSavePath&WordSaveName),"") <> Ucase(Filenames(i) ) Then
					WordSaveName = Replace(WordSaveName,".docx","")&"("&Replace(Date,"/","")&" "&Replace(Time,":","��")&")"&".docx"
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
	'�����ļ�
	For Each file In files
			Extname = fso.GetExtensionName(file.name)
			If UCase(Extname) = UCase(fileExt) Then
					filenames(filecount) = pathname & file.name
					filecount = filecount+1
			End If
	Next
End Function


'����ֽ�Ŵ�ӡ��ʽ����ʽ
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

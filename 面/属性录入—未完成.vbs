Dim MJKID,LCID,arrmjkmc(1000),mjkcount,arrmjksygn(1000),arrmjkmjxs(1000),arrmjkmjjrxs(1000),arrsfjr(1000)
Sub OnInitScript()
	mode = 0 '=0 �޲����Ի��� =1 �в����Ի���
	title="ѡ�������"
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
		sql= "SELECT JG_��������ֲ�ͼ��Ϣ���Ա�.ID FROM JG_��������ֲ�ͼ��Ϣ���Ա� INNER JOIN GeoLineTB ON JG_��������ֲ�ͼ��Ϣ���Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark mod 2<>0 )"
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
	sql = "SELECT JG_��������ֲ�ͼ��Ϣ���Ա�.JianZWMC,GuiHXKZBH FROM JG_��������ֲ�ͼ��Ϣ���Ա� INNER JOIN GeoLineTB ON JG_��������ֲ�ͼ��Ϣ���Ա�.ID = GeoLineTB.ID WHERE [GeoLineTB].[Mark] Mod 2 <>0 "
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
					arrmjkmc(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[MianJKMC]")'���������
					arrmjksygn(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[GongNYT]") '������;
					arrmjkmjxs(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[MianJXS]") '���ϵ��
					arrmjkmjjrxs(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[JiRMJXS]") '�������ϵ��
					arrsfjr(mjkcount)=SSProcess.GetObjectAttr (MJKID, "[ShiFJR]") '�Ƿ�Ƽ���
					mjkcount=mjkcount+1
					MJKCG=SSProcess.GetObjectAttr(MJKID,"[CengG]") 
		  'GetInfo dh,gcbh

			xkzxx= SSProcess.GetObjectAttr (MJKID, "[GuiHXKZBH]")'�滮���֤����
			lzhxx=SSProcess.GetObjectAttr (MJKID, "[JianZWMC]")'����������
			getGNYTXX lzhxx,xkzxx,GNYTXX,GNMCXX
			End If
			mode = 1 '=0 �޲����Ի��� =1 �в����Ի���
			title="�������Ϣ¼��"
			SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title,"DlgWidth" , "240" 
			SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title, "DlgHeight" ,"250" 
			SSProcess.WriteEpsIni "CRunScriptDlgMJK_" & title, "ColumnWidth" ,"110" 
			'SSProcess.ClearInputParameter
			for c=0 to mjkcount-1
				If arrmjksygn(c) <>"" Then 
					SSProcess.AddInputParameter "ʹ�ù�������", arrmjksygn(c),3,"SYS_DROPDOWNLIST,"&GNYTXX, ""

					'SSProcess.AddInputParameter "���������", arrmjkmc(c),0, "סլ,��ҵ,�칫,��ҵ��ͨ�ִ�,����ҽ����������,�Ļ���������,����,����������,������,�������,�����ҳ����,�ܿղ�,���ݶ�,���ݿ�̨�µĽ���,��Χ���ĳ��ۿ�̨,����,����,������̨,�ܿ�����,��Χ���ļܿ�����,�������,����ֿ�,���峵��,������,��س���,͹(Ʈ)��,�������ȣ����ȣ�,����,�Ŷ�,����,��������,��������,¥�ݼ䡢ˮ��䡢���ݻ���,����¥��,���ݾ�,���ﾮ,ͨ����������,�̵�,�ж��ǲɹ⾮,����¥��,��̨,����̨,����,����,վ̨,����վ,�շ�վ,���η�,�豸��,�ܵ���,���Ѳ�,�����ݹ�����,δ���������,����", ""
					SSProcess.AddInputParameter "���������", arrmjkmc(c),0, GNMCXX, ""

					SSProcess.AddInputParameter "���ϵ��" ,arrmjkmjxs(c), 0, "0,1,0.5", ""
					'SSProcess.AddInputParameter "Ȩ������", "",0, "˽��,����", ""
					'SSProcess.AddInputParameter "����Զ��ۼ�", "��",3, "SYS_DROPDOWNLIST,��,��", ""
					SSProcess.AddInputParameter "�Ƿ����", arrsfjr(c),3, "SYS_DROPDOWNLIST,��,��", ""
					SSProcess.AddInputParameter "�������ϵ��", arrmjkmjjrxs(c),0, "1,0.5", ""
				Else
					MJKlx1 = "סլ,��ҵ,�칫,��ҵ��ͨ�ִ�,����ҽ����������,�Ļ���������,����,����������,������,�������,�����ҳ����,�ܿղ�,���ݶ�,���ݿ�̨�µĽ���,��Χ���ĳ��ۿ�̨,����,����,������̨,�ܿ�����,��Χ���ļܿ�����,�������,����ֿ�,���峵��,������,��س���,͹(Ʈ)��,�������ȣ����ȣ�,����,�Ŷ�,����,��������,��������,¥�ݼ䡢ˮ��䡢���ݻ���,����¥��,���ݾ�,���ﾮ,ͨ����������,�̵�,�ж��ǲɹ⾮,����¥��,��̨,����̨,����,����,վ̨,����վ,�շ�վ,���η�,�豸��,�ܵ���,���Ѳ�,�����ݹ�����,δ���������,����"


					MJKGN1=SSProcess.ReadEpsIni("MJKGNXX", "MJKSYGN" ,"")
					MJKGN2=SSProcess.ReadEpsIni("MJKGNXX", "MJKMCZ" ,"")
					MJKGN3=SSProcess.ReadEpsIni("MJKGNXX", "SFJRZ" ,"")
					MJKGN4=SSProcess.ReadEpsIni("MJKGNXX", "MianJXSZ" ,"")
					MJKGN5=SSProcess.ReadEpsIni("MJKGNXX", "JiRMJXSZ" ,"")
'&,,
					SSProcess.AddInputParameter "ʹ�ù�������", MJKGN1,3,"SYS_DROPDOWNLIST,"&GNYTXX, ""
					SSProcess.AddInputParameter "���������", MJKGN2,0,GNMCXX&MJKlx1, ""
					SSProcess.AddInputParameter "���ϵ��" ,MJKGN4, 0, "0,1,0.5", ""
					SSProcess.AddInputParameter "�Ƿ����",MJKGN3,3, "SYS_DROPDOWNLIST,��,��", ""
					SSProcess.AddInputParameter "�������ϵ��", MJKGN5,0, "1,0.5", ""
				End If
			next 
			SSProcess.ShowScriptDlg mode,title
			SSProcess.SetCursorStatus 0
		End If
End Function

dim MJKCG
'����ֵ�����ı�
Function OnPropertyChanged( strName, strValue)
If isnumeric(mjkcg) = false Then mjkcg = 0
		SSProcess.UpdateScriptDlgParameter 1
			if  strName="ʹ�ù�������"  Then
				if strValue <> "����"  then 
					SSProcess.AddInputParameter "���������", arrmjkmc(c),0, "סլ,��ҵ,�칫,��ҵ��ͨ�ִ�,����ҽ����������,�Ļ���������,����,����������,������,�������,�����ҳ����,�ܿղ�,���ݶ�,���ݿ�̨�µĽ���,��Χ���ĳ��ۿ�̨,����,����,������̨,�ܿ�����,��Χ���ļܿ�����,�������,����ֿ�,���峵��,������,��س���,͹(Ʈ)��,�������ȣ����ȣ�,����,�Ŷ�,����,��������,��������,¥�ݼ䡢ˮ��䡢���ݻ���,����¥��,���ݾ�,���ﾮ,ͨ����������,�̵�,�ж��ǲɹ⾮,����¥��,��̨,����̨,����,����,վ̨,����վ,�շ�վ,���η�,�豸��,�ܵ���,���Ѳ�,�����ݹ�����,δ���������,����", ""
				SSProcess.AddInputParameter "�Ƿ����", "��",3, "SYS_DROPDOWNLIST,��,��", ""
				else
					xkzxx= SSProcess.GetObjectAttr (MJKID, "[GuiHXKZBH]")'�滮���֤����
					lzhxx=SSProcess.GetObjectAttr (MJKID, "[JianZWMC]")'����������
					getGNMCXX lzhxx,xkzxx,GNMCZB
					SSProcess.AddInputParameter "���������", arrmjkmc(c),0,GNMCZB, ""
				SSProcess.AddInputParameter "�Ƿ����", "��",3, "SYS_DROPDOWNLIST,��,��", ""
				end if
 
		elseif  strName="�Ƿ����"  Then
				if strValue = "��"  then 
				dyjmjkmc = SSProcess.GetInputParameter ("���������")'��һ�����������
				dyjgnyt = SSProcess.GetInputParameter ("ʹ�ù�������")'��һ�����������
				If dyjgnyt="סլ"   then
					if dyjmjkmc = "����" or dyjmjkmc = "����" then 
					jrmjxs = 1
					else
						IF MJKCG <= 3.6 then 
							jrmjxs = 1
						ELSe
							a = fix((mjkcg-3.6)/2.2)
							jrmjxs = a + 2
						End if 
					end if
				Elseif dyjgnyt = "�����칫"  Then
					if dyjmjkmc = "����������"  then 
							IF MJKCG <= 4.5 then 
								jrmjxs = 1
							ELSe
								a = fix((mjkcg-4.5)/2.2)
								jrmjxs = a + 2
							End if 
					else
						jrmjxs = 1
					end if
				Elseif dyjgnyt = "��ҵ"  Then
					if dyjmjkmc = "����������"  then 
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
					SSProcess.SetInputParameter "�������ϵ��", jrmjxs
				else 
					SSProcess.SetInputParameter "�������ϵ��", "0"
				end if
		elseif  strName="���������"  Then
				If strValue="δ���������"   then
					SSProcess.SetInputParameter "ʹ�ù�������", strValue
					SSProcess.SetInputParameter "���ϵ��", "1"
				ElseIf strValue="���Ѳ�" or strValue="�ܵ���" or strValue="�豸��" or strValue="¥�ݼ䡢ˮ��䡢���ݻ���"  or strValue="�Ŷ�"  or strValue="��س���"  or strValue="������"  or strValue="���峵��"  or  strValue="����ֿ�"  or  strValue="�������"  or strValue="����"  or strValue="����"  or strValue="�ܿղ�"  or strValue="����������"  or   strValue="������" or strValue="�������"  Then
					'SSProcess.SetInputParameter "ʹ�ù�������", "סլ"
					IF MJKCG < 2.2 THEN 
					SSProcess.SetInputParameter "���ϵ��", "0.5"
					ELSE 
					SSProcess.SetInputParameter "���ϵ��", "1"
					end if 
				ElseIf  strValue="���ݶ�"  or strValue="���ݿ�̨�µĽ���"  Then
					'SSProcess.SetInputParameter "ʹ�ù�������", "סլ"
					IF MJKCG < 1.2 THEN 
					SSProcess.SetInputParameter "���ϵ��", "0"
					ELSEif MJKCG < 2.1 THEN 
					SSProcess.SetInputParameter "���ϵ��", "0.5"
					ELSE
					SSProcess.SetInputParameter "���ϵ��", "1"
					END IF
				ElseIf replace(strValue,"�в�","") <>  strValue  Then
					'SSProcess.SetInputParameter "ʹ�ù�������", "סլ"
					IF MJKCG < 2 THEN 
					SSProcess.SetInputParameter "���ϵ��", "0"
					END IF
				ElseIf   strValue="����" or strValue="���η�" or  strValue="���²�" or  strValue="��̨" or strValue="�̵�" or strValue="ͨ����������" or strValue="�ܵ���" or strValue="���ﾮ" or strValue="���ݾ�" or strValue="����¥��" or strValue="�ܿ�����"  Then
					'SSProcess.SetInputParameter "ʹ�ù���", "סլ"
					SSProcess.SetInputParameter "���ϵ��", "1"
				ElseIf strValue="�շ�վ" or strValue="����վ" or strValue="վ̨" or strValue="����" or strValue="����" or strValue="����̨" or strValue="����¥��" or strValue="��������" or strValue="����" or strValue="����" or strValue="�������ȣ����ȣ�" or strValue="�����ҳ����" or strValue="��Χ���ļܿ�����"  Then
					'SSProcess.SetInputParameter "ʹ�ù���", "סլ"
					SSProcess.SetInputParameter "���ϵ��", "0.5"
				ElseIf  strValue="�����ݹ�����"  Then
					'SSProcess.SetInputParameter "ʹ�ù���", "סլ"
					SSProcess.SetInputParameter "���ϵ��", "0"
				ElseIf   strValue="��������" or strValue="͹��Ʈ����"  Then
					IF MJKCG >= 2.1 THEN 
					'SSProcess.SetInputParameter "ʹ�ù���", "סլ"
					SSProcess.SetInputParameter "���ϵ��", "0.5"
					else
					SSProcess.SetInputParameter "���ϵ��", "0"
				End If
				ElseIf   strValue="�ж��ǲɹ⾮"   Then
					IF MJKCG >= 2.1 THEN 
					'SSProcess.SetInputParameter "ʹ�ù���", "סլ"
					SSProcess.SetInputParameter "���ϵ��", "1"
					else
					SSProcess.SetInputParameter "���ϵ��", "0.5"
					End If
				ELSE
					SSProcess.SetInputParameter "���ϵ��", "1"
					'SSProcess.SetInputParameter "�������ϵ��", "1"
				End If

				SYGNSXZ = SSProcess.GetInputParameter ("ʹ�ù�������")'��һ��ʹ�ù���
				If SYGNSXZ="סլ"   then
				if MJKCG="" then MJKCG=1
					if strValue = "����" or strValue = "����" then 
						jrmjxs = 1
					else
						IF MJKCG <= 3.6 then 
							jrmjxs = 1
						ELSe
							a = fix((mjkcg-3.6)/2.2)
							jrmjxs = a + 2
						End if 
					end if
				Elseif SYGNSXZ = "�����칫"  Then
					if strValue = "����������"  then 
							IF MJKCG <= 4.5 then 
								jrmjxs = 1
							ELSe
								a = fix((mjkcg-4.5)/2.2)
								jrmjxs = a + 2
							End if 
					else
						jrmjxs = 1
					end if
				Elseif SYGNSXZ = "��ҵ"  Then
					if strValue = "����������"  then 
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
					SSProcess.SetInputParameter "�������ϵ��", jrmjxs
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
				'dyxh=SSProcess.GetInputParameter ("��Ԫ���")
            If dyxh<>"" Then
                  If isnumeric(dyxh)=false Then msgbox "��Ԫ���ӦΪ��ֵXX"  :  Exit Function
            End If
				mjkmc=SSProcess.GetInputParameter ("���������")
				sygn=SSProcess.GetInputParameter ("ʹ�ù�������")
				mjxs=SSProcess.GetInputParameter ("���ϵ��")
            JR=SSProcess.GetInputParameter ("�Ƿ����")
            JRMJXSZ=SSProcess.GetInputParameter ("�������ϵ��")
				if mjxs ="" then mjxs=0
				if JRMJXSZ ="" then JRMJXSZ=0
				if JR="��" then JRMJXSZ=0
				if JR="" then JRMJXSZ=0
				JZMJ = cdbl(JHMJ) * cdbl(mjxs)
				JRMJ = cdbl(JZMJ) * cdbl(JRMJXSZ)
				JZMJ = FormatNumber(JZMJ, 2) 
				JRMJ = FormatNumber(JRMJ, 2) 
				if JR="��" then
					BJRMJ = JZMJ
				else
					BJRMJ = 0
				end if 
				if  SYGN = "סլ"  Then col = RGB(255,0,0)
				if  SYGN = "��ҵ��ͨ�ִ�"  Then col = RGB(255,255,0)
				if  SYGN = "��ҵ"  Then col = RGB(0,255,0)
				if  SYGN = "����ҽ����������"  Then col = RGB(0,255,255)
				if  SYGN = "�Ļ���������"  Then col = RGB(0,0,255)
				if  SYGN = "�칫"  Then col = RGB(255,0,255)
				if  SYGN = "����"  Then col = RGB(128,128,128)
				if  SYGN = "δ���������"  Then col = RGB(255,255,255)
				if  SYGN = "����"  Then col = RGB(192,192,192)
				'qsxz=SSProcess.GetInputParameter ("Ȩ������")
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
            'LJ=SSProcess.GetInputParameter ("����Զ��ۼ�")
				'SSProcess.AddNewObjToSaveObjList


            'If LJ<>"��" Then Exit Function
           ' XH=LJXH(dyxh)
				'mode = 1 '=0 �޲����Ի��� =1 �в����Ի���
				'title="�������Ϣ¼��"
				'SSProcess.ClearInputParameter
				'SSProcess.AddInputParameter "���������", "",0, "סլ,��ҵ,�칫,��ҵ��ͨ�ִ�,����ҽ����������,�Ļ���������,����,����������,������,�������,�����ҳ����,�ܿղ�,���ݶ�,���ݿ�̨�µĽ���,��Χ���ĳ��ۿ�̨,����,����,������̨,�ܿ�����,��Χ���ļܿ�����,�������,����ֿ�,���峵��,������,��س���,͹(Ʈ)��,�������ȣ����ȣ�,����,�Ŷ�,����,��������,��������,¥�ݼ䡢ˮ��䡢���ݻ���,����¥��,���ݾ�,���ﾮ,ͨ����������,�̵�,�ж��ǲɹ⾮,����¥��,��̨,����̨,����,����,վ̨,����վ,�շ�վ,���η�,�豸��,�ܵ���,���Ѳ�,�����ݹ�����,δ���������,����", ""
				'SSProcess.AddInputParameter "ʹ�ù�������", sygn,0, "סլ,��ҵ��ͨ�ִ�,��ҵ,����ҽ����������,�Ļ���������,�칫,����,δ���������,����", ""
				'SSProcess.AddInputParameter "���������", mjkmc,0, "סլ,��ҵ,�칫,��ҵ��ͨ�ִ�,����ҽ����������,�Ļ���������,����,������,�������,�����ҳ����,�ܿղ�,���ݶ�,���ݿ�̨,��Χ���ĳ��ۿ�̨,����,����,������̨,�ܿ�����,��Χ���ļܿ�����,�������,����ֿ�,���峵��,������,��س���,͹(Ʈ)��,�������ȣ����ȣ�,����,�Ŷ�,����,��������,��������,¥�ݼ䡢ˮ��䡢���ݻ���,����¥�ݡ����ݾ������ﾮ��ͨ�������������̵�,�ж��ǲɹ⾮,����¥��,��̨,����̨,������վ̨������վ���շ�վ,���η�,�豸�㡢�ܵ��㡢���Ѳ�,�����ݹ�����,δ���������,����", ""
				'SSProcess.AddInputParameter "���������", "",0, "סլ,��ҵ,�칫,��ҵ��ͨ�ִ�,����ҽ����������,�Ļ���������,����,������,�������,�����ҳ����,�ܿղ�,���ݶ�,���ݿ�̨,��Χ���ĳ��ۿ�̨,����,����,������̨,�ܿ�����,��Χ���ļܿ�����,�������,����ֿ�,���峵��,������,��س���,͹(Ʈ)��,�������ȣ����ȣ�,����,�Ŷ�,����,��������,��������,¥�ݼ䡢ˮ��䡢���ݻ���,����¥�ݡ����ݾ������ﾮ��ͨ�������������̵�,�ж��ǲɹ⾮,����¥��,��̨,����̨,������վ̨������վ���շ�վ,���η�,�豸�㡢�ܵ��㡢���Ѳ�,�����ݹ�����,δ���������,����", ""	
				'SSProcess.AddInputParameter "ʹ�ù�������", sygn,0, "סլ,��ҵ��ͨ�ִ�,��ҵ,����ҽ����������,�Ļ���������,�칫,����,����", ""
				'SSProcess.AddInputParameter "���ϵ��", mjxs,0, "", ""
			'	SSProcess.AddInputParameter "���ϵ��" ,mjxs, 3, "SYS_DROPDOWNLIST", ""
				'SSProcess.AddInputParameter "Ȩ������", qsxz,0, "˽��,����", ""
				'SSProcess.AddInputParameter "����Զ��ۼ�", LJ,0, "��,��", ""
				'SSProcess.AddInputParameter "�Ƿ����", JR,0, "��,��", ""
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
'��ȡ������;��Ϣ
Function getGNYTXX(DH,XKZBH,GNYTXX,GNMCXX)
	Dim Fvalues(2)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName
	sql = "SELECT JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.GongNLX,GongNMC FROM JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա� WHERE (JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.[JianZWMC]) = '"&DH&"' AND (JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.[GuiHXKZBH]) = '"&XKZBH&"';"
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

'��ȡ����������Ϣ
Function getGNMCXX(DH,XKZBH,GNMCZB)
	Dim Fvalues(2)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName
	sql = "SELECT JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.GongNMC FROM JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա� WHERE (JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.[JianZWMC]) = '"&DH&"' AND (JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.[GuiHXKZBH]) = '"&XKZBH&"' And (JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�.[GongNLX]) = '����';"
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




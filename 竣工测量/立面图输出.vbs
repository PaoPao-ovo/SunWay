'�������ͼ�Ի���
Sub OnInitScript()
	mode = 1 '=0 �޲����Ի��� =1 �в����Ի���
	title="�������ͼ....."
	SSProcess.ClearInputParameter 
	SSProcess.AddInputParameter "����ͼ����", "ƽ������",  0, "ƽ������,б������,��ͬ�߶�����", ""
	SSProcess.AddInputParameter "����λ��", "����",  3,"����,����,����,����", ""
	SSProcess.UpdateScriptDlgParameter 1  '���½ű����жԻ������(�����µ��ڴ�)
	SSProcess.ShowScriptDlg mode,title
	SSProcess.WriteEpsIni "CRunScriptDlg2_" & LMtitle,"DlgWidth" , "250" 
	SSProcess.WriteEpsIni "CRunScriptDlg2_" & LMtitle, "DlgHeight" ,"200" 
	SSProcess.WriteEpsIni "CRunScriptDlg2_" & LMtitle, "ColumnWidth" ,"80" 
	SSProcess.RefreshView 
	'LMtitle �� WriteEpsIni
End Sub
Sub OnExitScript()
	'��Ӵ���
End Sub


Sub OnOK()
		SSProcess.UpdateScriptDlgParameter 1
		LMLX = SSProcess.GetInputParameter ("����ͼ����" )
		LMfx = SSProcess.GetInputParameter ("����λ��" )
		SSProcess.UpdateSysSelection 0 'ϵͳѡ�����ݸ��µ��ű�ѡ��
		geoCount = SSProcess.GetSelGeoCount()
		if geoCount<>1 then msgbox "��ѡ�񿢹���������ͼͼ����":exit sub
			GeoCode= SSProcess.GetSelGeoValue( 0, "SSObj_Code" )
		if GeoCode<>"9400604"  then  msgbox "��ѡ�񿢹���������ͼͼ����":exit sub
		If geoCount=1 then 
			LMTTKID = SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
			JZWMC =SSProcess.GetObjectAttr( LMTTKID, "[JianZWMC]") '����
			JZGHXKZH =SSProcess.GetObjectAttr( LMTTKID, "[GuiHXKZBH]")'���
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
		'б��߶�
		ZGzgdgd=CDBL(zgdgc)-CDBL(zflbg)-CDBL(jzwzgd)
		
		If  LMLX = "ƽ������" Then 
			if wjgd <>"" then wjgd=0
			if jzwjbgd <> 0 then jzwjbgd=0
			if neqgc = ""  then neqgc=0
			if jzwjbgd = ""  then jzwjbgd =0
			If neqgd=0 then
				If zgdgc <> 0 and neqgc <> 0 Then
					neqgd = cdbl(zgdgc) - cdbl(neqgc) 
				End If
			End If
			'�ա�ID�������������Ҹ߶ȡ��ܸ߶ȡ�Ů��ǽ�߶ȡ�wjgd���ֲ��߶ȡ���������
			Createlimiantu scsndpg,LMTTKID,geoid,dxsgd,jzwzgd,neqgd,wjgd,jzwjbgd,zflbg

		Elseif  LMLX = "б������" Then 
			if jzwjbgd <> 0 then jzwjbgd=0
			if neqgd <> 0 then neqgd=0
			If wjgd=0 then
				wjgd = cdbl(zgdgc) - cdbl(zflbg)  - cdbl(jzwzgd) 
				wjgd=formatnumber(wjgd,2,-1,0,0) '��ʽ��Ϊ��λС��������
			End If
			Createlimiantu scsndpg,LMTTKID,geoid,dxsgd,jzwzgd,neqgd,wjgd,jzwjbgd,zflbg
		Elseif  LMLX = "��ͬ�߶�����" Then 
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
		TMDXSGD=Dxsgd*blxs  'ͼ������Ҹ߶�
		TMZGD=Zgd*blxs  'ͼ���ܸ߶�
		TMNEQGD=Neqgd*blxs  'ͼ��Ů��ǽ�߶�
		TMZLG=Lzg*blxs 'ͼ����¥��
		TMWJG=Wjgd*blxs 'ͼ���ݼ��߶�
		TMDEZTGD=Deztgd*blxs 'ͼ��ڶ�����߶�
		TMZTGDC=Ztgdc*blxs 'ͼ������߶Ȳ�
		If Deztgd<> 0 then
	'****************���Ƶ�������*********************
			If Dxsgd <> 0 then
				addline jdX0,jdY0,jdX1, jdY0,"9400506",geoid'���Ƶ����ҵ�һ����
				addline jdX0, jdY0+TMDXSGD,jdX0,jdY0,"9400506",geoid'����������
				addline jdX1,jdY0,jdX1, jdY0+TMDXSGD,"9400507",geoid'����������
				addnote "�����ҵװ�",jdX0-5.5,jdY0-1,250*notaise,250*notaise,"����",geoid

	'��߱�ע
				addline jdX0-2.5,jdY0,jdX0-2.5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע��
				addline jdX0,jdY0,jdX0-5, jdY0,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addline jdX0,jdY0+TMDXSGD,jdX0-5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addPoint "9400504",jdX0-2.5,jdY0+((cdbl(TMDXSGD))/2),"",Dxsgd,geoid
				addline jdX0-2.5-1,jdY0-0.5,jdX0-2.5+1, jdY0+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
				addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����

	'�ұ߱�ע
				addline jdX1+2.5,jdY0,jdX1+2.5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע��
				addline jdX1,jdY0,jdX1+5,jdY0,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addline jdX1,jdY0+TMDXSGD,jdX1-5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addPoint "9400504",jdX1+5,jdY0+((cdbl(TMDXSGD))/2),"�����ܸ߶�",Dxsgd,geoid
				addline jdX1+2.5+1,jdY0-0.5,jdX1+2.5-1, jdY0+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
				addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����

			End If
	'****************�����ܸ߶���*********************

			If Dxsgd<> 0 then 
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"940050602",geoid'���ơ�0�����
			Else
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"9400506",geoid'���ơ�0�����
			End If
			addPoint "9400502",jdX0-5,jdY0+TMDXSGD,zflbg, "",geoid '�������
			addPoint "9400502",jdX1+5,jdY0+TMDXSGD, zflbg,"",geoid'�������


			addline jdX0,jdY0,jdX0, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'�����ܸ߶�����
			addline jdX1,jdY0,jdX1, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܸ߶�����


			addline jdX0,jdY0+TMDXSGD+TMDEZTGD,jdX0+2/5*kuandu, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'���Ƶڶ��߶Ⱥ���
			addline jdX0+2/5*kuandu,jdY0+TMDXSGD+TMDEZTGD,jdX0+2/5*kuandu, jdY0+TMDXSGD+TMZGD,"940050602",geoid'���Ƹ߶���ֱ��
			addline jdX0+2/5*kuandu,jdY0+TMDXSGD+TMZGD,jdX1,jdY0+TMDXSGD+TMZGD,"940050602",geoid'���Ƹ߶���ֱ��


			addline jdX0-2.5,jdY0+TMDXSGD,jdX0-2.5, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'��������ܸ߶Ⱦ����ע��
			addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMDEZTGD))/2),"",Deztgd,geoid
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))-0.5,jdX0+1-2.5, jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
			addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0+1-2.5, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����
			addline jdX0-5,jdY0+TMDXSGD,jdX0, jdY0+TMDXSGD,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����


			addline jdX0-2.5,jdY0+TMDXSGD+TMDEZTGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�������߶Ȳ�����ע��
			addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMDEZTGD)+(cdbl(TMZTGDC))/2),"",Ztgdc,geoid
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))-0.5,jdX0-2.5+1, jdY0+cdbl(TMDXSGD)+(cdbl(TMDEZTGD))+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX0-2.5+1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����

			addline jdX0,jdY0+TMDXSGD+TMZGD,jdX0-5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'���Ƶ�һ�߶��ܿ���
			addnote "¥��",jdX0-5.5,jdY0+TMDXSGD+TMZGD+1,250*notaise,250*notaise,"����",geoid
			addline jdX0,jdY0+TMDXSGD+TMDEZTGD,jdX0-5, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'���Ƶڶ��߶��ܿ���
			addnote "¥��",jdX0-5.5,jdY0+TMDXSGD+TMDEZTGD+1,250*notaise,250*notaise,"����",geoid


			addline jdX1+2.5,jdY0+TMDXSGD,jdX1+2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܸ߶Ⱦ����ע��
			addline jdX1,jdY0+TMDXSGD+TMZGD,jdX1+5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'����Ů��ǽ�߶�����
			addPoint "9400504",jdX1+2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMZGD))/2),Zgd, "�����ܸ߶�",geoid
			addline jdX1+2.5+1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX1+2.5-1, jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
			addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����



	'****************����Ů��ǽ*********************
			If Neqgd<> 0 then 
				addline jdX0+2/5*kuandu,jdY0+TMDXSGD+TMZGD,jdX0+2/5*kuandu, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����
				addline jdX1,jdY0+TMDXSGD+TMZGD,jdX1,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����

				addline jdX0-2.5,jdY0+TMDXSGD+TMZGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶ȱ�ע��
				addline jdX0, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����
				addline jdX0-2.5-1, jdY0+TMDXSGD+TMZGD+TMNEQGD-0.5,jdX0-2.5+1,  jdY0+TMDXSGD+TMZGD+TMNEQGD+0.5,"940050602",geoid'����Ů��ǽ�߶ȱ�ע����
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+0.5,"940050602",geoid'����Ů��ǽ�߶ȱ�ע����
				addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMDEZTGD)+cdbl(TMZTGDC)+(cdbl(TMNEQGD)/2)),"",Neqgd,geoid

				addline jdX0, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ��߶�������
				addnote "Ů��ǽ��",jdX0-5.5, jdY0+TMDXSGD+TMZGD+TMNEQGD+1,250*notaise,250*notaise,"����",geoid



				addline jdX0+2/5*kuandu, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0+2/5*kuandu+5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶����ߺ���
				addline jdX1, jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX1-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶����ߺ���
				addline jdX0+2/5*kuandu+5,jdY0+TMDXSGD+TMZGD,jdX0+2/5*kuandu+5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶���������
				addline jdX1-5,jdY0+TMDXSGD+TMZGD,jdX1-5,  jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶���������




			'���Ƶڶ��߶�Ů��ǽ

				addline jdX0,jdY0+TMDXSGD+TMDEZTGD,jdX0, jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����
				addline jdX0,jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,jdX0+5, jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶����ߺ���
				addline jdX0+5,jdY0+TMDXSGD+TMDEZTGD+TMNEQGD,jdX0+5, jdY0+TMDXSGD+TMDEZTGD,"940050602",geoid'����Ů��ǽ�߶���������
			End If

		Else
	'****************���Ƶ�������*********************
			If Dxsgd <> 0 then
				addline jdX0,jdY0,jdX1, jdY0,"9400506",geoid'���Ƶ����ҵ�һ����
				addline jdX0, jdY0+TMDXSGD,jdX0,jdY0,"9400506",geoid'����������
				addline jdX1,jdY0,jdX1, jdY0+TMDXSGD,"9400507",geoid'����������
				addnote "�����ҵװ�",jdX0-5.5,jdY0-1,250*notaise,250*notaise,"����",geoid

	'��߱�ע
				addline jdX0-2.5,jdY0,jdX0-2.5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע��
				addline jdX0,jdY0,jdX0-5, jdY0,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addline jdX0,jdY0+TMDXSGD,jdX0-5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addPoint "9400504",jdX0-2.5,jdY0+((cdbl(TMDXSGD))/2),"",Dxsgd,geoid
				addline jdX0-2.5-1,jdY0-0.5,jdX0-2.5+1, jdY0+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
				addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����

	'�ұ߱�ע
				addline jdX1+2.5,jdY0,jdX1+2.5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע��
				addline jdX1,jdY0,jdX1+5,jdY0,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addline jdX1,jdY0+TMDXSGD,jdX1-5, jdY0+TMDXSGD,"940050602",geoid'���Ƶ��¸߶Ⱦ����ע����
				addPoint "9400504",jdX1+5,jdY0+((cdbl(TMDXSGD))/2),"�����ܸ߶�",Dxsgd,geoid
				addline jdX1+2.5+1,jdY0-0.5,jdX1+2.5-1, jdY0+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
				addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����

			End If
	'****************�����ܸ߶���*********************

			If Dxsgd<> 0 then 
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"940050602",geoid'���ơ�0�����
			Else
				addline jdX0-5,jdY0+TMDXSGD,jdX1+5, jdY0+TMDXSGD,"9400506",geoid'���ơ�0�����
			End If
			addPoint "9400502",jdX0-5,jdY0+TMDXSGD, zflbg,"",geoid '�������
			addPoint "9400502",jdX1+5,jdY0+TMDXSGD, zflbg, "",geoid'�������


			addline jdX0,jdY0,jdX0, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܸ߶�����
			addline jdX1,jdY0,jdX1, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܸ߶�����
			addline jdX0-5,jdY0+TMDXSGD+TMZGD,jdX1+5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܿ���
			addnote "�ܶ�",jdX0-5.5,jdY0+TMDXSGD+TMZGD-1,250*notaise,250*notaise,"����",geoid

 '******************�����ܸ߶��ұ�ע*********************
			addline jdX1+2.5,jdY0+TMDXSGD,jdX1+2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܸ߶Ⱦ����ע��
			addPoint "9400504",jdX1+2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMZGD))/2),Zgd, "�����ܸ߶�",geoid
			addline jdX1+2.5+1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX1+2.5-1, jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
			addline jdX1+2.5+1,jdY0+TMDXSGD-0.5,jdX1+2.5-1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����

 '******************�����ܸ߶����ע*********************
			addline jdX0-2.5,jdY0+TMDXSGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD,"940050602",geoid'�����ܸ߶Ⱦ����ע��
			addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+(cdbl(TMZGD))/2), "",Zgd,geoid
			addline jdX0-2.5-1,jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))-0.5,jdX0-2.5+1, jdY0+cdbl(TMDXSGD)+(cdbl(TMZGD))+0.5,"940050602",geoid'�����ܸ߶ȱ�ע����
			addline jdX0-2.5-1,jdY0+TMDXSGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+0.5,"940050602",geoid'�����ܸ߶ȸ߶ȱ�ע����




	'****************����Ů��ǽ*********************
			If Neqgd<> 0 then 
				addline jdX0,jdY0+TMDXSGD+TMZGD,jdX0, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����
				addline jdX1,jdY0+TMDXSGD+TMZGD,jdX1, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����
				addline jdX0-2.5,jdY0+TMDXSGD+TMZGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶ȱ�ע��
				addline jdX0,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶�����
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD+TMNEQGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+TMNEQGD+0.5,"940050602",geoid'����Ů��ǽ�߶ȱ�ע����
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+0.5,"940050602",geoid'����Ů��ǽ�߶ȱ�ע����
				addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMZGD)+(cdbl(TMNEQGD))/2), "",Neqgd,geoid
				addline jdX0,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ��߶�������
				addnote "Ů��ǽ��",jdX0-5.5,jdY0+TMDXSGD+TMZGD+TMNEQGD+1,250*notaise,250*notaise,"����",geoid

				addline jdX0,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX0+5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶����ߺ���
				addline jdX1,jdY0+TMDXSGD+TMZGD+TMNEQGD,jdX1-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶����ߺ���

				addline jdX0+5,jdY0+TMDXSGD+TMZGD,jdX0+5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶���������
				addline jdX1-5,jdY0+TMDXSGD+TMZGD,jdX1-5, jdY0+TMDXSGD+TMZGD+TMNEQGD,"940050602",geoid'����Ů��ǽ�߶���������
			End If

			If Wjgd<> 0 then
				WJZXDXZB=jdX0+CDBL(kuandu/2) '�ݼ����ĵ�X����
				addline jdX0,jdY0+TMDXSGD+TMZGD,WJZXDXZB, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'�����ݼ��߶�����
				addline jdX1,jdY0+TMDXSGD+TMZGD,WJZXDXZB, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'�����ݼ��߶�����		
				
				addline jdX0,jdY0+TMDXSGD+TMZGD+TMWJG,jdX0-5, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'�����ݼ��߶ȱ�ע����
				
				addline jdX0-2.5,jdY0+TMDXSGD+TMZGD,jdX0-2.5, jdY0+TMDXSGD+TMZGD+TMWJG,"940050602",geoid'�����ݼ��߶ȱ�ע��
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD+TMWJG-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+TMWJG+0.5,"940050602",geoid'�����ݼ�ǽ�߶ȱ�ע����
				addline jdX0-2.5-1,jdY0+TMDXSGD+TMZGD-0.5,jdX0-2.5+1, jdY0+TMDXSGD+TMZGD+0.5,"940050602",geoid'�����ݼ�ǽ�߶ȱ�ע����
				addPoint "9400504",jdX0-2.5,jdY0+(cdbl(TMDXSGD)+cdbl(TMZGD)+(cdbl(TMWJG))/2),"",Wjgd,geoid

				addnote "�ݼ�",jdX0-5.5,jdY0+TMDXSGD+TMZGD+TMWJG+1,250*notaise,250*notaise,"����",geoid
			End If
		End If

end function 




function addline(x00,y00,x11,y11,code,geoid)
		if code="" then code="1"
		SSProcess.CreateNewObj 1
		SSProcess.SetNewObjValue "SSObj_Code", code
		SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
		SSProcess.SetNewObjValue "SSObj_DataMark", "������������ͼ��Ϣ"
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
		SSProcess.SetNewObjValue "SSObj_DataMark", "������������ͼ��Ϣ"
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
		SSProcess.SetNewObjValue "SSObj_DataMark", "������������ͼ��Ϣ"
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
		SSProcess.SetNewObjValue "SSObj_DataMark", "������������ͼ��Ϣ"
		SSProcess.SetNewObjValue "SSObj_GroupID ", GeoID
		SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
		SSProcess.SetNewObjValue "SSObj_FontName", zt
		SSProcess.SetNewObjValue "SSObj_FontWidth", zk
		SSProcess.SetNewObjValue "SSObj_FontHeight", zg
		SSProcess.AddNewObjPoint xx,yy, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
end function 

'��ȡ��������Ϣ
Function GetJZWFWXXX(ghxkzbh,jzwmc,jzwzgd,jzwjbgd,neqgd,neqgc,zgdgc,zflbg,dxsgd,dxsdbgc)
	dim Fvalues(10)
	projectName = SSProcess.GetProjectFileName  
	SSProcess.OpenAccessMdb projectName 
	sql = "SELECT JG_�½������ﷶΧ����Ϣ���Ա�.JianZWZGD,JianZWJBZGD,NvEQGD,NvEQGC,JianZDBZGCGC,ZhengFLBG,DiXGD,DiXSDBGC FROM JG_�½������ﷶΧ����Ϣ���Ա� INNER JOIN GeoAreaTB ON JG_�½������ﷶΧ����Ϣ���Ա�.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 and ([JG_�½������ﷶΧ����Ϣ���Ա�].[ID] > 0 And ([JG_�½������ﷶΧ����Ϣ���Ա�].[JianZWMC] = '"&jzwmc&"') And ([JG_�½������ﷶΧ����Ϣ���Ա�].[GuiHXKZBH] = '"&ghxkzbh&"'));"
	SSProcess.OpenAccessRecordset projectName, sql 
	rscount = SSProcess.GetAccessRecordCount (projectName, sql )
	If rscount >  0 Then
		SSProcess.AccessMoveFirst projectName, sql
		while (SSProcess.AccessIsEOF (projectName, sql ) = False)
			SSProcess.GetAccessRecord projectName, sql, fields, values
			SSFunc.Scanstring values,",",Fvalues,Fvaluescount		
			jzwzgd=Fvalues(0) '�ܸ߶�
			jzwjbgd=Fvalues(1) '�ֲ��߶�
			neqgd=Fvalues(2) 'Ů��ǽ�߶�
			neqgc=Fvalues(3) 'Ů��ǽ�߳�
			zgdgc=Fvalues(4) '��ߵ�߳�
			zflbg=Fvalues(5) '��������
			dxsgd=Fvalues(6) '�����Ҹ߶�
			dxsdbgc=Fvalues(7) '�����ҵװ�߳�
			SSProcess.AccessMoveNext projectName, sql 
		Wend
	End If
	SSProcess.CloseAccessRecordset projectName, sql 
	SSProcess.CloseAccessMdb projectName 
End Function

'��ֵ���㴦��
Function SZZLCL (SJvalue)
		If  SJvalue="" and SJvalue="NULL"   and SJvalue="*"  then SJvalue=0
		If Isnumeric(SJvalue) = False Then SJvalue = 0
End Function





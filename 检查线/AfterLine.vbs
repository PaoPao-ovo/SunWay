'���ݽṹ
Const FWJG_HX = "��,ש,��,��,ʯ,ľ,ձ,��,��,��,����,����,����"

'���ݲ���
Const FWCS_HX = "2,3,4,5,6,7,8,9,10,11,12,13"


Const TDYT = "011:ˮ��,012:ˮ����,013:����,021:��԰,022:��԰,023:����԰��,031:���ֵ�,032:��ľ�ֵ�,033:�����ֵ�,041:��Ȼ���ݵ�,042:�˹����ݵ�,043:�����ݵ�,051:���������õ�,052:ס�޲����õ�,053:��������õ�,054:�����̷��õ�,0601:��ҵ�õ�,0602:�ɿ��õ�,0604:�ִ��õ�,0701:����סլ�õ�,0702:ũ��լ����,0801:���������õ�,0802:���ų����õ�,0803:�ƽ��õ�,0804:ҽ�������õ�,0805:���������õ�,0806:������ʩ�õ�,0807:��԰���̵�,0808:�羰��ʤ��ʩ�õ�,091:������ʩ�õ�,092:ʹ����õ�,093:��̳����õ�,094:�ڽ��õ�,095:�����õ�,101:��·�õ�,102:��·�õ�,103:�����õ�,104:ũ���·,105:�����õ�,106:�ۿ���ͷ�õ�,107:�ܵ������õ�,111:����ˮ��,112:����ˮ��,113:ˮ��ˮ��,114:����ˮ��,115:�غ�̲Ϳ,116:��½̲Ϳ,117:����,118:ˮ�������õ�,119:���������û�ѩ,121:���е�,122:��ʩũ�õ�,123:�￲,124:�μ��,125:�����,126:ɳ��,127:���"

QuanShuLeiXing = "A:������������Ȩ�ڵ�,B:�����õ�ʹ��Ȩ�ڵأ��ر�,S:�����õ�ʹ��Ȩ�ڵأ����ϣ�,X:�����õ�ʹ��Ȩ�ڵأ����£�,C:լ����ʹ��Ȩ�ڵ�,D:���سа���ӪȨ�ڵأ����أ�,E:���سа���ӪȨ�ڵأ��ֵأ�,F:���سа���ӪȨ�ڵأ��ݵأ�,H:����ʹ��Ȩ�ں�,G:�޾��񺣵�ʹ��Ȩ,W:ʹ��Ȩδȷ��������������ػ��򺣵�,Y:����ʹ��Ȩ���ء����򡢺���"                            '����Ȩ������


Const QLLXS = "1:������������Ȩ,2:������������Ȩ,3:���н����õ�ʹ��Ȩ,4:���н����õ�ʹ��Ȩ/���ݣ����������Ȩ,5:լ����ʹ��Ȩ,6:լ����ʹ��Ȩ/���ݣ����������Ȩ,7:���彨���õ�ʹ��Ȩ,8:���彨���õ�ʹ��Ȩ/���ݣ����������Ȩ,9:���سа���ӪȨ,10:���سа���ӪȨ/ɭ�֡���ľ����Ȩ,11:�ֵ�ʹ��Ȩ,12:�ֵ�ʹ��Ȩ/ɭ�֡���ľʹ��Ȩ,13:��ԭʹ��Ȩ,14:ˮ��̲Ϳ��ֳȨ,15:����ʹ��Ȩ,16:����ʹ��Ȩ/����������������Ȩ,17:�޾��񺣵�ʹ��Ȩ,18:�޾��񺣵�ʹ��Ȩ/����������������Ȩ,19:����Ȩ,20:ȡˮȨ,21:̽��Ȩ,22:�ɿ�Ȩ,23:����Ȩ��"

Const QLXZS = "100:��������,101:����,102:����,103:���۳��ʣ���ɣ�,104:����,105:��Ȩ��Ӫ,200:��������,201:��ͥ�а�,202:������ʽ�а�,203:��׼����,204:���,205:��Ӫ"

'���ϵ������(x,y,z,name)
Dim PointArr1(2,4)

'��鼯����
Dim strGroupName1
strGroupName1 = "���߼��"

'��鼯�����
Dim strCheckName1
strCheckName1 = "����߼��"

'�����־
Dim strPromptMessage1
strPromptMessage1 = "���ֶ���д��վ��źͼ����"

'===================================================================================================================

'���ϵ������(x,y,z,name) ���� ��վ�㡢����
Dim PointArr2(2,4)
'��鼯����

Dim strGroupName2
strGroupName2 = "���߼��"
'��鼯�����

Dim strCheckName2
strCheckName2 = "�����߼��"
'�����־

Dim strPromptMessage2
strPromptMessage2 = "���ֶ���д��վ��źͷ�����"

'=============================================================================================================================
'���ϵ������
Dim PointArr3(2,4)

'��鼯����
Dim strGroupName3
strGroupName3 = "���߼��"

'��鼯�����
Dim strCheckName3
strCheckName3 = "���Ƶ������߼��"

'�����־
Dim strPromptMessage3
strPromptMessage3 = "���ֶ���д��վ��źͼ����"


#include"֧����_֧����.vbs"


Sub OnClick()
    
    
    SSParameter.GetParameterINT "AfterAddLine", "CurrentObjID", "0", ObjID
    If ObjID = 0 Then Exit Sub
    ObjCode = SSProcess.GetObjectAttr (objID, "SSObj_Code")
    If ObjCode = "" Then
        OBJID = SSProcess.GetGeoMaxID()
        ObjCode = SSProcess.GetObjectAttr (OBJID, "SSObj_Code")
        If ObjCode = "" Then  Exit Sub
    End If
    
    '=============================================================================================================================================================
    If objcode = 9130241 Then
        GetOnlinePoint1(objID)
        SearchNear1(objID)
    End If
    
    If objcode = 9130251 Then
        GetOnlinePoint2(objID)
        SearchNear2(objID)
    End If
    
    If objcode = 1130212 Then
        GetOnlinePoint3(objID)
        SearchNear3(objID)
        SetYZBC(objID)
        comparelong(objID)
    End If
    
    If objcode = 9414032  Then'�õغ��߱�ע
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "��ע����", "�õغ���",0, "�����ҷ�Χ��", ""
        res = SSProcess.ShowInputParameterDlg ("��ע����" )
        bzny = SSProcess.GetInputParameter ("��ע����" )
        
        SSProcess.SetObjectAttr CLng(objID), "[BiaoZNR]", bzny
        If bzny = "�õغ���"  Then
            SSProcess.SetObjectAttr CLng(objID), "SSObj_Color", RGB(255,0,0)
        ElseIf bzny = "�����ҷ�Χ��"  Then
            SSProcess.SetObjectAttr CLng(objID), "SSObj_Color", RGB(0,0,255)
        Else
            SSProcess.SetObjectAttr CLng(objID), "SSObj_Color", RGB(255,255,255)
        End If
    End If
    
    If objCode = 9470013 Then'�̻���Χ�����Ա�
        SSProcess.ClearInputParameter
        LDBH = SSProcess.ReadEpsDBIni("�̵ر��", "���" ,"")
        SCMJ = SSProcess.GetObjectAttr( ObjID, "SSObj_Area")
        SCMJ = FormatNumber(SCMJ,3, - 1,0,0)
        SSProcess.AddInputParameter "�̵�ͼ�ߺ�",LDBH, 0, "", ""
        SSProcess.AddInputParameter "�̵�����", "�����̻�",  0, "�����̻�,�����Ҷ��̻�,�ݶ��̻�,԰·��԰����װ,����ˮ��", ""
        SSProcess.AddInputParameter "�̵�ϸ��", "԰·��԰����װ",  0, "԰·��԰����װ,����ˮ��,����,�����Ҷ�,�ݶ�", ""
        SSProcess.AddInputParameter "�������", "",0,"", "��д��ֵ"
        SSProcess.AddInputParameter "�Ƿ����̵�", "��", 0, "��,��", ""
        'SSProcess.AddInputParameter "�滮�����̵����", "", 0, "", ""
        result = SSProcess.ShowInputParameterDlg ("¼������")
        If result = 1 Then
            bh = SSProcess.GetInputParameter ("�̵�ͼ�ߺ�")
            lx = SSProcess.GetInputParameter ("�̵�����")
            xl = SSProcess.GetInputParameter ("�̵�ϸ��")
            SFJZLD = SSProcess.GetInputParameter ("�Ƿ����̵�")
            fthd = SSProcess.GetInputParameter ("�������")
            
            If ghspmj = "" Then ghspmj = 0
            SSProcess.SetObjectAttr ObjID, "[LvHTBH]", bh
            SSProcess.SetObjectAttr ObjID, "[LvHLX]", lx
            SSProcess.SetObjectAttr ObjID, "[LvHXL]", xl
            SSProcess.SetObjectAttr ObjID, "[TuBMJ]",SCMJ
            SSProcess.SetObjectAttr ObjID, "[FuTHD]",fthd
            SSProcess.SetObjectAttr ObjID, "[SFJZLD]",SFJZLD
            If  IsNumeric (bh) = True Then
                LDBH = CDbl(bh) + 1
                SSProcess.WriteEpsDBIni "�̵ر��", "���" ,LDBH
            Else
                SSProcess.WriteEpsDBIni "�̵ر��", "���" ,bh
            End If
        End If
    End If
    
    '��Ȼ������
    If objCode = 9210123 Then
        '��Ȼ����Ϣ
        strLCXX_ZRZ = "LCFZXX"                     '¥����Ϣ-��Ȼ��
        strCHZT_ZRZ = "CHZT"                         '���״̬-��Ȼ��
        strLJZLB = "LJZHLB"                         '�߼����б�
        strZRZH_ZRZ = "ZRZH"                          '����-��Ȼ��
        strFWJG_ZRZ = "FWJG"                          '���ݽṹ-��Ȼ��
        strFWJGM_ZRZ = "FWJGNAME"                  '���ݽṹ��-��Ȼ��
        
        '���ԶԻ���
        ZRZ_AttrDlg strFWZH, strFWJG, strZCS, strZTS, strCHZT, strLCFZ2, strsxh,strQH
        SSProcess.SetObjectAttr ObjID, "[QiuH]", strQH
        SSProcess.SetObjectAttr ObjID, "[ZRZH]", strFWZH
        SSProcess.SetObjectAttr ObjID, "[FWJG]", strFWJG
        SSProcess.SetObjectAttr ObjID, "[ZCS]", strZCS
        SSProcess.SetObjectAttr ObjID, "[ZTS]", strZTS
        SSProcess.SetObjectAttr ObjID, "[CHZT]", strCHZT
        SSProcess.SetObjectAttr ObjID, "[LCFZXX]", strLCFZ2
        SSProcess.SetObjectAttr ObjID, "[ZRZSXH]", strsxh
        
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    End If
    
    
    If objCode = 9210872 Then
        '¥��������
        strIDs = SSProcess.SearchNearObjIDs2 (ObjID, 2, "9210123", 0 )
        If strIDs <> ""  And InStr(strIDs,",") = 0 Then
            zrzguid = SSProcess.GetObjectAttr (strIDs, "[ZRZGUID]" )
            strFields = "LCGUID"
            fieldsCount = 1
            sql = "select " & strFields & " from FC_¥����Ϣ���Ա� inner join GeoAreaTB on FC_¥����Ϣ���Ա�.ID=GeoAreaTB.ID where (GeoAreaTB.mark mod 2) <> 0 and ZRZGUID=" & zrzguid & " order BY val(CH)"
            GetMdbValues sql,strFields,fieldsCount,lcAr,lcCount
            lcguid = lcAr(0,0)
            SSProcess.SetObjectAttr ObjID, "SSObj_DataMark", lcguid
        End If
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    End If
    
    '====================================================������
    If objcode = 9130221 Then '֧����
        zd(objID)
    End If
    '====================================================�滮��������
    If objcode = 9310013 Then '�滮��������
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "����������", "",0, "1#¥", ""
        res = SSProcess.ShowInputParameterDlg ("����������" )
        If res = 1 Then
            JianZWMC = SSProcess.GetInputParameter ("����������" )
            SSProcess.SetObjectAttr objID, "[JianZWMC]",  JianZWMC
        End If
    End If
    
    
    '=============================================================================================================================================
    
    If objCode = 3103013 Or objCode = 3103014 Or objCode = 3104003 Or objCode = 3105003 Or objCode = 3108003 Or objCode = 31030131 Then'310301301  Then
        obj_area = SSProcess.GetObjectAttr (objID, "SSObj_Area")
        obj_area = FormatNumber(obj_area,3)
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "������ṹ", "��",0, FWJG_HX, ""
        SSProcess.AddInputParameter "¥����Ŀ", "1",0, FWCS_HX, ""
        
        
        res = SSProcess.ShowInputParameterDlg ("������ṹ" )
        If res = 0 Then
            '���Ż�ˢ��
            SSProcess.ObjectDeal objID, "AddToSelection", "", result
            SSProcess.ObjectDeal 0, "FreeSelectionObjectDisplayList", "", result
            Exit Sub
        End If
        
        FWJG = SSProcess.GetInputParameter ("������ṹ" )
        FWCS = SSProcess.GetInputParameter ("¥����Ŀ" )
        
        SSProcess.SetObjectAttr CLng(objID), "[CONSTRUCT]", FWJG
        SSProcess.SetObjectAttr CLng(objID), "[OGLAYER]", FWCS
        
        
        pointcount = SSProcess.GetObjectAttr (objID, "SSObj_PointCount")
        
        SSProcess.GetObjectPoint objID, 0, x0,  y0,  z0,  ptype0,  name0
        For i = 1 To pointcount - 1
            SSProcess.GetObjectPoint objID, i, x1,  y1,  z1,  ptype1,  name1
            SSProcess.SetObjectPoint objID, i, x1,  y1,  z0,  ptype1,  name1, 1
        Next
        
        
    End If
    '���Ż�ˢ��
    SSProcess.ObjectDeal objID, "AddToSelection", "", result
    SSProcess.ObjectDeal 0, "FreeSelectionObjectDisplayList", "", result
    SSProcess.RefreshView
End Sub

'��ȡMDB��Ϣ
Function GetMdbValues(ByVal sql,ByVal strFields,ByVal fieldsCount,ByRef rs,ByRef rscount)
    
    mdbName = SSProcess.GetProjectFileName()
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    rscount = SSProcess.GetAccessRecordCount (mdbName, sql)
    ReDim rs(rscount,fieldsCount)
    'addloginfo "sql=" & sql & ",fieldsCount=" & fieldsCount
    If rscount > 0 Then
        SSProcess.AccessMoveFirst mdbName, sql
        n = 0
        While SSProcess.AccessIsEOF (mdbName, sql) = False
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            If IsNull(values) Then values = ""
            values = Replace(values,"|","��")
            strs = Split(values,",")
            If UBound(strs) <> - 1 Then
                For i = 0 To fieldsCount - 1
                    rs(n,i) = strs(i)
                Next
            End If
            SSProcess.AccessMoveNext mdbName, sql
            n = n + 1
        WEnd
    End If
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
    
End Function

'�ڵع���
Function ZDGM()
    SSProcess.PushUndoMark
    
    'ɾ��ԭ�����ڵ���
    'SSProcess.ClearSelection
    'SSProcess.ClearSelectCondition
    'SSProcess.SetSelectCondition  "SSObj_Type","=","AREA"
    'SSProcess.SetSelectCondition  "SSObj_Code","=","6803153"
    'SSProcess.SetSelectCondition  "SSObj_LayerName","=","�ڵ�"
    'SSProcess.SelectFilter
    'SSProcess.DeleteSelectionObj
    
    SSProcess.ClearFunctionParameter
    '���ҵ㴦���޾�
    SSProcess.AddFunctionParameter "limitdist=0.0001"
    '���˻��α���
    SSProcess.AddFunctionParameter "SrcArcCodes=9130242,6801332,6803232,6803152"
    'ɾ��Դ����
    SSProcess.AddFunctionParameter "DelSrcArc=0"
    'ɾ���ϴ����ɵ��ص�����
    SSProcess.AddFunctionParameter "DelNewArc=0"
    'ɾ���ϴ����ɵ�ԭ������
    SSProcess.AddFunctionParameter "DelOldTopArea=0"
    '���ݴ�����Ƿ����
    SSProcess.AddFunctionParameter "SaveDB=1"
    '�Ƿ�����������
    SSProcess.AddFunctionParameter "CreateTopArea=1"
    '�������������  ���Ե����1,�����1,ͼ������1/���Ե����2,�����2,ͼ������2
    SSProcess.AddFunctionParameter "NewObject=913022301,9130223"
    '�ж����Ե��ظ��Ĺؼ���
    SSProcess.AddFunctionParameter "LabelKeyFields="
    '�������˻���ѡ��
    '0 �����ɻ���
    '1 ����ͳһ���뻡�Σ�������UniqueArcCodeָ��
    '2 ���ɻ���, ���ж�����״�����ص�ʱ���� ReserveArcOrder���õı���˳�����ȴ�ǰѡȡ
    '3 �Զ����������������ص����»���, ��CreateOverlayArc����
    SSProcess.AddFunctionParameter "CreateTopArc=0"
    '�����ɷ���
    jx = ""
    jx = jx & "250200/��������������/733001/�ڵ�������"
    jx = jx & ",250201/���������߳���/733001/�ڵ�������"
    SSProcess.AddFunctionParameter "CreateOverlayArc=" & jx
    
    SSProcess.TopProcess "�ڵع���"
End Function

'��Ȼ�����ԶԻ���
Function ZRZ_AttrDlg(ByRef strFWZH,ByRef strFWJG,ByRef strZCS,ByRef strZTS,ByRef strCHZT,ByRef strLCFZ2,ByRef strsxh,ByRef strQH)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "���", "", 0, "", ""
    SSProcess.AddInputParameter "��Ȼ����", "1��", 0, "1��,2��,3��", ""
    SSProcess.AddInputParameter "��˳���", "0001", 0, "0001,0002,0003", ""
    SSProcess.AddInputParameter "���ݽṹ", "5:שľ�ṹ", 0, "1:�ֽṹ,2:�ֺ͸ֽ�������ṹ,3:�ֽ�������ṹ,4:��Ͻṹ,5:שľ�ṹ,6:�����ṹ", "���ݽṹȡֵ 1:�ֽṹ,2:�ֺ͸ֽ�������ṹ,3:�ֽ�������ṹ,4:��Ͻṹ,5:שľ�ṹ,6:�����ṹ"
    SSProcess.AddInputParameter "�ܲ���", "1", 0, "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16", ""
    SSProcess.AddInputParameter "������", "1", 0, "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16", ""
    SSProcess.AddInputParameter "���״̬", "2:ʵ��", 0, "1:Ԥ��,2:ʵ��", ""
    SSProcess.AddInputParameter "¥�������Ϣ", "1", 0, "", strLCFZ2
    'SSProcess.AddInputParameter "�߼������б�", "1", 0, "1,1��2,1��2��3", ""
    
    SSProcess.ShowInputParameterDlg title
    strQH = SSProcess.GetInputParameter ("���")
    strFWZH = SSProcess.GetInputParameter ("��Ȼ����")
    strFWJG = SSProcess.GetInputParameter ("���ݽṹ")
    strZCS = SSProcess.GetInputParameter ("�ܲ���")
    strZTS = SSProcess.GetInputParameter ("������")
    strCHZT = SSProcess.GetInputParameter ("���״̬")
    strLCFZ2 = SSProcess.GetInputParameter ("¥�������Ϣ")
    strsxh = SSProcess.GetInputParameter ("��˳���")
    
    '���ݽṹ
    If Replace(strFWJG,":","") <> strFWJG Then
        arFWJG = Split(strFWJG,":")
        strFWJG = arFWJG(0)
        strFWJGMC = arFWJG(1)
    End If
    '���״̬
    If Replace(strCHZT,":","") <> strCHZT Then
        arCHZT = Split(strCHZT,":")
        strCHZT = arCHZT(0)
        'strFWJGMC =arCHZT(1)
    End If
End Function

'��ȡ������Ϣ
'��ȡ���֤������������
Function GetDTByJSGCGHXKZBH (DT)
    
    Dim Fvalues(1000)
    DT = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_���蹤�̽���������Ϣ���Ա�.JianZWMC,GuiHXKZBH FROM (JG_���蹤�̽���������Ϣ���Ա� inner join JG_�õغ�����Ϣ���Ա� on JG_���蹤�̽���������Ϣ���Ա�.YDHXGUID = JG_�õغ�����Ϣ���Ա�.YDHXGUID)  inner join GeoAreaTB on GeoAreaTB.ID = JG_�õغ�����Ϣ���Ա�.ID  WHERE ((GeoAreaTB.mark mod 2) <> 0)  ORDER BY JG_���蹤�̽���������Ϣ���Ա�.JianZWMC;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            
            If values <> "" And  values <> "*" And values <> "NULL" Then
                SSFunc.ScanString values, ",", Fvalues, FvaluesCount
                GHXKZH = Fvalues(1)
                DTMC = Fvalues(0)
                'GHSPJDMJ=Fvalues(2)
                If DT = "" Then
                    DT = GHXKZH & "|" & DTMC
                Else
                    DT = DT & "," & GHXKZH & "|" & DTMC
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function


'*************��ȡ�������彨����Ϣ**************************
Function GetjgclGDxx(ghxkzbh,jzwmc,GHSPZFL,GHSPZGD,GHSPDXZGD,JGDSCS,JGDXCS,JGCLJGLX,GHSPJDMJ,YDHXGUID,JSGHXKZGUID,JZWMCGUID,GuiHYDXKZBH)
    Dim Fvalues(1000)
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_���蹤�̽���������Ϣ���Ա�.GuiHSPZFL,GuiHSPDSCS,GuiHSPDXCS,JunGCLDSCS,JunGCLDXCS,JunGCLJGLX,GuiHSPJDMJ,YDHXGUID,JSGHXKZGUID,JZWMCGUID,GuiHYDXKZBH FROM JG_���蹤�̽���������Ϣ���Ա� WHERE ([JG_���蹤�̽���������Ϣ���Ա�].[ID] > 0 And ([JG_���蹤�̽���������Ϣ���Ա�].[JianZWMC] = '" & jzwmc & "') And ([JG_���蹤�̽���������Ϣ���Ա�].[GuiHXKZBH] = '" & ghxkzbh & "'));"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            SSFunc.Scanstring values,",",Fvalues,Fvaluescount
            GHSPZFL = Fvalues(0)
            GHSPZGD = Fvalues(1)
            GHSPDXZGD = Fvalues(2)
            JGDSCS = Fvalues(3)
            JGDXCS = Fvalues(4)
            JGCLJGLX = Fvalues(5)
            GHSPJDMJ = Fvalues(6)
            YDHXGUID = Fvalues(7)
            JSGHXKZGUID = Fvalues(8)
            JZWMCGUID = Fvalues(9)
            GuiHYDXKZBH = Fvalues(10)
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function




'*************��ȡ�滮���֤��Ϣ**************************
Function Getjhxkzxx(ghxkzbh,jzwmc,XiangMMC)
    Dim Fvalues(6)
    'MSGBOX ghxkzbh
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_���蹤�̹滮���֤��Ϣ���Ա�.XiangMMC FROM JG_���蹤�̹滮���֤��Ϣ���Ա� WHERE ([JG_���蹤�̹滮���֤��Ϣ���Ա�].[ID] > 0  And ([JG_���蹤�̹滮���֤��Ϣ���Ա�].[GuiHXKZBH] = '" & ghxkzbh & "'));"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            SSFunc.Scanstring values,",",Fvalues,Fvaluescount
            XiangMMC = Fvalues(0)
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function



Function GetDTJZWJJXX (DT)
    
    Dim Fvalues(1000)
    DT = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_���蹤�̽���������Ϣ���Ա�.JianZWMC,GuiHXKZBH,JZWMCGUID,JSGHXKZGUID FROM (JG_���蹤�̽���������Ϣ���Ա� inner join JG_�õغ�����Ϣ���Ա� on JG_���蹤�̽���������Ϣ���Ա�.YDHXGUID = JG_�õغ�����Ϣ���Ա�.YDHXGUID)  inner join GeoAreaTB on GeoAreaTB.ID = JG_�õغ�����Ϣ���Ա�.ID  WHERE ((GeoAreaTB.mark mod 2) <> 0)  ORDER BY JG_���蹤�̽���������Ϣ���Ա�.JianZWMC;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            
            If values <> "" And  values <> "*" And values <> "NULL" Then
                SSFunc.ScanString values, ",", Fvalues, FvaluesCount
                GHXKZH = Fvalues(1)
                DTMC = Fvalues(0)
                JZWGUID = Fvalues(2)
                GHXKZGUID = Fvalues(3)
                'GHSPJDMJ=Fvalues(2)
                If DT = "" Then
                    DT = GHXKZH & "|" & DTMC & "|" & JZWGUID & "|" & GHXKZGUID
                Else
                    DT = DT & "," & GHXKZH & "|" & DTMC & "|" & JZWGUID & "|" & GHXKZGUID
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function

'==============================================================================================================================================================================================
'��ֵ����
Function SearchNear1(id)
    x1 = PointArr1(0,0)
    y1 = PointArr1(0,1)
    x2 = PointArr1(1,0)
    y2 = PointArr1(1,1)
    SetLinepoiname1 x1,y1,x2,y2,id
    SetProp1 x1,y1,x2,y2,id
End Function' SearchNear

'��ȡ���ϵĿռ����Ϣ
Function GetOnlinePoint1(id)
    Dim x, y, z, pointtype, name
    pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
    'MsgBox pointcount
    pointcount = transform(pointcount)
    For j = 0 To pointcount - 1
        SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
        x = transform(x)
        y = transform(y)
        z = transform(z)
        PointArr1(j,0) = x
        PointArr1(j,1) = y
        PointArr1(j,2) = z
        PointArr1(j,3) = name
    Next
    'MsgBox PointArr(1,0)
End Function' GetOnlinePoint

'�����ߵķ���ֵ��ˮƽ����(����ֵ����)
Function SetProp1(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = FormatNumber(longtitude,3)
    If x1 < x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 < x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
    End If
    angarr = Split(angles,".", - 1,1)
    If UBound(angarr) > 0 Then
        str = angarr(1)
        dd = ""
        ss = ""
        If Len(str) > 4 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,2)
        End If
        If Len(str) = 3 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,1) & "0"
        End If
        If Len(str) = 2 Then
            dd = Mid(str,1,2)
            ss = "00"
        End If
        If Len(str) = 1 Then
            dd = Mid(str,1,1) & "0"
            ss = "00"
        End If
        If Len(str) = 0 Then
            dd = "00"
            ss = "00"
        End If
    ElseIf UBound(angarr) = 0 Then
        dd = "00"
        ss = "00"
    End If
    SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
    SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "��" & dd & "��" & ss & "��"
End Function' SetProp

'�������ۿ��Ƶ�����
Function SetLinepoiname1(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName1, strCheckName1
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo1 x1,y1,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            'MsgBox id
            ExportInfo1 x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0.001,0,"9130311,9130312,9130217",0)
    idarr = Split(idstring,",", - 1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo1 x2,y2,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[JianCDH]",pointname
        code = SSProcess.GetObjectAttr(idarr(0),"SSObj_Code")
        If code = "9130217" Then
            DiffXY id,"9130216"
        ElseIf code = "9130311" Then
            DiffXY id,"9130211"
        ElseIf code = "9130312" Then
            DiffXY id,"9130212"
        End If
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            ExportInfo1 x2,y2,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[ZhiZDH]",Firstname
        End If
    End If
End Function' SetLinepoiname

'����X,Y��ֵ
Function DiffXY(id,Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SetSelectCondition "SSObj_PointName", "==",PointArr1(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    'MsgBox PointArr(1,3)
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y)
        diffx = Abs(X - PointArr1(1,0))
        diffy = Abs(Y - PointArr1(1,1))
        diffx = FormatNumber(diffx,3)
        diffy = FormatNumber(diffy,3)
        SSProcess.SetObjectAttr id,"[XZuoBCZ]",diffx
        SSProcess.SetObjectAttr id,"[YZuoBCZ]",diffy
    Else
        'MsgBox "������ͬ����" 
        Exit Function
    End If
End Function' DiffXY

'�����鼯����
Function ExportInfo1(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName1, strCheckName1, "�Զ���ű������->" & strCheckName1, strPromptMessage1, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo

'��������ת��
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        MsgBox "��������"
        Exit Function
    End If
    transform = content
End Function

'=============================================================================================================================
Function SearchNear2(id)
    x1 = PointArr2(0,0)
    y1 = PointArr2(0,1)
    x2 = PointArr2(1,0)
    y2 = PointArr2(1,1)
    SetLinepoiname2 x1,y1,x2,y2,id
    SetProp2 x1,y1,x2,y2,id
End Function' SearchNear

'��ȡ���ϵĿռ����Ϣ
Function GetOnlinePoint2(id)
    Dim x, y, z, pointtype, name
    pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
    'MsgBox pointcount
    pointcount = transform(pointcount)
    For j = 0 To pointcount - 1
        SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
        x = transform(x)
        y = transform(y)
        z = transform(z)
        PointArr2(j,0) = x
        PointArr2(j,1) = y
        PointArr2(j,2) = z
        PointArr2(j,3) = name
    Next
    'MsgBox PointArr(1,0)
End Function' GetOnlinePoint

'�����ߵķ���ֵ��ˮƽ����(����ֵ����)
Function SetProp2(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = FormatNumber(longtitude,3)
    If x1 < x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 < x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
    End If
    angarr = Split(angles,".", - 1,1)
    If UBound(angarr) > 0 Then
        str = angarr(1)
        dd = ""
        ss = ""
        If Len(str) > 4 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,2)
        End If
        If Len(str) = 3 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,1) & "0"
        End If
        If Len(str) = 2 Then
            dd = Mid(str,1,2)
            ss = "00"
        End If
        If Len(str) = 1 Then
            dd = Mid(str,1,1) & "0"
            ss = "00"
        End If
        If Len(str) = 0 Then
            dd = "00"
            ss = "00"
        End If
    ElseIf UBound(angarr) = 0 Then
        dd = "00"
        ss = "00"
    End If
    SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
    SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "��" & dd & "��" & ss & "��"
End Function' SetProp

'�������ۿ��Ƶ�����
Function SetLinepoiname2(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName2, strCheckName2
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo2 x1,y1,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            'MsgBox id
            ExportInfo2 x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo2 x2,y2,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[FangXDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            ExportInfo2 x2,y2,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[FangXDH]",Firstname
        End If
    End If
End Function' SetLinepoiname

'�����鼯����
Function ExportInfo2(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName2, strCheckName2, "�Զ���ű������->" & strCheckName2, strPromptMessage2, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo


'===========================================================================================================================================

'��ֵ����
Function SearchNear3(id)
    x1 = PointArr3(0,0)
    y1 = PointArr3(0,1)
    x2 = PointArr3(1,0)
    y2 = PointArr3(1,1)
    SetLinepoiname3 x1,y1,x2,y2,id
    SetProp3 x1,y1,x2,y2,id
End Function' SearchNear

'��ȡ���ϵĿռ����Ϣ
Function GetOnlinePoint3(id)
    Dim x, y, z, pointtype, name
    pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
    'MsgBox pointcount
    pointcount = transform(pointcount)
    For j = 0 To pointcount - 1
        SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
        x = transform(x)
        y = transform(y)
        z = transform(z)
        PointArr3(j,0) = x
        PointArr3(j,1) = y
        PointArr3(j,2) = z
        PointArr3(j,3) = name
    Next
    'MsgBox PointArr(1,0)
End Function' GetOnlinePoint

'�����ߵķ���ֵ��ˮƽ����
Function SetProp3(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = FormatNumber(longtitude,3)
    SSProcess.SetObjectAttr id,"[JCBC]",longtitude
    'SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "��" & dd & "��" & ss & "��"
End Function' SetProp

'������֪�߳�
Function SetYZBC(id)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130211"
    SSProcess.SetSelectCondition "SSObj_PointName", "==",PointArr3(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    'msgbox PointArr3(1,3)
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y)
        yzbc = Sqr((PointArr3(0,0) - X) ^ 2 + (PointArr3(0,1) - Y) ^ 2)
        yzbc = FormatNumber(yzbc,3)
        SSProcess.SetObjectAttr id,"[YZBC]",yzbc
    End If
End Function' SetYZBC

'����߳��ϲ�
Function comparelong(id)
    yzbc = SSProcess.GetObjectAttr(id,"[YZBC]")
    jcbc = SSProcess.GetObjectAttr(id,"[JCBC]")
    yzbc = transform(yzbc)
    jcbc = transform(jcbc)
    bcjc = Abs(yzbc - jcbc)
    SSProcess.SetObjectAttr id,"[BCJC]",bcjc
End Function' comparelong

'���ò�վ���������
Function SetLinepoiname3(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo3 x1,y1,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            'MsgBox id
            ExportInfo3 x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo3 x2,y2,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[JianCDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            ExportInfo3 x2,y2,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[JianCDH]",Firstname
        End If
    End If
End Function' SetLinepoiname

'�����鼯����
Function ExportInfo3(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName3, strCheckName3, "�Զ���ű������->" & strCheckName3, strPromptMessage3, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo


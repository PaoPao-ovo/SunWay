Rem autor
 < Administrator > 
Rem email
XXXX@xxx.com
Rem �ű��ļ���
F
 \ 1208 \ ���Ͽ��� \ DeskTop \ ������ \ Script \ ���� \ ����ͼ���.vbs
Rem ��Ӧ�����ļ���
F
 \ 1208 \ ���Ͽ��� \ DeskTop \ ������ \ ����ģ�� \ ����ͼ.Map
Rem ��������
����ͼ���
Rem ���ű��ļ�Ӧ������ EPS��װĿ¼ \ desktop \ XX̨�� \ Script \ 
Rem framework
gq
Rem framework
471b1e20fe69040339fca38c3d3a189b




Rem special
[����ͼ] ��ͼǰ����ʼ�����ã��ɴ˽���
Function VBS_preMap0(MSGID,mapName,selectID)
    
    Rem �������ؼ�������SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 ֹͣ����ɹ�ͼ
    Rem return = 0 ��������ɹ�ͼ���������á�Ĭ��ֵΪ0��
    
    If MSGID = 0 Then '// �¹��̳�ͼ 
        '// ������Ĵ���.... 
        '// ���ó�ͼ�������ơ��������.... ,������ͼ��·��ÿ�λ���ýű����ص�·�������̲���ͬ����ͨ�������÷�Χ�ߵ������չ����ƴ��
        FileFolder = SSProcess.GetSysPathName (5)
        'CreateFolders FileFolder,"����һ���ݳɹ�"
        SaveFile = FileFolder & "\3�ɹ�\����ͼ.edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",SaveFile
        
    ElseIf MSGID = 1 Then '// �����̳�ͼ 
        '// ������Ĵ���.... 
        
    ElseIf MSGID = 2 Then '// �¹����Զ���Ŀ¼��ͼ(����ѡ�񱣴�·��) 
        '// ������Ĵ���.... 
        
    End If
    
End Function



Rem special
[����ͼ] ��ͼ����ɴ˽���
Function VBS_postMap0(MSGID,mapName,selectID)
    
    Rem ͼ��ID,�ű����������
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem �ű�����������,�ű����������,�ű�������Ӳ���
    Dim str_Name,str_para,str_paraex
    Rem ��ȡ�ֲ�ͼͼ��IDS,���Ӣ�Ķ������
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem ��ȡͼ���ڵ���IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem ��ȡ�ű����������
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    
    '// ������ĳɹ�ͼ������� 
    
    
    Rem �ɹ�ͼϸ�ڷֿ�����
    For i = 0 To ScriptChangeCount - 1
        Rem ��ȡ����������
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem ��ȡ���������
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem ��ȡ������Ӳ���
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        
        
        '// �˴��޴��롢˵��û�нű�������..
    Next
    
    GetTKSX
    DaHui
End Function



Dim g_MapList,g_MapPrePtrfun,g_MapPostPtrfun
Rem �����������޸�
Sub OnClick()
    
    Rem ��ʼ��
    g_MapList = Array("����ͼ")
    g_MapPrePtrfun = Array("VBS_preMap0")
    g_MapPostPtrfun = Array("VBS_postMap0")
    
    Rem ϵͳ��������Ϣ,�û�ѡ��ķ�Χ��ID,�ɹ�ͼ����
    Dim str_msg,str_selectObjid,str_mapName
    
    Rem ��ȡϵͳ���� -  - �û�ѡ��Χ��ID
    SSParameter.GetParameterINT "printMap", "SelectID", - 1, str_selectObjid
    
    Rem ��ȡϵͳ���� -  - ϵͳ��Ϣ ��0���¹��̶̹�Ŀ¼��ͼ��ʼ����Ϣ  1�������̳�ͼ��ʼ����Ϣ  2
    �¹����Զ���Ŀ¼��ͼ��ʼ����Ϣ  3����ͼ����ɽ����ڽű�����ϸ�ڣ�
    SSParameter.GetParameterINT "printMap", "printMSG", - 1, str_msg
    
    Rem ��ȡϵͳ���� -  - ר������
    SSParameter.GetParameterSTR "printMap", "SpecialMapName", "", str_mapName
    
    DistributeMSG str_msg,str_mapName,str_selectObjid
End Sub




Rem ���������������޸�
Function DistributeMSG(MSGid,str_MapName,selectID)
    Dim pFun
    
    For i = 0 To UBound(g_MapList)
        If UCase(g_MapList(i)) = UCase(str_MapName) Then
            If MSGid = 3 Then
                
                Set pFun = GetRef(g_MapPostPtrfun(i))
                Call pFun(MSGid,str_MapName,selectID)
                
            Else
                
                Set pFun = GetRef(g_MapPrePtrfun(i))
                Call pFun(MSGid,str_MapName,selectID)
                
            End If
            Exit For
        End If
    Next
End Function

Function GetTKSX
    If 0 Then '����
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_CODE", "==", "9130224"
        SSProcess.SelectFilter
        geocount = SSProcess.GetSelGeoCount
        If  geocount > 0 Then
            hxid = SSProcess.GetSelGeoValue(0,"SSObj_ID")
            XMMC = SSProcess.GetObjectAttr (hxid,"[XiangMMC]")
        End If
    End If
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_CODE", "==", "9130224"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    If  geocount > 0 Then
        id = SSProcess.GetSelGeoValue(i, "SSObj_id")
        SSProcess.SetObjectAttr id, "[ͼ������]","����ͼ"
    End If
End Function

Function DaHui
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "<>", "9130224,9130411,9310013,9130611"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        SSProcess.SetObjectAttr geoID, "SSObj_Color", RGB(0,0,0)
    Next
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "<>", "����ͼ��,������,����,������Ŀ��,���ؽ�ַ��,����ͼ����,����ע��,���ش�ע��,�����Ե�,�������Ե�"
    SSProcess.SelectFilter
    notecount = SSProcess.GetSelNoteCount()
    For i1 = 0 To notecount - 1
        id = SSProcess.GetSelNoteValue(i1 ,"SSObj_ID" )
        SSProcess.SetObjectAttr id, "SSObj_Color", RGB(0,0,0)
    Next
    
    
End Function

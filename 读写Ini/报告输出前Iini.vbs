
' [���߱�����Ϣ]
' ��� = ""
' ��Ŀ���� = ""
' ��Ŀ��ַ = ""
' ��Ƶ�λ = ""
' ���赥λ = ""
' ί�е�λ = ""
' ��浥λ = ""
' ��ҵʱ�� = ""
' �����ϲ�ֵ = ""
' �߳����ϲ�ֵ = ""

'======================================����Ini&������ֶ�============================================

KeyStr = "���,��Ŀ����,��Ŀ��ַ,��Ƶ�λ,���赥λ,ί�е�λ,��ҵʱ��,���ʱ��,�����ϲ�ֵ,�߳����ϲ�ֵ"

'==========================================��������=====================================================

'�������
Sub OnClick()
    SSProcess.ClearInputParameter
    
    KeyArr = Split(KeyStr,",", - 1,1)
    
    For i = 0 To UBound(KeyArr) - 2
        SSProcess.AddInputParameter KeyArr(i) , SSProcess.ReadEpsIni("���߱�����Ϣ", KeyArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("���߱�����Ϣ¼��")
    
    For i = 0 To UBound(KeyArr)
        SSProcess.WriteEpsIni "���߱�����Ϣ", KeyArr(i) ,SSProcess.GetInputParameter(KeyArr(i))
    Next 'i
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    
    SqlStr = "Update ���¹��ߵ����Ա� SET XMBH = " & "'" & SSProcess.ReadEpsIni("���߱�����Ϣ", "���" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr
    
    SqlStr = "Update ���¹��������Ա� SET XMBH = " & "'" & SSProcess.ReadEpsIni("���߱�����Ϣ", "���" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr

    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
    
    SSProcess.MapMethod "clearattrbuffer",  "���¹��ߵ����Ա�"
    SSProcess.MapMethod "clearattrbuffer",  "���¹��������Ա�"
    SSProcess.RefreshView()
End Sub' OnClick

' [����CAD���]
' ��ע=
' ͼ������=
' ��ҵ��λ=�����ز��Ժ
' ί�е�λ=
' ��������=2023��7�¼������ͼ
' ƽ��������ϵ=���ϳ�������ϵ
' �߳���ϵ=1985���Ҹ̻߳�׼���ȸ߾�0.5�ס�
' ͼʽ=2017���ͼʽ
' ̽��Ա=����
' ����Ա=����
' ��ͼԱ=����
' ���Ա=����


AttrStr = "��ҵ��λ,ί�е�λ,��������,ƽ��������ϵ,�߳���ϵ,ͼʽ,̽��Ա,����Ա,��ͼԱ,���Ա"

Sub OnClick()
    
    SSProcess.ClearInputParameter

    AttrArr = Split(AttrStr,",", - 1,1)
    
    For i = 0 To UBound(AttrArr)
        SSProcess.AddInputParameter AttrArr(i) , SSProcess.ReadEpsIni("����CAD���", AttrArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("����ͼ����Ϣ¼��")
    
    For i = 0 To UBound(AttrArr)
        SSProcess.WriteEpsIni "����CAD���", AttrArr(i) ,SSProcess.GetInputParameter(AttrArr(i))
    Next 'i
    
End Sub

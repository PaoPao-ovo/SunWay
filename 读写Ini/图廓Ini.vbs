
' ���֮ǰ�������

' [����ͼ����Ϣ]
' ��浥λ = �����ز��Ժ
' ���ʱ�估��ʽ = 2020��12�����ֲ�ͼ
' ����ϵ = ���ϳ�������ϵ
' �߳���ϵ = 1985���Ҹ̻߳�׼���ȸ߾�Ϊ0.5m��
' ͼʾ = 2007���ͼʽ
' ����Ա = ��  ��
' ��ͼԱ = ��  ǿ
' ���Ա = ������

KeyStr = "��浥λ,���ʱ�估��ʽ,����ϵ,�߳���ϵ,ͼʾ,����Ա,��ͼԱ,���Ա"

Sub OnClick()
    SSProcess.ClearInputParameter
    
    KeyArr = Split(KeyStr,",", - 1,1)
    
    For i = 0 To UBound(KeyArr)
        SSProcess.AddInputParameter KeyArr(i) , SSProcess.ReadEpsIni("����ͼ����Ϣ", KeyArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("ͼ����Ϣ¼��")
    
    For i = 0 To UBound(KeyArr)
        SSProcess.WriteEpsIni "����ͼ����Ϣ", KeyArr(i) ,SSProcess.GetInputParameter(KeyArr(i))
    Next 'i
    
End Sub' OnClick
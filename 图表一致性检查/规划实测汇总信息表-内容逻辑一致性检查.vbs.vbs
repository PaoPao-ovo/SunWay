
'===================================================����������==================================================

'���������
Dim strGroupName

'���������
Dim strCheckName

'���ģ������
Dim CheckmodelName

'�������
Dim strDescription

'====================================================���=========================================================

'������
Sub OnClick()
    
    ClearCheckRecord
    
    ZhuangCheck
    
    BasementCheck
    
    LvAreaCheck
    
    ConstractDensityCheck
    
    LHPercrntCheck
    
    DSJDCCheck
    
    DXJDCCheck
    
    DSFJDCWCheck
    
    DSFJDCHES
    
    DXFJDCWCheck
    
    DXFJDCHES
    
    LvDAreaCheck
    
    DKLVCheck
    
    JZLDCheck
    
    DGCDCheck
    
    RFMJCheck
    
    ShowCheckRecord
    
End Sub' OnClick

'===================================================��麯��=======================================================

'�������ֵ�봱�������ֵ�Ƿ�һ��
Function ZhuangCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�JZMJ��
    ' 2:��Ȼ����JG_��Ȼ�����Ա����С�JZMJ���ۼƻ��ܡ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "����ָ�������߼����"
    CheckmodelName = "�Զ���ű������->����ָ�������߼����"
    strDescription = "�������ֵ�봱�������ֵ��һ��"
    
    '��ȡ�ܽ������ JZMJ
    SqlStr = "Select Sum(JGSCHZXX.JZMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMJArr,SearchCount
    
    If SearchCount > 0 Then
        JZMJ = Transform(JZMJArr(0))
    Else
        JZMJ = 0
    End If
    
    
    '��ȡ��Ȼ������� SumArea
    SqlStr = "Select Sum(FC_��Ȼ����Ϣ���Ա�.JZMJ) From FC_��Ȼ����Ϣ���Ա� Inner Join GeoAreaTB On FC_��Ȼ����Ϣ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    
    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If JZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' ZhuangCheck

'���������������������ֵ�Ƿ�һ��
Function BasementCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�JZJDMJ��
    ' 2: ����_��(JG_��������������Ա�)���Ա��еġ�JDMJ�������м�¼���ۼӺ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���������������������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->���������������������ֵһ���Լ��"
    strDescription = "���������������������ֵ��һ��"
    
    '��ȡ����� JDMJ
    SqlStr = "Select Sum(JGSCHZXX.JZJDMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JDMJArr,SearchCount
    
    If SearchCount > 0 Then
        JDMJ = Transform(JDMJArr(0))
    Else
        JDMJ = 0
    End If
    
    '��ȡ�������֮�� SumArea
    SqlStr = "Select Sum(JG_��������������Ա�.JDMJ) From JG_��������������Ա� Inner Join GeoAreaTB On JG_��������������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And JG_��������������Ա�.ID > 0"
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    
    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If JDMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' BasementCheck

'�̵�������̵ط�Χ���������ֵ�Ƿ�һ����
Function LvAreaCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�LDMJ��
    ' 2:�̻�Ҫ�����Ա�(LHYS)�С�LHMJ�������м�¼���ۼӺ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�̵�������̵ط�Χ���������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->�̵�������̵ط�Χ���������ֵһ���Լ��"
    strDescription = "�̵�������̵ط�Χ���������ֵ��һ��"
    
    '�̵������ LDMJ
    SqlStr = "Select Sum(JGSCHZXX.LDMJ) From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDMJArr,SearchCount
    
    If SearchCount > 0 Then
        LDMJ = Transform(LDMJArr(0))
    Else
        LDMJ = 0
    End If
    
    '�̻�Ҫ�����֮�� SumLhArea
    SqlStr = "Select Sum(GH_�̻�Ҫ�����Ա�.LHMJ) From GH_�̻�Ҫ�����Ա� Inner Join GeoAreaTB On GH_�̻�Ҫ�����Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_�̻�Ҫ�����Ա�.ID > 0"
    GetSQLRecordAll SqlStr,LHMJArr,LHCount
    
    If LHCount > 0 Then
        SumLhArea = Transform(LHMJArr(0))
    Else
        SumLhArea = 0
    End If
    
    If LDMJ - SumLhArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LvAreaCheck

'�����ܶ������������õ������ֵ�Ƿ�һ��
Function ConstractDensityCheck()
    
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�JZMD��
    ' 2���滮ʵ�������Ϣ��(JGSCHZXX)���С�JZJDMJ��/��YDMJ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�����ܶ������������õ����һ���Լ��"
    CheckmodelName = "�Զ���ű������->�����ܶ������������õ����һ���Լ��"
    strDescription = "�����ܶ������������õ������һ��"
    
    '��ȡ�����ܶ� JZMD
    SqlStr = "Select JGSCHZXX.JZMD From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZMDArr,SearchCount
    
    If SearchCount > 0 Then
        JZMD = Transform(JZMDArr(0))
    Else
        JZMD = 0
    End If
    
    
    '��ȡ������� JDMJ
    SqlStr = "Select JGSCHZXX.JZJDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JDMJArr,SearchCount
    
    If SearchCount > 0 Then
        JDMJ = Transform(JDMJArr(0))
    Else
        JDMJ = 0
    End If
    
    '��ȡ�õ���� YDMJ
    SqlStr = "Select JGSCHZXX.YDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,YDMJArr,SearchCount
    
    If SearchCount > 0 Then
        YDMJ = Transform(YDMJArr(0))
    Else
        YDMJ = 0
    End If
    
    '�����ܶ� Density
    If YDMJ <> 0 Then
        Density = (JDMJ / YDMJ) * 100
    Else
        MsgBox "�������Ϊ�ջ���"
        Exit Function
        Density = 100
    End If
    
    If JZMD - Density <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' ConstractDensityCheck

'�̻���ֵ���̵���������õ����ֵ�Ƿ�һ��
Function LHPercrntCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�LVL��
    ' 2���滮ʵ�������Ϣ��(JGSCHZXX)���С�LDMJ��/��YDMJ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�̻���ֵ���̵���������õ����һ���Լ��"
    CheckmodelName = "�Զ���ű������->�̻���ֵ���̵���������õ����һ���Լ��"
    strDescription = "�̻���ֵ���̵���������õ������һ��"
    
    '��ȡ�̻��� LVL
    SqlStr = "Select JGSCHZXX.LVL From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LVLArr,SearchCount
    
    If SearchCount > 0 Then
        LVL = Transform(LVLArr(0))
    Else
        LVL = 0
    End If
    
    
    '��ȡ�̵���� LDMJ
    SqlStr = "Select JGSCHZXX.LDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDMJArr,SearchCount
    
    If SearchCount > 0 Then
        LDMJ = Transform(LDMJArr(0))
    Else
        LDMJ = 0
    End If
    
    '��ȡ�õ���� YDMJ
    SqlStr = "Select JGSCHZXX.YDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,YDMJArr,SearchCount
    
    
    If SearchCount > 0 Then
        YDMJ = Transform(YDMJArr(0))
    Else
        YDMJ = 0
    End If
    
    'ʵ���ܶ� RealDensity
    If YDMJ <> 0 Then
        RealDensity = (LDMJ / YDMJ) * 100
    Else
        MsgBox "�õ����Ϊ�ջ���"
        Exit Function
        RealDensity = 100
    End If
    
    If RealDensity - LVL <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LHPercrntCheck

'���ϻ�����λ���������ͣ��λ�����Ƿ�һ��
Function DSJDCCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�DSJDCWGS��
    ' 2�����⳵λ���Ա�SWCW�����С�CWLX��<> ���ǻ�����λ�� �����ա�ZSXS��ֵ����ͳ�ƻ��ܣ����*����ϵ��������������ܣ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���ϻ�����λ���������ͣ��λ����һ���Լ��"
    CheckmodelName = "�Զ���ű������->���ϻ�����λ���������ͣ��λ����һ���Լ��"
    strDescription = "���ϻ�����λ���������ͣ��λ������һ��"
    
    '��ȡ���ϻ�������λ���� DSJDCWGS
    SqlStr = "Select JGSCHZXX.DSJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DSJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DSJDCWGS = Transform(DSJDCWGSArr(0))
    Else
        DSJDCWGS = 0
    End If
    
    
    '��ȡ������������� SWCWGS
    SqlStr = "Select GH_���⳵λ���Ա�.ID From GH_���⳵λ���Ա� Inner Join GeoAreaTB On GH_���⳵λ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_���⳵λ���Ա�.CWLX <> '�ǻ�����λ' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SWCWGS = SWCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SWCWGS = 0
    End If
    
    If DSJDCWGS - SWCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DSJDCCheck

'���»�����λ���������ͣ��λ�����Ƿ�һ��
Function DXJDCCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�DXJDCWGS��
    ' 2�����ڳ�λ���Ա�SNCW�����С�CWLX�� <> ���ǻ�����λ�� �����ա�ZSXS��ֵ���л��ܣ���� * ����ϵ��������������ܣ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���»�����λ���������ͣ��λ����һ���Լ��"
    CheckmodelName = "�Զ���ű������->���»�����λ���������ͣ��λ����һ���Լ��"
    strDescription = "���»�����λ���������ͣ��λ������һ��"
    
    '��ȡ���»�������λ���� DXJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DXJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DXJDCWGS = Transform(DXJDCWGSArr(0))
    Else
        DXJDCWGS = 0
    End If
    
    '��ȡ������������� SNCWGS
    SqlStr = "Select GH_���ڳ�λ���Ա�.ID From GH_���ڳ�λ���Ա� Inner Join GeoAreaTB On GH_���ڳ�λ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_���ڳ�λ���Ա�.CWLX <> '�ǻ�����λ' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SNCWGS = SNCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SNCWGS = 0
    End If
    
    If DXJDCWGS - SNCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DXJDCCheck

'���Ϸǻ�����λ��������Ϸǻ�����λ�����Ƿ�һ��
Function DSFJDCWCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�DSFJDCWGS��
    ' 2�����⳵λ���Ա�SWCW�����С�CWLX��=���ǻ�����λ�� �����ա�ZSXS��ֵ����ͳ�ƻ��ܣ����*����ϵ��������������ܣ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���Ϸǻ�����λ��������Ϸǻ�����λ����һ���Լ��"
    CheckmodelName = "�Զ���ű������->���Ϸǻ�����λ��������Ϸǻ�����λ����һ���Լ��"
    strDescription = "���Ϸǻ�����λ��������Ϸǻ�����λ������һ��"
    
    '��ȡ���»�������λ���� DSFJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DSFJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DSFJDCWGS = Transform(DSFJDCWGSArr(0))
    Else
        DSFJDCWGS = 0
    End If
    
    '��ȡ���⳵λ���� SWCWGS
    SqlStr = "Select GH_���⳵λ���Ա�.ID From GH_���⳵λ���Ա� Inner Join GeoAreaTB On GH_���⳵λ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_���⳵λ���Ա�.CWLX = '�ǻ�����λ' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SWCWGS = SWCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SWCWGS = 0
    End If
    
    If DSFJDCWGS - SWCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DSFJDCWCheck

'���Ϸǻ�����λ��ʵ�������
Function DSFJDCHES()
    
    ' 1�����⳵λ���Ա�SWCW�����С�CWLX��=���ǻ�����λ�� �������MJ��*����ϵ����ZSXS���Ƿ���ڳ�λ������CWGS��
    
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���Ϸǻ�����λ��ʵ�������"
    CheckmodelName = "�Զ���ű������->���Ϸǻ�����λ��ʵ�������"
    strDescription = "���Ϸǻ�����λ��ʵ������һ��"
    
    SqlStr = "Select GH_���⳵λ���Ա�.ID From GH_���⳵λ���Ա� Inner Join GeoAreaTB On GH_���⳵λ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_���⳵λ���Ա�.CWLX = '�ǻ�����λ' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    For i = 0 To IDCount - 1
        ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
        Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
        CWGS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWGS]"))
        If Round(Area * ZSXS) - CWGS <> 0  Then
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(IDArr(i),"SSObj_X"),SSProcess.GetObjectAttr(IDArr(i),"SSObj_Y"),0,2,IDArr(i),""
        End If
    Next 'i
    
End Function' DSFJDCHES

'���·ǻ�����λ��������·ǻ�����λ���Ƿ�һ��
Function DXFJDCWCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�DXFJDCWGS��
    ' 2�����ڳ�λ���Ա�SNCW�����С�CWLX��=���ǻ�����λ�� �����ա�ZSXS��ֵ���л��ܣ����*����ϵ��������������ܣ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���·ǻ�����λ��������·ǻ�����λ����һ���Լ��"
    CheckmodelName = "�Զ���ű������->���·ǻ�����λ��������·ǻ�����λ����һ���Լ��"
    strDescription = "���·ǻ�����λ��������·ǻ�����λ������һ��"
    
    '��ȡ���»�������λ���� DXFJDCWGS
    SqlStr = "Select JGSCHZXX.DXJDCWGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DXFJDCWGSArr,SearchCount
    
    If SearchCount > 0 Then
        DXFJDCWGS = Transform(DXFJDCWGSArr(0))
    Else
        DXFJDCWGS = 0
    End If
    
    '��ȡ���⳵λ���� SNCWGS
    SqlStr = "Select GH_���ڳ�λ���Ա�.ID From GH_���ڳ�λ���Ա� Inner Join GeoAreaTB On GH_���ڳ�λ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_���ڳ�λ���Ա�.CWLX = '�ǻ�����λ' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            SNCWGS = SNCWGS + Round(Area * ZSXS)
        Next 'i
    Else
        SNCWGS = 0
    End If
    
    If DXFJDCWGS - SNCWGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DXFJDCWCheck

'���·ǻ�����λ��ʵ�������
Function DXFJDCHES()
    
    ' 1�����ڳ�λ���Ա�SNCW�����С�CWLX��=���ǻ�����λ�� �������MJ��*����ϵ����ZSXS���Ƿ���ڳ�λ������CWGS��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "���·ǻ�����λ��ʵ�������"
    CheckmodelName = "�Զ���ű������->���·ǻ�����λ��ʵ�������"
    strDescription = "���·ǻ�����λ��ʵ������һ��"
    
    SqlStr = "Select GH_���ڳ�λ���Ա�.ID From GH_���ڳ�λ���Ա� Inner Join GeoAreaTB On GH_���ڳ�λ���Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 And GH_���ڳ�λ���Ա�.CWLX = '�ǻ�����λ' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            ZSXS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[ZSXS]"))
            Area = Transform(SSProcess.GetObjectAttr(IDArr(i),"[MJ]"))
            CWGS = Transform(SSProcess.GetObjectAttr(IDArr(i),"[CWGS]"))
            If Round(Area * ZSXS) - CWGS <> 0  Then
                SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(IDArr(i),"SSObj_X"),SSProcess.GetObjectAttr(IDArr(i),"SSObj_Y"),0,2,IDArr(i),""
            End If
        Next 'i
    End If
    
End Function' DXFJDCHES

'�̵�������Ƿ���ڼ����̵����+�����̵���������
Function LvDAreaCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�LDZMJ��=��JZLDMJ��+��DKLDMJ��
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�̵���������"
    CheckmodelName = "�Զ���ű������->�̵���������"
    strDescription = "�̵�������뼯���̵غ͵����̵����֮�Ͳ�һ��"
    
    '��ȡ�̵������ LDZMJ
    SqlStr = "Select JGSCHZXX.LDZMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,LDZMJArr,LDCount
    
    If LDCount > 0 Then
        LDZMJ = Transform(LDZMJArr(0))
    Else
        LDZMJ = 0
    End If
    
    '��ȡ�����̵غ͵����̵����֮�� SumArea
    SqlStr = "Select JGSCHZXX.JZLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZLDMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        JZLDMJ = Transform(JZLDMJArr(0))
    Else
        JZLDMJ = 0
    End If
    
    SqlStr = "Select JGSCHZXX.DKLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DKLDMJArr,DKLDCount
    
    If DKLDCount > 0 Then
        DKLDMJ = Transform(DKLDMJArr(0))
    Else
        DKLDMJ = 0
    End If
    
    SumArea = JZLDMJ + DKLDMJ
    
    If LDZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' LvDAreaCheck

'�����̵�����뵥���̵ط�Χ���������ֵ�Ƿ�һ��
Function DKLVCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�DKLDMJ��
    ' 2���̻�������Ϣ��LHHF�����еġ�MC��=�����̵أ���ͨ����ID_LDK���̵ؿ�ID���̻�Ҫ�����Ա�LHYS���еġ�ID_LDK��ȡ��LHMJ���Ļ���ֵ
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�����̵�����뵥���̵ط�Χ���������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->�����̵�����뵥���̵ط�Χ���������ֵһ���Լ��"
    strDescription = "�����̵�����뵥���̵ط�Χ���������ֵ��һ��"
    
    '�����̵������ DKLDMJ
    SqlStr = "Select JGSCHZXX.DKLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DKLDMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        DKLDMJ = Transform(DKLDMJArr(0))
    Else
        DKLDMJ = 0
    End If
    
    '�����̻���� SumArea
    SqlStr = "Select LHHF.ID_LDK From LHHF Where LHHF.MC = '�����̵�' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            SumArea = SumArea + Transform(SSProcess.GetObjectAttr(IDArr(i),"[LHMJ]"))
        Next 'i
    Else
        SumArea = 0
    End If
    
    
    If DKLDMJ - SumArea <> 0  Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' DKLVCheck

'�����̵�����뼯���̵ط�Χ���������ֵ�Ƿ�һ��
Function JZLDCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�JZLDMJ��
    ' 2���̻�������Ϣ��LHHF�����еġ�MC��=�����̵أ���ͨ����ID_LDK���̵ؿ�ID���̻�Ҫ�����Ա�LHYS���еġ�ID_LDK��ȡ��LHMJ���Ļ���ֵ
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�����̵�����뼯���̵ط�Χ���������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->�����̵�����뼯���̵ط�Χ���������ֵһ���Լ��"
    strDescription = "�����̵�����뼯���̵ط�Χ���������ֵ��һ��"
    
    '�����̵���� JZLDMJ
    SqlStr = "Select JGSCHZXX.JZLDMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,JZLDMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        JZLDMJ = Transform(JZLDMJArr(0))
    Else
        JZLDMJ = 0
    End If
    
    '�����̻���� SumArea
    SqlStr = "Select LHHF.ID_LDK From LHHF Where LHHF.MC = '�����̵�' "
    GetSQLRecordAll SqlStr,IDArr,IDCount
    
    If IDCount > 0 Then
        For i = 0 To IDCount - 1
            SumArea = SumArea + Transform(SSProcess.GetObjectAttr(IDArr(i),"[LHMJ]"))
        Next 'i
    Else
        SumArea = 0
    End If
    
    If DKLDMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' JZLDCheck

'�Ǹ߳��ظ�����Ǹ߳���������Ƿ�һ��
Function DGCDCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�DGCDGS��
    ' 2����GH_����Ҫ�������Ա�Ҫ�ظ���
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�Ǹ߳��ظ�����Ǹ߳��������һ���Լ��"
    CheckmodelName = "�Զ���ű������->�Ǹ߳��ظ�����Ǹ߳��������һ���Լ��"
    strDescription = "�Ǹ߳��ظ�����Ǹ߳����������һ��"
    
    '��ȡ�˷������ DGCDGS
    SqlStr = "Select JGSCHZXX.DGCDGS From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,DGCDGSArr,DGCDGSCount
    
    If DGCDGSCount > 0 Then
        DGCDGS = Transform(DGCDGSArr(0))
    Else
        DGCDGS = 0
    End If
    
    '��ȡ����Ҫ������� XFMGS
    SqlStr = "Select GH_����Ҫ�������Ա�.ID From GH_����Ҫ�������Ա� Inner Join GeoAreaTB On GH_����Ҫ�������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0"
    GetSQLRecordAll SqlStr,XFMGSArr,XFMGSCount
    If XFMGSCount > 0 Then
        XFMGS = Transform(XFMGSCount)
    Else
        XFMGS = 0
    End If
    
    If DGCDGS - XFMGS <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
End Function' DGCDCheck

'�˷���������˷��������������ֵ�Ƿ�һ��
Function RFMJCheck()
    
    ' 1���滮ʵ�������Ϣ��(JGSCHZXX)���С�RFZMJ��
    ' 2���˷����������Ա�RFGNQ���С�JZMJ��ֵ�ۼӺ�
    
    '����¼����
    strGroupName = "ͼ��һ���Լ��"
    strCheckName = "�˷���������˷��������������ֵһ���Լ��"
    CheckmodelName = "�Զ���ű������->�˷���������˷��������������ֵһ���Լ��"
    strDescription = "�˷���������˷��������������ֵ��һ��"
    
    '��ȡ�˷������ RFZMJ
    SqlStr = "Select JGSCHZXX.RFZMJ From JGSCHZXX Where JGSCHZXX.ID > 0 "
    GetSQLRecordAll SqlStr,RFZMJArr,JZLDCount
    
    If JZLDCount > 0 Then
        RFZMJ = Transform(RFZMJArr(0))
    Else
        RFZMJ = 0
    End If
    
    
    '�����˷���� SumArea
    SqlStr = "Select Sum(RF_�˷����������Ա�.JZMJ) From RF_�˷����������Ա� Inner Join GeoAreaTB On RF_�˷����������Ա�.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2) <> 0 "
    GetSQLRecordAll SqlStr,SumAreaArr,SumCount
    
    If SumCount > 0 Then
        SumArea = Transform(SumAreaArr(0))
    Else
        SumArea = 0
    End If
    
    If RFZMJ - SumArea <> 0 Then
        SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,0,0,0,2,0,""
    End If
    
End Function' RFMJCheck

'======================================================�����ຯ��====================================================

'��ջ�������м���¼
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'��ʾ���м���¼
Function ShowCheckRecord()
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' ShowCheckRecord

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

'��������ת��
Function Transform(ByVal Values)
    If Values <> "" Then
        If IsNumeric(Values) = True Then
            Values = CDbl(Values)
        End If
    Else
        Values = 0
    End If
    Transform = Values
End Function'Transform
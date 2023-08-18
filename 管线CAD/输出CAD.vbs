
Sub OnClick()
    '��Ӵ���
    'SSProcess.PushUndoMark 
    fileName = SSProcess.GetSysPathName(5)
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_CODE", "==", "9130235"
    'SSProcess.SetSelectCondition "SSObj_CODE", "==", "9130235,9130225" 
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    If  geocount > 0 Then
        id = SSProcess.GetSelGeoValue(0, "SSObj_id")
        CODE = SSProcess.GetSelGeoValue(0, "SSObj_CODE")
        If CODE = "9130235" Then pathName = fileName & "����ͼ.dwg" '��Ŀ����+�ۺϹ���ͼ.dwg
    End If
    SZDWT id,pathName
End Sub


Function SZDWT(TKID,fileName)
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType", "1"
    SSProcess.SetDataXParameter "Version", "2004"
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_kcad"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_cad"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_cad"
    SSProcess.SetDataXParameter "ExportPathName",fileName
    SSProcess.SetDataXParameter "DataBoundMode", "5"'0(��������)�� 1(ѡ������)�� 2(��ǰͼ��)�� 3(������)��4(ָ������պϵ���)�� 5(ָ��ID�պϵ���)�� 6(����ͼ��)
    SSProcess.SetDataXParameter "DataBoundID", TKID
    'SSProcess.SetDataXParameter "ZoomInOutDataBound", "0.0001"  '���������Χ����������Ϊ��λ��ȱʡֵΪ-0.0001��
    SSProcess.SetDataXParameter "ExportLayerCount", "0" '���ͼ��������������0����ֻ�����ǰ�򿪵�ͼ�㡣
    SSProcess.SetDataXParameter "ZeroLineWidth", "0" '���AutoCAD����ʱ��0�߿�ֽ�ֵ��С�ڻ���ڸ�ֵ���߿����ʱ����Ϊ0��
    SSProcess.SetDataXParameter "AcadColorMethod", "0" '���DWG��ɫʹ�÷�ʽ 0 ����ɫ�ţ��� 1��RGB��ɫֵ��
    SSProcess.SetDataXParameter "ColorUseStatus", "0"       '��ɫʹ��״̬��0����������趨��ɫ�������1���������趨��ɫ�����
    SSProcess.SetDataXParameter "ExplodeObjColorStatus", "0"      '��Ƕ������ɫ�����ʽ��0�������������趨������� 1����������ͬɫ�����
    SSProcess.SetDataXParameter "FontHeightScale", "0.75"
    SSProcess.SetDataXParameter "FontWidthScale", "0.75"
    'SSProcess.SetDataXParameter "FontWidthScale", "FontClass_1190001=0.8,FontClass_1190002=0.8,FontClass_1990001=0.8,FontClass_1990002=0.8,FontClass_1990011=0.8,FontClass_1990012=0.8,FontClass_1990013=0.8,FontClass_1990014=0.8,FontClass_1990015=0.8,FontClass_1990016=0.8,FontClass_1990017=0.8,FontClass_1990018=0.8,FontClass_1990019=0.8,FontClass_1990020=0.8,FontClass_1990021=0.8,FontClass_1990022=0.8,FontClass_1990023=0.8,FontClass_1990031=0.8,FontClass_1990032=0.8,FontClass_1990033=0.8,FontClass_1990034=0.8,FontClass_1990035=0.8,FontClass_1990036=0.8,FontClass_1990037=0.8,FontClass_1990038=0.8,FontClass_1990039=0.8,FontClass_1990040=0.8,FontClass_1990041=0.8,FontClass_1990042=0.8,FontClass_1990043=0.8,FontClass_2190001=0.8,FontClass_2190002=0.8,FontClass_2190003=0.8,FontClass_2190004=0.8,FontClass_2190005=0.8,FontClass_2190006=0.8,FontClass_2290001=0.8,FontClass_2390001=0.8,FontClass_2490001=0.8,FontClass_2590001=0.8,FontClass_2690001=0.8,FontClass_2790001=0.8,FontClass_2990001=0.8,FontClass_3190001=0.8,FontClass_3190002=0.8,FontClass_3190003=0.8,FontClass_3190004=0.8,FontClass_3190005=0.8,FontClass_3190006=0.8,FontClass_3190011=0.8,FontClass_3190012=0.8,FontClass_3190013=0.8,FontClass_3190014=0.8,FontClass_3190015=0.8,FontClass_3290001=0.8,FontClass_3390001=0.8,FontClass_3490001=0.8,FontClass_3590001=0.8,FontClass_3690001=0.8,FontClass_3790001=0.8,FontClass_3890001=0.8,FontClass_3990041=0.8,FontClass_3990042=0.8,FontClass_3990043=0.8,FontClass_3990044=0.8,FontClass_3990045=0.8,FontClass_3990046=0.8,FontClass_3990047=0.8,FontClass_3990048=0.8,FontClass_3990049=0.8,FontClass_3990050=0.8,FontClass_3990051=0.8,FontClass_3990052=0.8,FontClass_3990053=0.8,FontClass_3990054=0.8,FontClass_3990055=0.8,FontClass_3990056=0.8,FontClass_3990057=0.8,FontClass_3990058=0.8,FontClass_3990059=0.8,FontClass_3990060=0.8,FontClass_4190001=0.8,FontClass_4290001=0.8,FontClass_4290002=0.8,FontClass_4290003=0.8,FontClass_4390001=0.8,FontClass_4390002=0.8,FontClass_4390003=0.8,FontClass_4390004=0.8,FontClass_4490001=0.8,FontClass_4590001=0.8,FontClass_4590002=0.8,FontClass_4590003=0.8,FontClass_4690001=0.8,FontClass_4790001=0.8,FontClass_4890001=0.8,FontClass_4990001=0.8,FontClass_4990011=0.8,FontClass_4990012=0.8,FontClass_4990013=0.8,FontClass_4990014=0.8,FontClass_5190001=0.8,FontClass_5290001=0.8,FontClass_5390001=0.8,FontClass_5490001=0.8,FontClass_5990001=0.8,FontClass_6390001=0.8,FontClass_6490001=0.8,FontClass_6590001=0.8,FontClass_6690001=0.8,FontClass_6790001=0.8,FontClass_6790002=0.8,FontClass_6790003=0.8,FontClass_6790004=0.8,FontClass_6790005=0.8,FontClass_6990001=0.8,FontClass_7190001=0.8,FontClass_7290001=0.8,FontClass_7290002=0.8,FontClass_7390001=0.8,FontClass_7490001=0.8,FontClass_7490002=0.8,FontClass_7590001=0.8,FontClass_7590002=0.8,FontClass_7590003=0.8,FontClass_7690001=0.8,FontClass_7690002=0.8,FontClass_7990001=0.8,FontClass_7990002=0.8,FontClass_7990003=0.8,FontClass_7990004=0.8,FontClass_8190001=0.8,FontClass_8290001=0.8,FontClass_8390001=0.8,FontClass_8990001=0.8,FontClass_Z0001=0.7,FontClass_Z0002=1,FontClass_Z0003=1,FontClass_Z0004=1,FontClass_Z0005=1,FontClass_Z0006=0.8,FontClass_Z0007=0.8,FontClass_Z0008=0.8,FontClass_Z0009=1.33,FontClass_Z0010=1,FontClass_Z0011=1,FontClass_Z0012=1,FontClass_Z0013=1,FontClass_Z0212=0.5,FontClass_Z0213=0.8" '���ע���ֿ����űȣ�FontClass_�����=���ű�,���ֱ����д���ű�,��Ĭ��Ϊȫ�����ű�,�ж�������ʱ,�ö��ŷָ�(�� FontClass_0=0.6,FontClass_1=0.7),���ű�ȡֵ��Χ0-1��
    'SSProcess.SetDataXParameter "FontSizeUseStatus","0"               '�����Сʹ��״̬ 0 ����ע�Ƿ���������ָ߿�������� 1 ����ע�������ָ߿������
    SSProcess.SetDataXParameter "OthersExportMode", "3"'���AutoCAD����ʱ����������ʽ�� 0��������룩�� 1��������еĺ�ȣ��� 2��������еı�������3���ó�0��
    SSProcess.SetDataXParameter "OthersExportToZFactor", "0"       '���AutoCAD����ʱ������������Z������ʽ�� 0����������� 1�������
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"   '���Ŵ�ɢ��ʽ�� 0���Զ���ɢ���� 1�����ݱ�����趨��ɢ���� 2��ȫ������ɢ��
    SSProcess.SetDataXParameter "LayerUseStatus", "0"     '�����������ʹ��״̬��0����������趨�����������1���������趨���������
    SSProcess.SetDataXParameter "ExplodeObjLayerStatus", "1"  '��Ƕ����ͼ�������ʽ��0�������������趨������� 1����������ͬ�������
    SSProcess.SetDataXParameter "LineExportMode", "1" '���AutoCAD����ʱ�������������ʽ�� 0 ��ȱʡ��ʽ������ͬ�߳�ʱ��3DPolyline��������ఴ2DPolyline������� 1��ǿ�ư�2DPolyline������� 2�� ǿ�ư�3DPolyline����� 3�� ǿ�ư�Polyline�����
    SSProcess.SetDataXParameter "LineWidthUseStatus", "0"  '�߿�ʹ��״̬��0����������趨�߿��������1���������趨�߿������
    SSProcess.SetDataXParameter "GotoPointsMode", "0"                     '���ͼ�����߻���ʽ�� 0 �������߻����� 1 ��ֻ���߻����ߣ��� 2 ������ͼ�����߻���
    SSProcess.SetDataXParameter "AcadLineWidthMode", "1" 'Acad�߿������ʽ��0 ����� 1 ���
    SSProcess.SetDataXParameter "AcadLineScaleMode", "1"                'Acad���ͱ��������ʽ��0 ������߳�������� 1 ���ǰ�1���
    SSProcess.SetDataXParameter "AcadLineWeightMode","1"               'Acad���������ʽ��0 �����߿� 1 ��� 2 ��� 3 ���߶���
    SSProcess.SetDataXParameter "AcadBlockUseColorMode", "1"        'Acadͼ�������ɫʹ�÷�ʽ��0 ��� 1 ��� 2 �����ʵ��
    SSProcess.SetDataXParameter "AcadLinetypeGenerateMode", "1" '���AutoCAD����ʱ�����������Ƿ����á� 0�����ã� 1�����ã�
    SSProcess.SetDataXParameter "ExplodeObjMakeGroup ", "0"       'AutoCAD����ʱ����ɢ������������ʽ�� 0�������飩�� 1�� ���飬ͬʱҪ��FeatureCodeTB���е�ExtraInfo=1 
    SSProcess.SetDataXParameter "AcadUsePersonalBlockScaleCodes ", "1=7601023"       'AcadUsePersonalBlockScaleCodes ָ��ʹ�����������ı��롣��ʽ1�� ����1=����1,����2;����2=����1,����2��ʽ2�� ���� (�÷�ʽָ�����б����ʹ��ָ���Ŀ����) 
    dwt_path = SSProcess.GetSysPathName (0 ) & "\Acadlin\acad.dwt"
    SSProcess.SetDataXParameter "AcadDwtFileName", dwt_path
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DEFAULT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ӷ���������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ά��ͼ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�������Ե�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����Ե�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�������Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���Ƶ�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ѧ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ͼ����Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ַ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"¥ַ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"POI"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ص�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��·����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ͨ������ʩ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��·�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ߵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ȸ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�̵߳�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ֲ����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ֲ�������ʵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ֲ����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����ע��Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ͼ��Χ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ش�ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ؼ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ؼ�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ۿ��Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����Ȩ�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GPS����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���۲�վ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ʵ���վ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"֧����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�������õط�Χ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ڵؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ڵؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ʵ����Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ں�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ں���ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ں���ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���۷�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ʵ�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ڵ�ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��Ȼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"¥��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ֻ���ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ּ���ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�滮��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ƫ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�滮������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�滮�����ﷶΧ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����ﷶΧ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��������׷�Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������Ĥ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�滮Χǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������ע��Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��ͼ��Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�滮�����ɹ�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"λ��ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ռ�����������ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����߶ȼ���߲�����ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����Ǹ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������ͣ��λ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ǻ�����ͣ��λ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ͣ��λ�ֲ�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�̵ط�Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�̵ؿ���ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"Ժ�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"TERP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GTFA"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GTFL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ط�����ԭ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ�߸�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���Ҳ�һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ط���һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"Ȩ������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������Ŀ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"������������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ͼ������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"����ƽ��ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"�����Ա�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"���غ���ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ͣ��λ��Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"KZ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"KZ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"SX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JMD_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JT_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"GX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"JJ_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"DM_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ZB_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"QT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"TK"
    
    startIndex = 0
    
    SSProcess.SetDataXParameter "LayerRelationCount","2000"
    'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"DEFAULT:0:0:0:0:0"
    SSProcess.ExportData
End Function

Function AddOne(ByRef startIndex)
    startIndex = startIndex + 1
    AddOne = startIndex
End Function


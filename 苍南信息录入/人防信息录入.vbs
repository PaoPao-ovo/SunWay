
' strfieldsName = "��Ŀ����,��Ŀ��ַ,��Ŀ���,��Ƶ�λ,���赥λ,ί�е�λ,��浥λ,���ʱ��,սʱ����,ƽʱ����,�����˷�Ӧ�����,סլϵ��,��סլϵ��,�����ṹ"
' strfields = "XiangMMC,XiangMDZ,XiangMBH,SheJDW,JianSDW,WeiTDW,CeLDW,CeLRQ,ZSGN,PSGN,SPRFYJMJ,ZhuZXS,FZhuZXS,JZJG"
strfieldsName = "��浥λ,���ʱ��,սʱ����,ƽʱ����,�����˷�Ӧ�����,סլϵ��,��סլϵ��,�����ṹ,��Ŀ����,��Ŀ��ַ,��Ŀ���,��Ƶ�λ,���赥λ,ί�е�λ"
strfields = "CeLDW,CeLRQ,ZSGN,PSGN,SPRFYJMJ,ZhuZXS,FZhuZXS,JZJG,XiangMMC,XiangMDZ,XiangMBH,SheJDW,JianSDW,WeiTDW"
Sub OnClick()
    
    geocount = GetFeatureCount( "9130223", geocount)
    strfieldsNameList = Split(strfieldsName,",")
    strfieldsList = Split(strfields,",")
    SSProcess.ClearInputParameter
    For i = 0 To geocount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ObjID = id
        GetFieldValues id, strfields, strValues, fieldCount
        For i1 = 0 To fieldCount - 1
            If fieldCount - 1 = UBound(strfieldsNameList) Then
                If strfieldsNameList(i1) = "סլϵ��" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "7%,8%", "�س�סլ8%���ص���סլ7%��"
                ElseIf strfieldsNameList(i1) = "��סլϵ��" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "5%,4%", "�س�ס��סլ5%���ص����סլ4%��"
                ElseIf strfieldsNameList(i1) = "�����ṹ" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "�ֻ�", ""
                ElseIf strfieldsNameList(i1) = "���ʱ��" Then
                    SSProcess.AddInputParameter strfieldsNameList(i1), GetNowTime,0, "", ""
                Else
                    SSProcess.AddInputParameter strfieldsNameList(i1), strValues(i1),0, "", ""
                End If
            End If
        Next
    Next
    res = SSProcess.ShowInputParameterDlg ("�˷���Ŀ��Ϣ¼��" )
    If res = 1 Then
        For i = 0 To UBound(strfieldsNameList)
            value = SSProcess.GetInputParameter (strfieldsNameList(i))
            SSProcess.SetObjectAttr ObjID, "[" & strfieldsList(i) & "]", value
        Next
    End If
    SSProcess.ObjectDeal objID, "FreeDisplayList", parameters, result
    SSProcess.RefreshView
    
End Sub


'��ȡҪ�ظ���
Function GetFeatureCount(ByVal Code,ByRef geocount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code","==",Code
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    GetFeatureCount = geocount
End Function

'��ȡ�ֶ���������
Function GetFieldValues(ByVal id,ByVal fields,ByRef strValues(),ByRef fieldCount)
    ReDim strValues(fieldCount)
    fieldCount = 0
    fieldsList = Split(fields,",")
    For i = 0 To UBound(fieldsList)
        values = SSProcess.GetObjectAttr (id, "[" & fieldsList(i) & "]")
        ReDim Preserve strValues(fieldCount)
        strValues(fieldCount) = values
        fieldCount = fieldCount + 1
    Next
End Function

'��ȡ��ǰϵͳʱ��
Function GetNowTime()
    GetNowTime = FormatDateTime(Now(),1)
End Function' GetNowTime
'ȫ�ֱ���
Dim vArray(1000)

'ͼ�����ƺ���
Function DrawTuLi()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9460093 'ͼ��
    SSProcess.SelectFilter
    GeoCount = SSProcess.GetSelGeoCount()
    If Geocount > 0 Then
        For i = 0 To GeoCount - 1
            ID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            SSProcess.GetObjectPoint ID, 1, x, y, z, pointtype, name '���½ǵ�
        Next
        innerids = SSProcess.SearchInnerObjIDs(ID , 10 ,"9460081,9460033,9460003,9450013,9420005,9450014", 0)
        If innerids <> "" Then
            SSFunc.ScanString innerids, ",", vArray, nCount
            ZDrawCode = ""
            For j = 0 To nCount - 1
                DrawCode = SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
                DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
                DrawName = SSProcess.GetFeatureCodeInfo (DrawCode,"ObjectName")
                'MsgBox DrawName 
                If ZDrawCode = "" Then
                    ZDrawCode = DrawCode
                    ZDrawColor = DrawColor
                    ZDrawName = DrawName
                Else
                    If Replace(ZDrawCode,DrawCode,"") = ZDrawCode Then
                        ZDrawCode = ZDrawCode & "," & DrawCode
                        ZDrawColor = ZDrawColor & "," & DrawColor
                        ZDrawName = ZDrawName & "," & DrawName
                    End If
                End If
            Next
            
            '�������
            arDrawCode = Split(ZDrawCode,",")
            count = UBound(arDrawCode) + 2
            DrawBorder x,y,0,"RGB(255,255,255)",ID,count
            
            CreateWindows ZDrawCode,ZDrawColor,ZDrawName,Code,Color,Name

            If Code <> "" Then
                '�����ڲ�ͼ��
                DrawInner x,y,ID,Code,Color,Name
                
                '���ƹ̶���ע��
                
                DrawPoint x - 43,y + 11,"9000001",ID
            End If
        End If
    End If
End Function

'����ѡ�񵯴�,����ѡ���Code
Function CreateWindows(ByVal ZDrawCode,ByVal ZDrawColor,ByVal ZDrawName,ByRef Code,ByRef Color,ByRef Name)
    
    ZDrawCodeArr = Split(ZDrawCode,",", - 1,1)
    ReDim Preserve ZDrawCodeArr(UBound(ZDrawCodeArr) + 1)
    
    For i = 0 To UBound(ZDrawCodeArr)
        If i < UBound(ZDrawCodeArr) Then
            ZDrawCodeArr(i) = ZDrawCodeArr(i) & "��" & SSProcess.GetFeatureCodeInfo(ZDrawCodeArr(i),"ObjectName") & "��"
        Else
            ZDrawCodeArr(i) = ""
        End If
    
    Next 'i

    '�ƿ�
    Code = ""
    Color = ""
    Name = ""
    
    RecordShortListCount = UBound(ZDrawCodeArr) + 1

    ResVal_Dlg = SSFunc.SelectListAttr("ѡ���б�","��ѡ�����б�","ѡ�������б�",ZDrawCodeArr,RecordShortListCount)
    If ResVal_Dlg = 1 Then
        If RecordShortListCount > 0 Then
            Size = UBound(ZDrawCodeArr)
            If Size > 0 Then
                For i = 0 To RecordShortListCount  - 1
                    If Code = "" Then
                        CodeArr = Split(ZDrawCodeArr(i),"��", - 1,1)
                        Code = CodeArr(0)
                        Color = SSProcess.GetFeatureCodeInfo(Code,"LineColor")
                        Name = SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
                    Else
                        CodeArr = Split(ZDrawCodeArr(i),"��", - 1,1)
                        Code = Code & "," & CodeArr(0)
                        Color = Color & "," & SSProcess.GetFeatureCodeInfo(CodeArr(0),"LineColor")
                        Name = Name & "," & SSProcess.GetFeatureCodeInfo(CodeArr(0),"ObjectName")
                    End If
                Next 'i
            End If
        End If
    End If
End Function' CreateWindows

'����ͼ��
Function DrawInner(x,y,polygonID,ZDrawCode,ZDrawColor,ZDrawName)
    FountWith = 200
    FountHight = 200
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    count = UBound(arDrawCode) + 3
    '�����ݻ���ͼ��
    
    If arDrawCode(i) = 9460033 Or arDrawCode(i) = 9460003  Then
        If arDrawCode(i) = 9460033  Then
            DrawLine x - 20,y + 3 * (count - 3 - i) + 3,x - 15,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),polygonID
            DrawNote x - 13,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),FountWith,FountHight,"����ǻ�����λ",polygonID
        End If
        
        If arDrawCode(i) = 9460033  Then
            DrawLine x - 20,y + 3 * (count - 3 - i) + 3,x - 15,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),polygonID
            DrawNote x - 13,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),FountWith,FountHight,"����ǻ�����λ",polygonID
        End If
        
    Else
        For i = 0 To count - 3
            DrawLine x - 20,y + 3 * (count - 3 - i) + 3,x - 15,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),polygonID
            DrawNote x - 13,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),FountWith,FountHight,arDrawName(i),polygonID
        Next
    End If
End Function

'���Ƶ�Ҫ��

Function DrawPoint(x,y,code,polygonID)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������߻���
Function DrawBorder(x,y,code,color,polygonID,count)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    'SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
    If count >= 5 Then
        '�ڿ���
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        SSProcess.AddNewObjPoint x, y + 3 * count + 10 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 60, y + 3 * count + 10 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 60, y , 0, 0, ""
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        
        '�����
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        SSProcess.AddNewObjPoint x, y + 3 * count + 1 + 10 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 61, y + 3 * count + 1 + 10 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 61, y , 0, 0, ""
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        SSProcess.AddNewObjToSaveObjList
        SSProcess.SaveBufferObjToDatabase
        
        '���Ʊ���
        DrawTitle x - 30,y + 3 * count - 2 + 10,400,400
    Else
        '�ڿ���
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        SSProcess.AddNewObjPoint x, y + 25 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 60, y + 25 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 60, y , 0, 0, ""
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        
        '�����
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        SSProcess.AddNewObjPoint x, y + 26 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 61, y + 26 , 0, 0, ""
        SSProcess.AddNewObjPoint x - 61, y , 0, 0, ""
        SSProcess.AddNewObjPoint x, y, 0, 0, ""
        SSProcess.AddNewObjToSaveObjList
        SSProcess.SaveBufferObjToDatabase
        
        '���Ʊ���
        DrawTitle x - 30,y + 23,400,400
    End If
End Function

'������Ҫ��

Function DrawArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������Ҫ��

Function DrawLine(x1,y1,x2,y2,code, color, polygonID)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'����ע��
Function DrawNote(x, y, code, color, width, height, fontString,polygonID)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'���Ʊ���
Function DrawTitle(x, y, width, height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", "ͼ ��"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ��"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Sub OnClick()
    '��Ӵ���
    DrawTuLi
End Sub
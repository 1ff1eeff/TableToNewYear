'Public objDicServiceDays As Dictionary '��� �� ������� ��������� ������������.
Public arr��������() '��� ���������� � �������������� ��������.
Public arr������() '������� ������� � �����������.
Public i��������, i���������, i������������������������, i������������������ As Integer

Sub fn����������������������()

    '=====================================================================================
    '           ��� ���������� ������� ���������� ��������!
    '=====================================================================================
    
    i������������������������ = 1 '���������� ����� ��������� � ������ � �������.
    i������������������ = 7
    i��������� = 2023 '����������� ��� � �������.
    i�������� = 2024  '����� ���.
    
    '���� ����������� � �������������� �������� ���� (� �������� ����, � ������� �������� � Excel):
    arr�������� = Array(45292, 45293, 45294, 45295, 45296, 45297, 45298, 45299, 45345, 45346, 45347, 45359, 45360, _
                45361, 45410, 45411, 45412, 45413, 45421, 45422, 45423, 45424, 45455, 45599, 45600, 45655, 45656, 45657)
    
    '���� ������� ������ ��� ����������� (� �������� ����, � ������� �������� � Excel):
    arr������ = Array(45409, 45654)
    
    '===================================================================================================================
    
'    If objDicServiceDays Is Nothing Then
'        Set objDicServiceDays = New Dictionary
'    End If
'    objDicServiceDays.RemoveAll
    
    Select Case Selection.Areas.Count '������������� ���������� ���������� ��������.
    Case 0: '���� �� ������� �� ���� �������.
        MsgBox "�� ������ �������� �����!"
        End
    Case 1: '���� ������� ���� �������.
        Set Table = Selection '������� �� ���������� �������.
    Case Else '���� ������� ��������� ��������.
        MsgBox "�������� ������� ������!"
        End
    End Select
    
    i������������������ = Table.Columns.Count
    
    For i = 3 To Table.Rows.Count '�������� �� ������� �������.
    
        For j = i������������������ To i������������������ Step i������������������������ '�������� �� �������� ������� � �������� �����.
            
'            If i = 26 Then
'                MsgBox "STOP"
'            End If
'
            i������������ = Table(i, j) '����� �� ������ �������.
            If IsEmpty(i������������) = False And IsNumeric(i������������) Then  '���� � ������ ���� �����.
                               
                s������������� = Table(1, j).MergeArea(1) '���������� ������� �����.
                i������������� = Month("01 " & s������������� & " 1900") '���������� ����� ������.
                d������������ = DateSerial(i���������, i�������������, i������������)  '��������� �������� ����.
                d�������������� = fn�������������������(d������������) '�������������.
                
                If j > 1 Then '���� ���� ������ �����.
                    s��������������� = Table(i, j - i������������������������) '������� ���������� ������.
                    If Not IsEmpty(s���������������) Then '���� � ������ ���� ����.
                        s����� = Table(1, j - i������������������������).MergeArea(1) '���������� �����.
                        s��������������������� = Table(1, j - 1).MergeArea(1)
                        i����� = Month("01 " & s��������������������� & " 1900") '���������� ����� ������.
                        d�������������� = DateSerial(i��������, i�����, s���������������) '��������� ����.
                        If d�������������� > d�������������� Then '���� ���� � ������� ������ ������ ��� � ����������.
                            Table(i, j - i������������������������) = Day(d��������������) '������ �������.
                            d�������������� = d��������������
                        End If
                    End If
                End If
                
                Table(i, j) = Day(d��������������) '��������� ���� � ������� ������.
                
            End If
        Next j
    Next i
End Sub

Function fn�������������������(d������������) As Date
   
    'Dim i��������� As Integer '�� ������� ���� �������� ����.
    's������������� = MonthName(Month(d������������)) '���������� ����� ��������� ���.
       
    i������������������������ = Weekday(CDate(d������������), vbMonday) '���������� ����� ��� � ������ ��������� ���.
    l�������������� = CLng(DatePart("ww", d������������)) '���������� ����� ������ ��������� ���.
    
    i����������������������� = GetDateByWeeknum(ByVal l��������������, i��������)
    d��������� = DateAdd("d", i������������������������ - 1, i�����������������������) '��������� ����� ����.
    i��������������������� = Weekday(CDate(d���������), vbMonday)
    i����������� = DatePart("ww", d���������)
        
    d�������������� = fn��������������������(ByVal DateAdd("d", -1, d���������))  '��������� ������� ���� �� ���������.
    i�������������������������� = Weekday(CDate(d��������������), vbMonday)
    i���������������� = DatePart("ww", d��������������)
        
    If i���������������� <> i����������� Then '���� ������ ����������.
        d�������������� = fn��������������������(ByVal d���������, -1) '�������� � ������ �����������.
        i�������������������������� = Weekday(CDate(d��������������), vbMonday)
        i���������������� = DatePart("ww", d��������������)
    End If
    
    'MsgBox "C����� ����: " & d������������ & vbLf & "����� ��� � �����e: " & i������������������������ & vbLf & "����� ������: " & i�������������� & _
        vbLf & vbLf & "����� ����: " & d��������� & vbLf & "����� ��� � �����e: " & i��������������������� & vbLf & "����� ������: " & i����������� & _
        vbLf & vbLf & "���������� ����: " & d�������������� & vbLf & "����� ��� � �����e: " & i�������������������������� & vbLf & "����� ������: " & i����������������
            
    
    'varTemp = objDicServiceDays.Item(CLng(d��������������))

    fn������������������� = d��������������
    
End Function
'---------------------------------------------------------------------------------------
' Purpose: GetDateByWeeknum - ��������� ���� ������������ �� ��������� ��������� ������ ������
'               lWeeknum    - ����� ������, ���� ������������ ��� ������� ���� ��������
'               lYear       - ����� ����. ���� �� ������ - ������� ������� ���
'               IsISO       - ���� True ��� 1, �� ������ ������� ��������� ��,
'                             � ������� ������ ��� ����������� �������
'                             ���� False, 0 ��� �� ������ - ������� ������ ������ ������
'---------------------------------------------------------------------------------------
Function GetDateByWeeknum(ByVal lWeeknum As Long, _
                          Optional ByVal lYear& = 0, _
                          Optional IsISO = False)
    If lYear = 0 Then
        lYear = Year(Date)
    End If
    
    Dim lr&, lcnt&, dt As Date, dStart As Date, dEnd As Date
    dStart = DateSerial(lYear, 1, 1)
    dEnd = DateSerial(lYear, 12, 31)
    If IsISO Then '���� ���� ����� ������ �������
        For lr = dStart To dStart + 8
            If VBA.Weekday(lr, vbMonday) = 4 Then
                dStart = lr - 3
                Exit For
            End If
        Next
    Else '���� ���� ����� �� ������ �������, � ������ ����� ������
        For lr = dStart To dStart + 8
            If VBA.Weekday(lr, vbMonday) = 1 Then
                lcnt = 1
            Else
                lcnt = lcnt + 1
            End If
            If lcnt = 7 Then
                dStart = lr - 6
                Exit For
            End If
        Next
    End If
    For lr = dStart To dEnd Step 7
        If Int((lr - dStart + 1) / 7) = lWeeknum - 1 Then
            GetDateByWeeknum = CDate(lr)
            Exit For
        End If
    Next
End Function

Function fn��������������������(ByVal d���� As Date, _
    Optional ByVal i������������ As Integer = 1) As Date
    If i������������ > 0 Then
        Do While i������������ > 0
            If fn��������������(d���� + 1) Then
                i������������ = i������������ - 1
            End If
            d���� = d���� + 1
        Loop
    Else
        Do While i������������ < 0
            If fn��������������(d���� - 1) Then
                i������������ = i������������ + 1
            End If
            d���� = d���� - 1
        Loop
    End If
    fn�������������������� = d����
End Function

Function fn��������������(ByVal d���� As Date) As Boolean
    If d���� = 0 Then Exit Function
    If Weekday(d����, vbMonday) <> 6 And Weekday(d����, vbMonday) <> 7 _
    And Not fn����������������������(d����) Or fn�����������������(d����) Then
    
'        If Not fn�������������������(d����) Then
            fn�������������� = True
'        End If
    Else
        fn�������������� = False
    End If
End Function

'Private Function fn�������������������(d���� As Date) As Boolean
'    b���������������� = objDicServiceDays.Exists(CLng(d����))
'    fn������������������� = b����������������
'End Function


Private Function fn����������������������(d���� As Date) As Boolean
    Dim i As Long
    Dim iRow As Integer, varTmp
    Static objDicHolidays As Dictionary
    If objDicHolidays Is Nothing Then
        Set objDicHolidays = New Dictionary
        '���� ���������� ����������� �������� "������������", ��������� ��� ���������� � �������.
        If TypeOf Evaluate("������������") Is Range Then
            With Range("������������")
                For iRow = 1 To .Rows.Count
                    varTmp = objDicHolidays.Item(CLng(.Cells(iRow, 1).Value))
                Next iRow
            End With
        End If

        For i = LBound(arr��������) To UBound(arr��������)
            varTmp = objDicHolidays.Item(arr��������(i))
        Next i
    End If
    fn���������������������� = objDicHolidays.Exists(CLng(d����))
End Function

Private Function fn�����������������(d���� As Date) As Boolean
    Dim i As Long
    Dim iRow As Integer, varTmp
    Static objDicWorkingSaturdays As Dictionary
    If objDicWorkingSaturdays Is Nothing Then
        Set objDicWorkingSaturdays = New Dictionary
        '���� ���������� ����������� �������� "��������������", ��������� ��� ���������� � �������.
        If TypeOf Evaluate("��������������") Is Range Then
            With Range("��������������")
                For iRow = 1 To .Rows.Count
                    varTmp = objDicWorkingSaturdays.Item(CLng(.Cells(iRow, 1).Value))
                Next iRow
            End With
        End If
    End If
    For i = LBound(arr������) To UBound(arr������)
        varTmp = objDicWorkingSaturdays.Item(arr������(i))
    Next i
    fn����������������� = objDicWorkingSaturdays.Exists(CLng(d����))
End Function
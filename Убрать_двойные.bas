Public i������������������������, i������������������ As Integer

Sub fn�������������()
    
    '=====================================================================================
    '           ��� ���������� ������� ���������� ��������!
    '=====================================================================================
    
    i������������������������ = 1 '���������� ����� ��������� � ������ � �������.
    i������������������ = 7 '������� � �������� ���������� ������ �������.
    
    '=====================================================================================
        
    Select Case Selection.Areas.Count '������������� ���������� ���������� ��������.
    Case 0: '���� �� ������� �� ���� �������.
        MsgBox "�� ������ �������� �����!"
        End
    Case 1: '���� ������� ���� �������.
        Set rng������� = Selection '������� �� ���������� �������.
    Case Else '���� ������� ��������� ��������.
        MsgBox "�������� ������� ������!"
        End
    End Select
    
    Dim i As Integer
    i = 3
    
    Do While i < rng�������.Rows.Count '�������� �� ������� �������.
        Do While fn������������������(ByVal rng�������, ByVal i)
        
            rng�������.Rows(i + 1).Insert
            rng�������.Range(Cells(i, 1), Cells(i + 1, 1)).Merge
            rng�������.Range(Cells(i, 4), Cells(i + 1, 4)).Merge
            rng�������.Range(Cells(i, 5), Cells(i + 1, 5)).Merge
            rng�������.Range(Cells(i, 6), Cells(i + 1, 6)).Merge
            Union(Selection, Selection.Offset(1, 0)).Select
            Set rng������� = Selection '������� �� ���������� �������.

            For j = i������������������ To rng�������.Columns.Count Step i������������������������ '�������� �� �������� �������.
                s������������� = rng�������(i, j) '������ �� ������� ������.
                If IsEmpty(s�������������) = False And IsNumeric(s�������������) Then  '���� � ������ ���� �����.
                    If InStr(s�������������, " ") > 0 Then '���� � ������� ������ ��������� ����� ��������� ���������.
                        
                        Dim arr���() As String
                        arr��� = Split(s�������������) '����������� �� ��������� - ������.
                        rng�������(i, j).ClearContents '������� ������ �������.
                        
                        rng�������(i + 1, j) = arr���(GetArrLength(arr���) - 1)
                        
                        ReDim Preserve arr���(UBound(arr���) - 1)
                        For Each i���� In arr���
                            If IsNumeric(i����) Then
                                rng�������(i, j) = rng�������(i, j) & " " & i����
                            End If
                        Next
                    End If
                End If
            Next j
        Loop
        i = i + 1
    Loop
End Sub

Function fn������������������(ByVal rng������� As Range, ByVal i����������� As Integer) As Boolean
        For j = i������������������ To rng�������.Columns.Count Step i������������������������ '�������� �� �������� �������.
            s������������� = rng�������(i�����������, j) '������ �� ������� ������.
            If IsEmpty(s�������������) = False And IsNumeric(s�������������) Then  '���� � ������ ���� �����.
                If InStr(s�������������, " ") > 0 Then '���� � ������� ������ ��������� ����� ��������� ���������.
                    fn������������������ = True
                    'MsgBox "������:   " & i����������� & vbLf & "�������: " & j & vbLf & vbLf & rng�������(i�����������, j)
                    Exit Function
                End If
            End If
        Next j
End Function


Public Function GetArrLength(a As Variant) As Long
   If IsEmpty(a) Then
      GetArrLength = 0
   Else
      GetArrLength = UBound(a) - LBound(a) + 1
   End If
End Function
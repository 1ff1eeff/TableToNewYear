Public iРасстояниеМеждуКолонками, iРасстояниеДоДанных As Integer

Sub fnУбратьДвойные()
    
    '=====================================================================================
    '           Тут необходимо указать актуальные значения!
    '=====================================================================================
    
    iРасстояниеМеждуКолонками = 1 'Расстояние между колонками с датами в таблице.
    iРасстояниеДоДанных = 7 'Столбец с которого начинаются данные таблицы.
    
    '=====================================================================================
        
    Select Case Selection.Areas.Count 'Рассматриваем количество выделенных областей.
    Case 0: 'Если не выбрана ни одна область.
        MsgBox "Не выбран диапазон ячеек!"
        End
    Case 1: 'Если выбрана одна область.
        Set rngТаблица = Selection 'Таблица из выделенной области.
    Case Else 'Если выбраны несколько областей.
        MsgBox "Выделите смежные ячейки!"
        End
    End Select
    
    Dim i As Integer
    i = 3
    
    Do While i < rngТаблица.Rows.Count 'Проходим по строкам таблицы.
        Do While fnВСтрокеЕстьДвойные(ByVal rngТаблица, ByVal i)
        
            rngТаблица.Rows(i + 1).Insert
            rngТаблица.Range(Cells(i, 1), Cells(i + 1, 1)).Merge
            rngТаблица.Range(Cells(i, 4), Cells(i + 1, 4)).Merge
            rngТаблица.Range(Cells(i, 5), Cells(i + 1, 5)).Merge
            rngТаблица.Range(Cells(i, 6), Cells(i + 1, 6)).Merge
            Union(Selection, Selection.Offset(1, 0)).Select
            Set rngТаблица = Selection 'Таблица из выделенной области.

            For j = iРасстояниеДоДанных To rngТаблица.Columns.Count Step iРасстояниеМеждуКолонками 'Проходим по столбцам таблицы.
                sТекущаяЯчейка = rngТаблица(i, j) 'Данные из текущей ячейки.
                If IsEmpty(sТекущаяЯчейка) = False And IsNumeric(sТекущаяЯчейка) Then  'Если в ячейке есть число.
                    If InStr(sТекущаяЯчейка, " ") > 0 Then 'Если в текущей ячейке несколько чисел разделены пробелами.
                        
                        Dim arrДни() As String
                        arrДни = Split(sТекущаяЯчейка) 'Разделитель по умолчанию - пробел.
                        rngТаблица(i, j).ClearContents 'Очищаем ячейку таблицы.
                        
                        rngТаблица(i + 1, j) = arrДни(GetArrLength(arrДни) - 1)
                        
                        ReDim Preserve arrДни(UBound(arrДни) - 1)
                        For Each iДень In arrДни
                            If IsNumeric(iДень) Then
                                rngТаблица(i, j) = rngТаблица(i, j) & " " & iДень
                            End If
                        Next
                    End If
                End If
            Next j
        Loop
        i = i + 1
    Loop
End Sub

Function fnВСтрокеЕстьДвойные(ByVal rngТаблица As Range, ByVal iНомерСтроки As Integer) As Boolean
        For j = iРасстояниеДоДанных To rngТаблица.Columns.Count Step iРасстояниеМеждуКолонками 'Проходим по столбцам таблицы.
            sТекущаяЯчейка = rngТаблица(iНомерСтроки, j) 'Данные из текущей ячейки.
            If IsEmpty(sТекущаяЯчейка) = False And IsNumeric(sТекущаяЯчейка) Then  'Если в ячейке есть число.
                If InStr(sТекущаяЯчейка, " ") > 0 Then 'Если в текущей ячейке несколько чисел разделены пробелами.
                    fnВСтрокеЕстьДвойные = True
                    'MsgBox "Строка:   " & iНомерСтроки & vbLf & "Колонка: " & j & vbLf & vbLf & rngТаблица(iНомерСтроки, j)
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
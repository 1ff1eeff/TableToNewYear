'Public objDicServiceDays As Dictionary 'Дни на которые назначено обслуживание.
Public arrПраздВых() 'Дни праздников и дополнительных выходных.
Public arrРабСуб() 'Рабочие субботы и воскресенья.
Public iНовыйГод, iСтарыйГод, iРасстояниеМеждуКолонками, iРасстояниеДоДанных As Integer

Sub fnАктуализироватьТаблицу()

    '=====================================================================================
    '           Тут необходимо указать актуальные значения!
    '=====================================================================================
    
    iРасстояниеМеждуКолонками = 1 'Расстояние между колонками с датами в таблице.
    iРасстояниеДоДанных = 7
    iСтарыйГод = 2023 'Изначальный год в таблице.
    iНовыйГод = 2024  'Новый год.
    
    'Даты праздничных и дополнительных выходных дней (в числовом виде, в котором хранятся в Excel):
    arrПраздВых = Array(45292, 45293, 45294, 45295, 45296, 45297, 45298, 45299, 45345, 45346, 45347, 45359, 45360, _
                45361, 45410, 45411, 45412, 45413, 45421, 45422, 45423, 45424, 45455, 45599, 45600, 45655, 45656, 45657)
    
    'Даты рабочих суббот или воскресений (в числовом виде, в котором хранятся в Excel):
    arrРабСуб = Array(45409, 45654)
    
    '===================================================================================================================
    
'    If objDicServiceDays Is Nothing Then
'        Set objDicServiceDays = New Dictionary
'    End If
'    objDicServiceDays.RemoveAll
    
    Select Case Selection.Areas.Count 'Рассматриваем количество выделенных областей.
    Case 0: 'Если не выбрана ни одна область.
        MsgBox "Не выбран диапазон ячеек!"
        End
    Case 1: 'Если выбрана одна область.
        Set Table = Selection 'Таблица из выделенной области.
    Case Else 'Если выбраны несколько областей.
        MsgBox "Выделите смежные ячейки!"
        End
    End Select
    
    iКоличествоСтолбцов = Table.Columns.Count
    
    For i = 3 To Table.Rows.Count 'Проходим по строкам таблицы.
    
        For j = iРасстояниеДоДанных To iКоличествоСтолбцов Step iРасстояниеМеждуКолонками 'Проходим по столбцам таблицы с заданным шагом.
            
'            If i = 26 Then
'                MsgBox "STOP"
'            End If
'
            iИсходныйДень = Table(i, j) 'Число из ячейки таблицы.
            If IsEmpty(iИсходныйДень) = False And IsNumeric(iИсходныйДень) Then  'Если в ячейке есть число.
                               
                sИсходныйМесяц = Table(1, j).MergeArea(1) 'Определяем текущий месяц.
                iИсходныйМесяц = Month("01 " & sИсходныйМесяц & " 1900") 'Порядковый номер месяца.
                dИсходныйДень = DateSerial(iСтарыйГод, iИсходныйМесяц, iИсходныйДень)  'Формируем исходную дату.
                dАктуальныйДень = fnАктуализироватьДату(dИсходныйДень) 'Актуализируем.
                
                If j > 1 Then 'Если есть ячейки слева.
                    sПредыдущаяНедля = Table(i, j - iРасстояниеМеждуКолонками) 'Смотрим предыдущую неделю.
                    If Not IsEmpty(sПредыдущаяНедля) Then 'Если в ячейке есть дата.
                        sМесяц = Table(1, j - iРасстояниеМеждуКолонками).MergeArea(1) 'Определяем месяц.
                        sМесяцПредыдущейНедели = Table(1, j - 1).MergeArea(1)
                        iМесяц = Month("01 " & sМесяцПредыдущейНедели & " 1900") 'Порядковый номер месяца.
                        dПредыдущийДень = DateSerial(iНовыйГод, iМесяц, sПредыдущаяНедля) 'Формируем дату.
                        If dПредыдущийДень > dАктуальныйДень Then 'Если дата в текущей ячейке меньше чем в предыдущей.
                            Table(i, j - iРасстояниеМеждуКолонками) = Day(dАктуальныйДень) 'Меняем местами.
                            dАктуальныйДень = dПредыдущийДень
                        End If
                    End If
                End If
                
                Table(i, j) = Day(dАктуальныйДень) 'Добавляем день в текущую ячейку.
                
            End If
        Next j
    Next i
End Sub

Function fnАктуализироватьДату(dИсходныйДень) As Date
   
    'Dim iСдвигДней As Integer 'На сколько дней сдвигать даты.
    'sИсходныйМесяц = MonthName(Month(dИсходныйДень)) 'Определяем месяц исходного дня.
       
    iНомерДняВНеделеИсходного = Weekday(CDate(dИсходныйДень), vbMonday) 'Определяем номер дня в неделе исходного дня.
    lИсходнаяНеделя = CLng(DatePart("ww", dИсходныйДень)) 'Определяем номер недели исходного дня.
    
    iНомерПонедельникаНового = GetDateByWeeknum(ByVal lИсходнаяНеделя, iНовыйГод)
    dНоваяДата = DateAdd("d", iНомерДняВНеделеИсходного - 1, iНомерПонедельникаНового) 'Формируем новую дату.
    iНомерДняВНеделеНового = Weekday(CDate(dНоваяДата), vbMonday)
    iНоваяНеделя = DatePart("ww", dНоваяДата)
        
    dАктуальныйДень = fnСледующийРабочийДень(ByVal DateAdd("d", -1, dНоваяДата))  'Проверяем текущий день на праздники.
    iНомерДняВНеделеАктуального = Weekday(CDate(dАктуальныйДень), vbMonday)
    iАктуальнаяНеделя = DatePart("ww", dАктуальныйДень)
        
    If iАктуальнаяНеделя <> iНоваяНеделя Then 'Если неделя изменилась.
        dАктуальныйДень = fnСледующийРабочийДень(ByVal dНоваяДата, -1) 'Сдвигаем в другом направлении.
        iНомерДняВНеделеАктуального = Weekday(CDate(dАктуальныйДень), vbMonday)
        iАктуальнаяНеделя = DatePart("ww", dАктуальныйДень)
    End If
    
    'MsgBox "Cтарая дата: " & dИсходныйДень & vbLf & "Номер дня в неделe: " & iНомерДняВНеделеИсходного & vbLf & "Номер недели: " & iИсходнаяНеделя & _
        vbLf & vbLf & "Новая дата: " & dНоваяДата & vbLf & "Номер дня в неделe: " & iНомерДняВНеделеНового & vbLf & "Номер недели: " & iНоваяНеделя & _
        vbLf & vbLf & "Актуальный день: " & dАктуальныйДень & vbLf & "Номер дня в неделe: " & iНомерДняВНеделеАктуального & vbLf & "Номер недели: " & iАктуальнаяНеделя
            
    
    'varTemp = objDicServiceDays.Item(CLng(dАктуальныйДень))

    fnАктуализироватьДату = dАктуальныйДень
    
End Function
'---------------------------------------------------------------------------------------
' Purpose: GetDateByWeeknum - получение даты понедельника на основании заданного номера недели
'               lWeeknum    - номер недели, дату понедельника для которой надо получить
'               lYear       - номер года. Если не указан - берется текущий год
'               IsISO       - если True или 1, то первой неделей считается та,
'                             в которой первый раз встречается четверг
'                             Если False, 0 или не указан - берется первая полная неделя
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
    If IsISO Then 'если надо брать первый четверг
        For lr = dStart To dStart + 8
            If VBA.Weekday(lr, vbMonday) = 4 Then
                dStart = lr - 3
                Exit For
            End If
        Next
    Else 'если надо брать не первый четверг, а первую целую неделю
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

Function fnСледующийРабочийДень(ByVal dДень As Date, _
    Optional ByVal iДобавитьДней As Integer = 1) As Date
    If iДобавитьДней > 0 Then
        Do While iДобавитьДней > 0
            If fnЭтоРабочийДень(dДень + 1) Then
                iДобавитьДней = iДобавитьДней - 1
            End If
            dДень = dДень + 1
        Loop
    Else
        Do While iДобавитьДней < 0
            If fnЭтоРабочийДень(dДень - 1) Then
                iДобавитьДней = iДобавитьДней + 1
            End If
            dДень = dДень - 1
        Loop
    End If
    fnСледующийРабочийДень = dДень
End Function

Function fnЭтоРабочийДень(ByVal dДень As Date) As Boolean
    If dДень = 0 Then Exit Function
    If Weekday(dДень, vbMonday) <> 6 And Weekday(dДень, vbMonday) <> 7 _
    And Not fnЭтоПраздникИлиВыходной(dДень) Or fnЭтоРабочаяСуббота(dДень) Then
    
'        If Not fnЭтоДеньОбслуживания(dДень) Then
            fnЭтоРабочийДень = True
'        End If
    Else
        fnЭтоРабочийДень = False
    End If
End Function

'Private Function fnЭтоДеньОбслуживания(dДень As Date) As Boolean
'    bДеньОбслуживания = objDicServiceDays.Exists(CLng(dДень))
'    fnЭтоДеньОбслуживания = bДеньОбслуживания
'End Function


Private Function fnЭтоПраздникИлиВыходной(dДень As Date) As Boolean
    Dim i As Long
    Dim iRow As Integer, varTmp
    Static objDicHolidays As Dictionary
    If objDicHolidays Is Nothing Then
        Set objDicHolidays = New Dictionary
        'Если существует именованный диапазон "НерабочиеДни", добавляем его содержимое в словарь.
        If TypeOf Evaluate("НерабочиеДни") Is Range Then
            With Range("НерабочиеДни")
                For iRow = 1 To .Rows.Count
                    varTmp = objDicHolidays.Item(CLng(.Cells(iRow, 1).Value))
                Next iRow
            End With
        End If

        For i = LBound(arrПраздВых) To UBound(arrПраздВых)
            varTmp = objDicHolidays.Item(arrПраздВых(i))
        Next i
    End If
    fnЭтоПраздникИлиВыходной = objDicHolidays.Exists(CLng(dДень))
End Function

Private Function fnЭтоРабочаяСуббота(dДень As Date) As Boolean
    Dim i As Long
    Dim iRow As Integer, varTmp
    Static objDicWorkingSaturdays As Dictionary
    If objDicWorkingSaturdays Is Nothing Then
        Set objDicWorkingSaturdays = New Dictionary
        'Если существует именованный диапазон "РабочиеСубботы", добавляем его содержимое в словарь.
        If TypeOf Evaluate("РабочиеСубботы") Is Range Then
            With Range("РабочиеСубботы")
                For iRow = 1 To .Rows.Count
                    varTmp = objDicWorkingSaturdays.Item(CLng(.Cells(iRow, 1).Value))
                Next iRow
            End With
        End If
    End If
    For i = LBound(arrРабСуб) To UBound(arrРабСуб)
        varTmp = objDicWorkingSaturdays.Item(arrРабСуб(i))
    Next i
    fnЭтоРабочаяСуббота = objDicWorkingSaturdays.Exists(CLng(dДень))
End Function
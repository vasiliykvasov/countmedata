Sub CountMeData()
' Переменные рядов
Dim FirstRow As Long ' Первый ряд
Dim CurrentRow As Long 'Текущий ряд
Dim LastRow As Long ' Последний ряд
FirstRow = 1
CurrentRow = FirstRow
LastRow = Selection.Rows.Count

' Переменные колонок
Dim FirstColumn As Long ' Первая колонка
Dim CurrentColumn As Long 'Текущая колонка
Dim LastColumn As Long ' Последняя колонка
FirstColumn = 1
CurrentColumn = FirstColumn
LastColumn = Selection.Columns.Count

' Переменные для операций над текущими ячейками вычислений
Dim ActiveRow As Long ' Номер активного ряда
Dim ActiveRangeRow As Range ' Диапазон активного ряда
Dim ActiveRangeColumn As Range ' Диапазон колонки, в которой производятся выявление ячейки

' Переменные для получения и вывода данных
Dim DataRange As String ' Диапазон ячеек для подстановки в формулу
Dim DataColumn As Long ' Числовой номер колонки на листе
Dim DataCell As String ' Какую функцию применять к диапазону
Dim TargetCell As Range   ' Целевая ячейка, куда записываются данные

' Переменные, на случай, если в первой строке нужно поделить число из одного ряда на число из другого, нужно написать буквы рядов через :
Dim LeftCell As String ' Буква слева от двоеточия
Dim RightCell As String ' Буква справа от двоеточия

' Цикл перебора колонок
For CurrentColumn = FirstColumn To LastColumn
Debug.Print "Текущий ряд - " & CurrentColumn
Set ActiveRangeColumn = Range(Selection.Cells(FirstRow, CurrentColumn), Selection.Cells(LastRow, CurrentColumn)) ' Получаем диапазон активной колонки
Debug.Print "Диапазон активной колонки = " & ActiveRangeColumn.Address

Debug.Print " "

DataColumn = Range(Selection.Cells(FirstRow, CurrentColumn), Selection.Cells(LastRow, CurrentColumn)).Column 'Получаем номер текущей колонки
DataCell = Cells(1, DataColumn) ' Получаем значение ячейки с функцией на английском языке из первой ячейки колонки
Set TargetCell = Selection.Cells(LastRow + 1, CurrentColumn)
Debug.Print "TargetCell = " & TargetCell.Address


' Если в значении ячейки, где должна храниться функция - :
If InStr(DataCell, ":") > 0 Then
Debug.Print "Двоеточие"
LeftCell = "$" & Left(DataCell, 1) & "$" & TargetCell.Row ' Получаем адрес ячейки-делимого - слева от двоеточия
RightCell = "$" & Right(DataCell, 1) & "$" & TargetCell.Row  ' Получаем адрес ячейки-делителя - справа от двоеточия
TargetCell = "=" & LeftCell & "/" & RightCell ' Применяем деление под активным диапазоном

' Если в значении ячейки, где должна храниться функция - ОЧИСТИТЬ
ElseIf DataCell = "ОЧИСТИТЬ" Then
Debug.Print "Очистить"
TargetCell = "" ' Заполняем ячейку под активным диапазоном пустотой

' Если значение ячейки, где должна храниться функция - СУММ
ElseIf DataCell = "СУММ" Then
Debug.Print "СУММ"

DataRange = RowIteration(CurrentColumn)
TargetCell.FormulaLocal = "=СУММ(" & DataRange & ")" ' Применяем формулу СУММ для ранее записанного диапазона ячеек

' Если значение ячейки, где должна храниться функция - СРЗНАЧ
ElseIf DataCell = "СРЗНАЧ" Then
Debug.Print "СРЗНАЧ"
DataRange = RowIteration(CurrentColumn)
TargetCell.FormulaLocal = "=СРЗНАЧ(" & DataRange & ")" ' ѕрименяем формулу СРЗНАЧ для ранее записанного диапазона ячеек

Else
Debug.Print "не работает"
End If

'DataRange = "" ' Очищаем диапазон ячеек для подстановки в формулу

Next CurrentColumn ' Конец цикла перебора текущей колонки

End Sub

Function RowIteration(CurrentColumn As Long) As String
' Переменные рядов
Dim FirstRow As Long ' Первый ряд
Dim CurrentRow As Long ' Текущий ряд
Dim LastRow As Long ' Последний ряд
FirstRow = 1
CurrentRow = FirstRow
LastRow = Selection.Rows.Count
Dim ActiveCell As Range ' Активная ячейка

' Цикл перебора рядов
For CurrentRow = FirstRow To LastRow ' Цикл проходит текущим рядом от первого ряда до последнего ряда
Set ActiveCell = Selection.Cells(CurrentRow, CurrentColumn) ' Получаем активную ячейку
Debug.Print "Активная ячейка - " & ActiveCell.Address
ActiveRow = ActiveCell.Row ' ѕолучаем активный ряд
Debug.Print "Активный ряд - " & ActiveRow

' Проверяем вложенность активного ряда
If Rows(ActiveRow).OutlineLevel = 1 Then
Debug.Print "не сгруппировано"
RowIteration = RowIteration + ActiveCell.Address + ";" ' Если ряд не сгруппирован, дописываем его в диапазон ячеек для подстановки в формулу
Debug.Print "RowIteration = "; RowIteration; ""
End If

Next CurrentRow ' Конец цикла перебора рядов
End Function

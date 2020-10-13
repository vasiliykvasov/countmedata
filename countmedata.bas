Sub CountMeData()

' ���������� �����
Dim FirstRow As Long ' ������ ���
Dim CurrentRow As Long '������� ���
Dim LastRow As Long '��������� ���
FirstRow = 1
CurrentRow = FirstRow
LastRow = Selection.Rows.Count

' ���������� �������
Dim FirstColumn As Long ' ������ �������
Dim CurrentColumn As Long '������� �������
Dim LastColumn As Long '��������� �������
FirstColumn = 1
CurrentColumn = FirstColumn
LastColumn = Selection.Columns.Count

' ���������� ��� �������� ��� �������� �������� ����������
Dim ActiveRow As Long '����� ��������� ����
Dim ActiveRangeRow As Range '�������� ��������� ����
Dim ActiveRangeColumn As Range '�������� �������, � ������� ������������ ��������� ������
Dim ActiveCell As Range '

' ���������� ��� ��������� � ������ ������
Dim DataRange As String '�������� ����� ��� ����������� � �������
Dim DataColumn As Long '�������� ����� ������� �� �����
Dim DataCell As String '������� �� ���������� �����

' ����������, �� ������, ���� � ������ ������ ����� �������� ����� �� ������ ���� �� ����� �� �������, ����� �������� ����� ����� ����� :
Dim LeftCell As String '����� ����� �� ���������
Dim RightCell As String '����� ������ �� ���������

' ���� �������� �������
For CurrentColumn = FirstColumn To LastColumn
Debug.Print "������� ��� - " & CurrentColumn
Set ActiveRangeColumn = Range(Selection.Cells(FirstRow, CurrentColumn), Selection.Cells(LastRow, CurrentColumn)) ' �������� �������� �������� �������
Debug.Print "�������� �������� ������� = " & ActiveRangeColumn.Address

Debug.Print " "

'���� �������� �����
For CurrentRow = FirstRow To LastRow
Set ActiveCell = Selection.Cells(CurrentRow, CurrentColumn) ' �������� �������� ������
Debug.Print "�������� ������ - " & ActiveCell.Address
ActiveRow = ActiveCell.Row ' �������� �������� ���
Debug.Print "�������� ��� - " & ActiveRow

' ��������� ����������� ��������� ����
If Rows(ActiveRow).OutlineLevel = 1 Then
Debug.Print "�� �������������"
DataRange = DataRange + ActiveCell.Address + ";" ' ���� ��� �� ������������, ���������� ��� � �������� ����� ��� ����������� � �������
Debug.Print DataRange
End If

Next CurrentRow ' ����� ����� �������� �����

Debug.Print " "

DataColumn = Range(Selection.Cells(FirstRow, CurrentColumn), Selection.Cells(LastRow, CurrentColumn)).Column '�������� ����� ������� �������
DataCell = Cells(1, DataColumn) ' �������� �������� ������ � �������� �� ���������� ����� �� ������ ������ �������

' ���� � �������� ������, ��� ������ ��������� ������� - :
If InStr(DataCell, ":") > 0 Then
Debug.Print "���������"
LeftCell = "$" & Left(DataCell, 1) & "$" & (ActiveRow + 1) '�������� ����� ������-�������� - ����� �� ���������
RightCell = "$" & Right(DataCell, 1) & "$" & (ActiveRow + 1) '�������� ����� ������-�������� - ������ �� ���������
Selection.Cells(LastRow + 1, CurrentColumn) = "=" & LeftCell & "/" & RightCell '��������� ������� ��� �������� ����������

' ���� � �������� ������, ��� ������ ��������� ������� - ��������
ElseIf DataCell = "��������" Then
Debug.Print "��������"
Selection.Cells(LastRow + 1, CurrentColumn) = "" '��������� ������ ��� �������� ���������� ��������

' ���� �������� ������, ��� ������ ��������� ������� - ����
ElseIf DataCell = "����" Then
Debug.Print "����"
Selection.Cells(LastRow + 1, CurrentColumn).FormulaLocal = "=����(" & DataRange & ")" '��������� ������� ���� ��� ����� ����������� ��������� �����

' ���� �������� ������, ��� ������ ��������� ������� - ������
ElseIf DataCell = "������" Then
Debug.Print "�� �������"
Selection.Cells(LastRow + 1, CurrentColumn).FormulaLocal = "=������(" & DataRange & ")" ' ��������� ������� ������ ��� ����� ����������� ��������� �����

Else
Debug.Print "�� ��������"
End If

DataRange = "" '������� �������� ����� ��� ����������� � �������

Next CurrentColumn '����� ����� �������� ������� �������

End Sub

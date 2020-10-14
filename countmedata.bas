Sub CountMeData()
'  
Dim FirstRow As Long '  
Dim CurrentRow As Long ' 
Dim LastRow As Long '  
FirstRow = 1
CurrentRow = FirstRow
LastRow = Selection.Rows.Count

'  
Dim FirstColumn As Long '  
Dim CurrentColumn As Long ' 
Dim LastColumn As Long '  
FirstColumn = 1
CurrentColumn = FirstColumn
LastColumn = Selection.Columns.Count

'       
Dim ActiveRow As Long '   
Dim ActiveRangeRow As Range '   
Dim ActiveRangeColumn As Range '  ,     

'      
Dim DataRange As String '      
Dim DataColumn As Long '     
Dim DataCell As String '     
Dim TargetCell As Range   '  ,   

' ,  ,              ,      :
Dim LeftCell As String '    
Dim RightCell As String '    

'   
For CurrentColumn = FirstColumn To LastColumn
Debug.Print "  - " & CurrentColumn
Set ActiveRangeColumn = Range(Selection.Cells(FirstRow, CurrentColumn), Selection.Cells(LastRow, CurrentColumn)) '    
Debug.Print "   = " & ActiveRangeColumn.Address

Debug.Print " "

DataColumn = Range(Selection.Cells(FirstRow, CurrentColumn), Selection.Cells(LastRow, CurrentColumn)).Column '   
DataCell = Cells(1, DataColumn) '            
Set TargetCell = Selection.Cells(LastRow + 1, CurrentColumn)
Debug.Print "TargetCell = " & TargetCell.Address


'    ,     - :
If InStr(DataCell, ":") > 0 Then
Debug.Print ""
LeftCell = "$" & Left(DataCell, 1) & "$" & TargetCell.Row '   - -   
RightCell = "$" & Right(DataCell, 1) & "$" & TargetCell.Row  '   - -   
TargetCell = "=" & LeftCell & "/" & RightCell '     

'    ,     - 
ElseIf DataCell = "" Then
Debug.Print ""
TargetCell = "" '      

'   ,     - 
ElseIf DataCell = "" Then
Debug.Print ""

DataRange = RowIteration(CurrentColumn)
TargetCell.FormulaLocal = "=(" & DataRange & ")" '        

'   ,     - 
ElseIf DataCell = "" Then
Debug.Print ""
DataRange = RowIteration(CurrentColumn)
TargetCell.FormulaLocal = "=(" & DataRange & ")" '        

Else
Debug.Print " "
End If

'DataRange = "" '       

Next CurrentColumn '     

End Sub

Function RowIteration(CurrentColumn As Long) As String
'  
Dim FirstRow As Long '  
Dim CurrentRow As Long '  
Dim LastRow As Long '  
FirstRow = 1
CurrentRow = FirstRow
LastRow = Selection.Rows.Count
Dim ActiveCell As Range '  

'   
For CurrentRow = FirstRow To LastRow '          
Set ActiveCell = Selection.Cells(CurrentRow, CurrentColumn) '   
Debug.Print "  - " & ActiveCell.Address
ActiveRow = ActiveCell.Row '   
Debug.Print "  - " & ActiveRow

'    
If Rows(ActiveRow).OutlineLevel = 1 Then
Debug.Print " "
RowIteration = RowIteration + ActiveCell.Address + ";" '    ,         
Debug.Print "RowIteration = "; RowIteration; ""
End If

Next CurrentRow '    
End Function

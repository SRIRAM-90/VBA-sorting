Attribute VB_Name = "Module1"
Public Sub DivisionSort()
Columns("A:F").sort key1:=Range("A2"), order1:=xlDescending, Header:=xlYes
End Sub
Public Sub CategorySort()
Columns("A:F").sort key1:=Range("B2"), order1:=xlDescending, Header:=xlYes
End Sub
Public Sub TotalSort()
Columns("A:F").sort key1:=Range("F2"), order1:=xlDescending, Header:=xlYes
End Sub

Public Sub makeIt()
Dim sort As Integer
Dim promptMsg As String
 Dim error As Integer
 On Error GoTo errHandler
promptMsg = "what do you want to sort?" & vbCrLf & _
"1 - sort by Division" & vbCrLf & _
"2 - sort by Category" & vbCrLf & _
"3 - sort by Total"
sort = InputBox(promptMsg, "Sort order")

If sort = 1 Then
DivisionSort
ElseIf sort = 2 Then
CategorySort
ElseIf sort = 3 Then
TotalSort
  Else
errHandler:
 error = MsgBox("Invalid try again!", vbYesNo)
 If error = 6 Then
 makeIt
 End If
End If
End Sub

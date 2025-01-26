Function MyFunction(param1, param2)
  On Error Resume Next
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 9999, , "Parameters cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of function code ...
End Function

' Example of how to handle the error in a calling function:
Sub CallMyFunction()
  On Error GoTo ErrHandler
  MyFunction 
  ' If execution reaches here, no errors occurred.
  Exit Sub

ErrHandler:
  If Err.Number = 9999 Then
    MsgBox "Error: " & Err.Description
  Else
    MsgBox "An unexpected error occurred: " & Err.Number & " - " & Err.Description
  End If
End Sub
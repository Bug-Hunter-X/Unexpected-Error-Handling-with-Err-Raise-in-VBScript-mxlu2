Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 9999, , "Parameters cannot be empty"
  End If
  ' ... rest of function code ...
End Function
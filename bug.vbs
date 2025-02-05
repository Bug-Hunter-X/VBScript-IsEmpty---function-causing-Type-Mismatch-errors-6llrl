Function MyFunc(param)
  If IsEmpty(param) Then
    Err.Raise 13, , "Type mismatch"
  End If
  ' ... rest of function code ...
End Function
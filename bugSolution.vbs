Function MyFunc(param)
  If VarType(param) = vbEmpty Then
    ' Handle empty Variant
    ' ... appropriate action ...
  ElseIf VarType(param) <> vbString Then
    Err.Raise 13, , "Parameter must be a string"
  ElseIf Len(Trim(param)) = 0 Then
    ' Handle empty string
    ' ... appropriate action ...
  Else
    ' ... rest of function code ...
  End If
End Function
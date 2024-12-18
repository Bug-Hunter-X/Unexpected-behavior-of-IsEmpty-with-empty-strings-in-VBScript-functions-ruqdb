Function MyFunction(param1, param2)
  If param1 = "" Or param2 = "" Then
    Err.Raise 13, , "Both parameters are required"
  ElseIf IsMissing(param1) Or IsMissing(param2) Then
    Err.Raise 13, , "Both parameters are required"
  End If
  ' ... rest of the function ...
End Function
Function f(a, b)
  If Not IsArray(a) And IsEmpty(a) Or Not IsArray(b) And IsEmpty(b) Then 
    Err.Raise 13, , "Type mismatch"
  ElseIf IsArray(a) And UBound(a) = -1 Or IsArray(b) And UBound(b) = -1 Then
    ' Handle empty arrays 
    ' ... appropriate action for empty arrays 
  Else
    ' ...rest of function
  End If
End Function
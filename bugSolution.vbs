Function f(a, b)
  If IsEmpty(a) Or a = vbEmpty Then
    a = 0
  End If
  If IsEmpty(b) Or b = vbEmpty Then
    b = 0
  End If
  f = a + b
End Function

MsgBox f(1, Empty) 'Output: 1
MsgBox f(Empty, 2) 'Output: 2
MsgBox f(Empty, Empty) 'Output: 0
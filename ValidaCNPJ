Public Function FU_ValidaCGC(cgc As String) As Boolean
  Dim retorno, a, J, I, d1, d2
  If Len(cgc) = 8 And Val(cgc) > 0 Then
    a = 0
    J = 0
    d1 = 0
    For I = 1 To 7
      a = Val(Mid(cgc, I, 1))
      If (I Mod 2) <> 0 Then
        a = a * 2
      End If
      If a > 9 Then
        J = J + Int(a / 10) + (a Mod 10)
      Else
        J = J + a
      End If
    Next I
    d1 = IIf((J Mod 10) <> 0, 10 - (J Mod 10), 0)
    If d1 = Val(Mid(cgc, 8, 1)) Then
      FU_ValidaCGC = True
    Else
      FU_ValidaCGC = False
    End If
  Else
    If Len(cgc) = 14 And Val(cgc) > 0 Then
      a = 0
      I = 0
      d1 = 0
      d2 = 0
      J = 5
      For I = 1 To 12 Step 1
        a = a + (Val(Mid(cgc, I, 1)) * J)
        J = IIf(J > 2, J - 1, 9)
      Next I
      a = a Mod 11
      d1 = IIf(a > 1, 11 - a, 0)
      a = 0
      I = 0
      J = 6
      For I = 1 To 13 Step 1
        a = a + (Val(Mid(cgc, I, 1)) * J)
        J = IIf(J > 2, J - 1, 9)
      Next I
      a = a Mod 11
      d2 = IIf(a > 1, 11 - a, 0)
      If (d1 = Val(Mid(cgc, 13, 1)) And d2 = Val(Mid(cgc, _
            14, 1))) Then
        FU_ValidaCGC = True
      Else
        FU_ValidaCGC = False
      End If
    Else
      FU_ValidaCGC = False
    End If
  End If
End Function

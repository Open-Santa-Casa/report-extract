Function getData(mesReferencia)
  If mesReferencia = 1 Then
    mesReferencia = 12
    shell.SendKeys"01"& mesReferencia & Year(Date())
    shell.SendKeys"30"& mesReferencia & Year(Date())
  ElseIf mesReferencia <= 10 Then
    mesReferencia = mesReferencia - 1
    shell.SendKeys"010"& mesReferencia & Year(Date())
    shell.SendKeys"300"& mesReferencia & Year(Date())
  Else 
    mesReferencia = mesReferencia - 1
    shell.SendKeys"01"& mesReferencia & Year(Date())
    shell.SendKeys"30"& mesReferencia & Year(Date())
  End If
End Function
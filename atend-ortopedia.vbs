Dim mesReferencia
' Abre o programa 💻
set shell = CreateObject("wscript.shell")

shell.run """C:\smart\Aplic125\AUDIT125.exe"""
WScript.sleep 3000

' Realiza o login 🔓
shell.SendKeys"FLAVIOE"
shell.SendKeys "{ENTER}"
shell.SendKeys"Shomertora*37"
shell.SendKeys"{ENTER}"
WScript.sleep 1000
' Seleciona a Base de Dados *HDPA | Unacon
shell.SendKeys"{TAB}"
WScript.sleep 200
shell.SendKeys"{ENTER}"
' Recusa ler mensagens
WScript.sleep 100
shell.SendKeys"{TAB}"
WScript.sleep 100
shell.SendKeys"{ENTER}"
WScript.sleep 3000
' Abrindo o emissor
shell.SendKeys"%"
WScript.sleep 55
shell.SendKeys"u"
WScript.sleep 55
shell.SendKeys"e"
WScript.sleep 100

' Hora de caçar o Relatório -- "Relatório de atendimento de Ortopedia" 🏹
shell.SendKeys"Relat"
shell.SendKeys"{TAB}"
shell.SendKeys"{TAB}"
shell.SendKeys"{TAB}"
shell.SendKeys"{TAB}"
shell.SendKeys"{TAB}"
shell.SendKeys"{DOWN}"
shell.SendKeys"{ENTER}"

mesReferencia = Month(Date())
WScript.sleep 100
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
shell.SendKeys "{F4}"
WScript.sleep 100
shell.SendKeys "{F6}"
shell.SendKeys "{TAB}"
shell.SendKeys "{DOWN}"
shell.SendKeys "{ENTER}"
WScript.sleep 20
shell.SendKeys "Atendimento Ortopedia - "& mesReferencia & Year(Date())
shell.SendKeys "{ENTER}"


shell.SendKeys "%{F4}"
shell.SendKeys "%{F4}"
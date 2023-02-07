Dim mesReferencia
' Abre o programa 💻
set shell = CreateObject("wscript.shell")
sub ()
  
end sub
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

' Hora de caçar o Relatório -- "ADM - Cirurgia por Especialidade 2" 🏹
shell.SendKeys"ADM - Cir"
WScript.sleep 10
shell.SendKeys"{ENTER}"
WScript.sleep 100
shell.SendKeys"{TAB}"
WScript.sleep 10
shell.SendKeys"{TAB}"
WScript.sleep 10
shell.SendKeys"{TAB}"
mesReferencia = Month(Date())
WScript.sleep 100

' Captura a a data 
set dat = shell.Run "C:\Users\flavi\Saved Games\py-excel-2-json\tools-function\getData.vbs"

shell.SendKeys "{F4}"
WScript.sleep 100
shell.SendKeys "{F6}"
WScript.sleep 13
shell.SendKeys "{TAB}"
shell.SendKeys "{DOWN}"
shell.SendKeys "{ENTER}"
WScript.sleep 80
shell.SendKeys "{ENTER}"
shell.SendKeys "Cirurgia Por Especialidade - "& mesReferencia & Year(Date())
shell.SendKeys "{ENTER}"


'MsgBox mesReferencia, vbOKOnly + vbInformation, "Aviso!"
WScript.sleep 80
MsgBox "Script em execução", vbOKOnly + vbInformation, "Aviso!"

shell.SendKeys "%{F4}"
shell.SendKeys "%{F4}"






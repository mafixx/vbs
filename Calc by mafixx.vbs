Dim mult, a, b, answer

a = InputBox("Digite um n�mero!","Calculadora by mafixx")
b = InputBox("Digite um segundo n�mero!","Calculadora by mafixx")

If Not (a = "" And b = "") Then
	MsgBox "Voc� n�o digitou em um dos campos.", vbCritical, "Calculadora by mafixx"
	MsgBox "Adeus.", vbInformation, "Calculadora by mafixx"
Else
End if


  mult = (a*b)
  MsgBox "O produto entre " & a & " e " & b & " � "

MsgBox mult

answer = MsgBox("O resultado est� correto?", vbQuestion + vbYesNo + vbDefaultButton2, "Calculadora by mafixx")

If answer = vbYes Then
	MsgBox "Ok!", vbInformation, "Calculadora by mafixx"
elseif answer = vbNo Then
	MsgBox "Voc� n�o est� sabendo o que quer da vida.", vbExclamation ,"Calculadora by mafixx"
else
	MsgBox "Tente de novo."
End if
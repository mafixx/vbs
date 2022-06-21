Dim mult, a, b, answer

a = InputBox("Digite um número!","Calculadora by mafixx")
b = InputBox("Digite um segundo número!","Calculadora by mafixx")

If Not (a = "" And b = "") Then
	MsgBox "Você não digitou em um dos campos.", vbCritical, "Calculadora by mafixx"
	MsgBox "Adeus.", vbInformation, "Calculadora by mafixx"
Else
End if


  mult = (a*b)
  MsgBox "O produto entre " & a & " e " & b & " é "

MsgBox mult

answer = MsgBox("O resultado está correto?", vbQuestion + vbYesNo + vbDefaultButton2, "Calculadora by mafixx")

If answer = vbYes Then
	MsgBox "Ok!", vbInformation, "Calculadora by mafixx"
elseif answer = vbNo Then
	MsgBox "Você não está sabendo o que quer da vida.", vbExclamation ,"Calculadora by mafixx"
else
	MsgBox "Tente de novo."
End if
Dim input
    input = InputBox("Digite seu nome")
    MsgBox "Seu nome e " & input
    'Dim é a variável que eu criei, e input é seu nome, depois eu chamei a variável e coloquei um inputbox nela, após isso chamei a funcção msgbox e exibi o que o client escreveu no input'

Dim input2
input2 = InputBox("what is your name?")
if input2 = "" Then
    MsgBox "please, write something!"
    else
        MsgBox "your name is " &input2
end if 
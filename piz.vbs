MsgBox "a palavra ta longe" &Space(15)& "to mesmo"
'essa função adciona 15 backspaces entre uma string e outra'

MsgBox String(5, "opa") &String(8, vbLf)& "a letra o apareceu 5 vezes", vbYesNo
'essa função cria uma string, e voce pode escolher quantas vezes ela irá se repetir'

MsgBox StrReverse("1234567")
'essa é autoexplicativa'

MsgBox Len("Matheus Alves")
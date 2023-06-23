set shell = CreateObject(WScript.Shell)
pesquisa = InputBox("o que voce quer saber?", "navegador")  
if pesquisa = "" Then
    MsgBox "digite algo para que a pesquisa seja feita"
Else
    shell.run "cmd / c start https://www.bing.com/search?q=" &pesquisa
end if  
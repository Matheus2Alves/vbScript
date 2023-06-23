set pesq = CreateObject(wscript)
pesquisa = InputBox("o que voce quer saber?", "navegador")  
if pesquisa = "" Then
    MsgBox "digite algo para que a pesquisa seja feita"
    Else
        pesq.run "https://www.bing.com/search?q=" &pesquisa& ""
end if
msgBox "olá mundo", 16, "calabresa"
'msgBox é a função, a primeira string é o texto dentro da msgBox, o primeiro número são os buttons que indicam o que será exibido no msgBox, como: ícone, modalidade e os botões.. , e por último temos o título da msgBox.'

msgBox "você quer pizza?", 4 or 64 or 0 or 16384, "calabresa" 
'o 4 representa os botões exibidos, o 64 representa o ícone ilustrador, o 0 representa que o primeiro botão é o padrão, e por último o 16384 que representa a adição do botão "ajuda" na msgBox'

msgBox "você quer pizza?", 4 or 64 or 0 , "calabresa" 

'voce tambem pode usar os números de forma escrita, desta maneira a seguir:'
msgBox "você tem certeza que deseja excluir o arquivo?", vbOkCancel or vbCritical or vbDefaultButton2 or vbSystemModal, "IMPORTANTE"
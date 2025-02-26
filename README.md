# Suporte para diversos formatos de imagem no Visual Basic 6 (VB6)
Formatos suportados BMP,ICO,CUR,WMF,EMF,JPG,PNG,TIF,GIF
Como utilizar:
1 - Importar o Class Module "StdPictureEx.cls" no projeto
2 - Adicionar um componente "Image" no formulário
3 - Instanciar a classe no formulário, exemplo:
  Public Sub LoadPNG(ByRef P_ComponenteImagem As Image, P_CaminhoImagem As String)
    Dim StdPictureExInstance As New StdPictureEx  
    Set P_ComponenteImagem.Picture = StdPictureExInstance.LoadPicture(P_CaminhoImagem)
  End Sub
4 - Chamar a Sub LoadPNG no Form_Load ou em outro evento desejado, passando os devidos parâmetros, exemplo:
  Private Sub Form_Load()
   LoadPNG imgTeste, "C:\Projects\VB6\SuportePNG\vb6.png"
   LoadPNG imgTeste2, "C:\Projects\VB6\SuportePNG\vb6.png"
  End Sub

Vídeo tutorial: https://youtu.be/uqtHHWsdHQE

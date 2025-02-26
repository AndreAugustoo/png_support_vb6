# Suporte para Diversos Formatos de Imagem no Visual Basic 6 (VB6)

Este projeto permite o **suporte a múltiplos formatos de imagem** no Visual Basic 6 (VB6). Ele inclui os formatos mais comuns, como **BMP, ICO, CUR, WMF, EMF, JPG, PNG, TIF e GIF**.

## Formatos Suportados
- BMP
- ICO
- CUR
- WMF
- EMF
- JPG
- PNG
- TIF
- GIF

## Como Utilizar

### 1. Importar o Class Module "StdPictureEx.cls" no Projeto
Baixe o arquivo **`StdPictureEx.cls`** e importe-o para o seu projeto VB6.

### 2. Adicionar um Componente "Image" no Formulário
Adicione um controle **`Image`** no formulário onde você deseja exibir as imagens.

### 3. Instanciar a Classe no Formulário

Crie uma Sub para carregar a imagem no componente `Image`. Exemplo:

```vb
Public Sub LoadPNG(ByRef P_ComponenteImagem As Image, P_CaminhoImagem As String)
    Dim StdPictureExInstance As New StdPictureEx
    Set P_ComponenteImagem.Picture = StdPictureExInstance.LoadPicture(P_CaminhoImagem)
End Sub

```vb
Private Sub Form_Load()
    LoadPNG imgTeste, "C:\Projects\VB6\SuportePNG\vb6.png"
    LoadPNG imgTeste2, "C:\Projects\VB6\SuportePNG\vb6.png"
End Sub

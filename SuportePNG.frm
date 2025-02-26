VERSION 5.00
Begin VB.Form SuportePNG 
   BackColor       =   &H80000007&
   Caption         =   "Suporte PNG - VB6"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgTeste 
      Height          =   2895
      Left            =   1680
      Top             =   1200
      Width           =   4695
   End
End
Attribute VB_Name = "SuportePNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadPNG(ByRef P_ComponenteImagem As Image, P_CaminhoImagem As String)
    Dim StdPictureExInstance As New StdPictureEx
    
    Set P_ComponenteImagem.Picture = StdPictureExInstance.LoadPicture(P_CaminhoImagem)
End Sub


Private Sub Form_Load()

   LoadPNG imgTeste, "C:\Projects\VB6\SuportePNG\vb6.png"

End Sub

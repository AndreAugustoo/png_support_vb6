VERSION 5.00
Begin VB.Form SuportePNG 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Suporte PNG - VB6"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Image imgTeste 
         Height          =   1695
         Left            =   0
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Image imgTeste2 
      Height          =   1095
      Left            =   0
      Top             =   480
      Width           =   1575
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
   LoadPNG imgTeste2, "C:\Projects\VB6\SuportePNG\vb6.png"


End Sub


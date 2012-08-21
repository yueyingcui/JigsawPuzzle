VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6105
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox CheckPicture 
      Height          =   9135
      Left            =   0
      ScaleHeight     =   9075
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CheckPicture.Picture = LoadPicture(UserSelectFile)
CheckPicture.AutoSize = True
End Sub

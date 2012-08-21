VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Time Counter"
   ClientHeight    =   630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2595
   LinkTopic       =   "Form4"
   ScaleHeight     =   630
   ScaleWidth      =   2595
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Hurry Up! "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

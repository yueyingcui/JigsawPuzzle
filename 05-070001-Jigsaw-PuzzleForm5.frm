VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "英雄榜"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2490
   LinkTopic       =   "Form5"
   ScaleHeight     =   2760
   ScaleWidth      =   2490
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem InputBox("请问尊姓大名：")
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "数学071班 070001崔越莹"
   ClientHeight    =   9600
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   14340
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   960
      Top             =   120
   End
   Begin VB.PictureBox OriPicture 
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   240
         Top             =   1080
      End
      Begin VB.Timer Timer2 
         Interval        =   50000
         Left            =   120
         Top             =   120
      End
      Begin VB.PictureBox SplitImage 
         Height          =   735
         Index           =   0
         Left            =   1200
         ScaleHeight     =   675
         ScaleWidth      =   555
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Menu menu 
      Caption         =   "拼图游戏"
      Begin VB.Menu menupic 
         Caption         =   "Select Your Picture!"
      End
      Begin VB.Menu menudegree 
         Caption         =   "Game Degree"
      End
      Begin VB.Menu menubegin 
         Caption         =   "Game Begin"
      End
      Begin VB.Menu menuline 
         Caption         =   "-"
      End
      Begin VB.Menu menuover 
         Caption         =   "Game Over"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "帮助"
      Begin VB.Menu menucheck 
         Caption         =   "Check OriPicture"
      End
   End
   Begin VB.Menu menumusic 
      Caption         =   "音乐"
      Begin VB.Menu menumidplay 
         Caption         =   "Play Music"
      End
   End
   Begin VB.Menu menuhistory 
      Caption         =   "英雄榜"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'装载图片可以使用CommandDialog对话框来进行，Filter属性设置为只允许装载Jpg,Gif,BMP类型的文件即可，当文件打开后，利用PictureBox的AutoSize特性可以得到真实大小，此时再调节窗体大小即可实现自适应大小。
Private Sub menupic_Click() '装载图片函数8
CommonDialog1.InitDir = "c:\windows" '指定初始文件目录P210
CommonDialog1.Filter = "Jpg File(*.jpg)|*.jpg|Gif File(*.gif)|*.gif|Bmp File(*.bmp)|*.bmp" '过滤器
CommonDialog1.FilterIndex = 1 '指定默认过滤器
On Error GoTo UserCancle
CommonDialog1.ShowOpen '显示"打开"对话框
UserSelectFile = CommonDialog1.FileName 'UserSelectFile获得选定的文件名
OriPicture.Picture = LoadPicture(UserSelectFile) '打开图片文件并显示
UserCancle: '出现错误退出过程
OriPicture.AutoSize = True 'PictureBox适应图片大小
End Sub

Private Sub menudegree_Click()
ControlCount = InputBox("请输入格数：", 难度设置, 4)
End Sub

Private Sub menubegin_Click()
Call splitpicture(ControlCount)
Call changepic
Timer2.Enabled = True
Timer3.Enabled = True
Form4.Visible = True
End Sub

'切割图片函数
Private Sub splitpicture(ControlCount As Integer)
'当用户开始游戏后，可以利用PictureBox的控件数组来存储切片图像，动态数组的大小可以根据切片数量来决定，至于切片上的图像，则可以用PaintPicture方法来从整体图像上获取。
Dim i As Integer, j As Integer, k As Integer
Dim splitcount As Integer
ReDim lOldSplitImage(ControlCount) As Integer
ReDim tOldSplitImage(ControlCount) As Integer
For i = 1 To ControlCount - 1
    Load SplitImage(i)
    SplitImage(i).Visible = True
Next
splitcount = Sqr(ControlCount)


SplitImageWidth = OriPicture.Width / splitcount '每格宽度值
SplitImageHeight = OriPicture.Height / splitcount  '每格高度值
For i = 0 To (splitcount - 1) '分解图片并按照顺序放入网格
    For j = 0 To (splitcount - 1)
        k = i * splitcount + j
        SplitImage(k).Width = SplitImageWidth '设置宽度P226 P166
        SplitImage(k).Height = SplitImageHeight '设置高度
        SplitImage(k).AutoRedraw = True
        SplitImage(k).PaintPicture OriPicture.Picture, 0, 0, OriPicture.Width / splitcount, OriPicture.Height / splitcount, j * SplitImageWidth, i * SplitImageHeight, OriPicture.Width / splitcount, OriPicture.Height / splitcount, vbSrcCopy
         'SplitImageWidth和SplitImageHeight是每个切片宽度和高度
         'Paintpicture方法裁剪图像[对象名.paintpicture 图像源,绘图起点坐标x,y,宽度,高度,裁剪起点坐标x,y,宽度,高度,绘制方式]P240
        SplitImage(k).AutoRedraw = True
        SplitImage(k).Top = i * SplitImageHeight '设置左上角坐标P226 j-x,i-y
        SplitImage(k).Left = j * SplitImageWidth
        SplitImage(k).Visible = True '控件可视P167
        SplitImage(k).ZOrder
    Next j
Next i

For i = 0 To ControlCount - 1
   lOldSplitImage(i) = SplitImage(i).Left
   tOldSplitImage(i) = SplitImage(i).Top
Next
End Sub

'重组图片函数
Private Sub changepic()
Dim i As Integer, j As Integer, k As Integer
Dim t As Integer, l As Integer, temp As Integer

'OriPicture.Visible = False '隐藏原图片

For i = 0 To ControlCount - 1 '打乱顺序并放入网格
    
    Randomize '随机选择两个网格
     j = Int(Rnd() * ControlCount)
    Randomize
     k = Int(Rnd() * ControlCount)
    
    t = SplitImage(j).Top '互换网格中图片
    l = SplitImage(j).Left
    SplitImage(j).Top = SplitImage(k).Top
    SplitImage(j).Left = SplitImage(k).Left
    SplitImage(k).Top = t
    SplitImage(k).Left = l
    'temp = j '更新图片所在位置
    'j = k
    'j = temp
Next i
End Sub



'用鼠标交换图片
Private Sub SplitImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t As Integer, l As Integer

If sign = 0 Then
   num1 = Index
   sign = 1
ElseIf sign = 1 Then
   num2 = Index
   sign = 0
   
t = SplitImage(num1).Top '互换网格中图片
l = SplitImage(num1).Left
SplitImage(num1).Top = SplitImage(num2).Top
SplitImage(num1).Left = SplitImage(num2).Left
SplitImage(num2).Top = t
SplitImage(num2).Left = l

End If


Call WINER(ControlCount)
End Sub

Private Sub menucheck_Click()
If n = 5 Then
  Exit Sub
End If
Form2.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer() '控制察看原图时间
Form2.Visible = False
n = n + 1
End Sub

Private Sub WINER(ControlCount As Integer)

For i = 0 To ControlCount - 1
  '通过PIC各数组控件位置数据的比较做出判断
  If SplitImage(i).Left <> lOldSplitImage(i) Or SplitImage(i).Top <> tOldSplitImage(i) Then
    Exit Sub
  End If
Next i
MsgBox "You Win!"
Timer3.Enabled = False
Form5.Visible = True

End Sub

Private Sub menumidplay_Click()
Form3.Visible = True
End Sub

Private Sub menuover_Click()
End
End Sub

Private Sub Timer2_Timer() '控制游戏总时间
 MsgBox "You Lose!"
 End
End Sub

Private Sub Timer3_Timer() '控制计时器显示Form4
If Val(Form4.Label2.Caption) > 0 Then
   Form4.Label2.Caption = Form4.Label2.Caption - 1
Else
   Timer3.Enabled = False
End If
End Sub
Private Sub menuhistory_Click()
Form5.Visible = True
End Sub

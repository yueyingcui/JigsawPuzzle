VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "��ѧ071�� 070001��ԽӨ"
   ClientHeight    =   9600
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   14340
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "ƴͼ��Ϸ"
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
      Caption         =   "����"
      Begin VB.Menu menucheck 
         Caption         =   "Check OriPicture"
      End
   End
   Begin VB.Menu menumusic 
      Caption         =   "����"
      Begin VB.Menu menumidplay 
         Caption         =   "Play Music"
      End
   End
   Begin VB.Menu menuhistory 
      Caption         =   "Ӣ�۰�"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'װ��ͼƬ����ʹ��CommandDialog�Ի��������У�Filter��������Ϊֻ����װ��Jpg,Gif,BMP���͵��ļ����ɣ����ļ��򿪺�����PictureBox��AutoSize���Կ��Եõ���ʵ��С����ʱ�ٵ��ڴ����С����ʵ������Ӧ��С��
Private Sub menupic_Click() 'װ��ͼƬ����8
CommonDialog1.InitDir = "c:\windows" 'ָ����ʼ�ļ�Ŀ¼P210
CommonDialog1.Filter = "Jpg File(*.jpg)|*.jpg|Gif File(*.gif)|*.gif|Bmp File(*.bmp)|*.bmp" '������
CommonDialog1.FilterIndex = 1 'ָ��Ĭ�Ϲ�����
On Error GoTo UserCancle
CommonDialog1.ShowOpen '��ʾ"��"�Ի���
UserSelectFile = CommonDialog1.FileName 'UserSelectFile���ѡ�����ļ���
OriPicture.Picture = LoadPicture(UserSelectFile) '��ͼƬ�ļ�����ʾ
UserCancle: '���ִ����˳�����
OriPicture.AutoSize = True 'PictureBox��ӦͼƬ��С
End Sub

Private Sub menudegree_Click()
ControlCount = InputBox("�����������", �Ѷ�����, 4)
End Sub

Private Sub menubegin_Click()
Call splitpicture(ControlCount)
Call changepic
Timer2.Enabled = True
Timer3.Enabled = True
Form4.Visible = True
End Sub

'�и�ͼƬ����
Private Sub splitpicture(ControlCount As Integer)
'���û���ʼ��Ϸ�󣬿�������PictureBox�Ŀؼ��������洢��Ƭͼ�񣬶�̬����Ĵ�С���Ը�����Ƭ������������������Ƭ�ϵ�ͼ���������PaintPicture������������ͼ���ϻ�ȡ��
Dim i As Integer, j As Integer, k As Integer
Dim splitcount As Integer
ReDim lOldSplitImage(ControlCount) As Integer
ReDim tOldSplitImage(ControlCount) As Integer
For i = 1 To ControlCount - 1
    Load SplitImage(i)
    SplitImage(i).Visible = True
Next
splitcount = Sqr(ControlCount)


SplitImageWidth = OriPicture.Width / splitcount 'ÿ����ֵ
SplitImageHeight = OriPicture.Height / splitcount  'ÿ��߶�ֵ
For i = 0 To (splitcount - 1) '�ֽ�ͼƬ������˳���������
    For j = 0 To (splitcount - 1)
        k = i * splitcount + j
        SplitImage(k).Width = SplitImageWidth '���ÿ��P226 P166
        SplitImage(k).Height = SplitImageHeight '���ø߶�
        SplitImage(k).AutoRedraw = True
        SplitImage(k).PaintPicture OriPicture.Picture, 0, 0, OriPicture.Width / splitcount, OriPicture.Height / splitcount, j * SplitImageWidth, i * SplitImageHeight, OriPicture.Width / splitcount, OriPicture.Height / splitcount, vbSrcCopy
         'SplitImageWidth��SplitImageHeight��ÿ����Ƭ��Ⱥ͸߶�
         'Paintpicture�����ü�ͼ��[������.paintpicture ͼ��Դ,��ͼ�������x,y,���,�߶�,�ü��������x,y,���,�߶�,���Ʒ�ʽ]P240
        SplitImage(k).AutoRedraw = True
        SplitImage(k).Top = i * SplitImageHeight '�������Ͻ�����P226 j-x,i-y
        SplitImage(k).Left = j * SplitImageWidth
        SplitImage(k).Visible = True '�ؼ�����P167
        SplitImage(k).ZOrder
    Next j
Next i

For i = 0 To ControlCount - 1
   lOldSplitImage(i) = SplitImage(i).Left
   tOldSplitImage(i) = SplitImage(i).Top
Next
End Sub

'����ͼƬ����
Private Sub changepic()
Dim i As Integer, j As Integer, k As Integer
Dim t As Integer, l As Integer, temp As Integer

'OriPicture.Visible = False '����ԭͼƬ

For i = 0 To ControlCount - 1 '����˳�򲢷�������
    
    Randomize '���ѡ����������
     j = Int(Rnd() * ControlCount)
    Randomize
     k = Int(Rnd() * ControlCount)
    
    t = SplitImage(j).Top '����������ͼƬ
    l = SplitImage(j).Left
    SplitImage(j).Top = SplitImage(k).Top
    SplitImage(j).Left = SplitImage(k).Left
    SplitImage(k).Top = t
    SplitImage(k).Left = l
    'temp = j '����ͼƬ����λ��
    'j = k
    'j = temp
Next i
End Sub



'����꽻��ͼƬ
Private Sub SplitImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t As Integer, l As Integer

If sign = 0 Then
   num1 = Index
   sign = 1
ElseIf sign = 1 Then
   num2 = Index
   sign = 0
   
t = SplitImage(num1).Top '����������ͼƬ
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

Private Sub Timer1_Timer() '���Ʋ쿴ԭͼʱ��
Form2.Visible = False
n = n + 1
End Sub

Private Sub WINER(ControlCount As Integer)

For i = 0 To ControlCount - 1
  'ͨ��PIC������ؼ�λ�����ݵıȽ������ж�
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

Private Sub Timer2_Timer() '������Ϸ��ʱ��
 MsgBox "You Lose!"
 End
End Sub

Private Sub Timer3_Timer() '���Ƽ�ʱ����ʾForm4
If Val(Form4.Label2.Caption) > 0 Then
   Form4.Label2.Caption = Form4.Label2.Caption - 1
Else
   Timer3.Enabled = False
End If
End Sub
Private Sub menuhistory_Click()
Form5.Visible = True
End Sub

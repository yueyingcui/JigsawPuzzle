VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form3 
   Caption         =   "Mp3 Player "
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form3"
   ScaleHeight     =   600
   ScaleWidth      =   2745
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   1085
      _Version        =   393216
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
CommonDialog2.InitDir = "c:\windows" 'ָ����ʼ�ļ�Ŀ¼P210
CommonDialog2.Filter = "Mp3 File(*.mp3)|*.mp3" '������
CommonDialog2.DialogTitle = "ָ��Ҫ���ŵ�Mp3�ļ�"
CommonDialog2.FilterIndex = 1 'ָ��Ĭ�Ϲ�����
CommonDialog2.ShowOpen '��ʾ"��"�Ի���
Form3.MMControl1.Command = "Close"
Form3.MMControl1.DeviceType = "MpegVideo" '�豸����Ps214
Form3.MMControl1.TimeFormat = mciFormatMilliseconds 'ָ��ʱ���ʽΪ����
MidFileName = CommonDialog2.FileName 'UserSelectFile���ѡ�����ļ���
Form3.MMControl1.FileName = MidFileName
Form3.MMControl1.Command = "Open"
Load Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MMControl1.Command = "Close"
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "12306����Ԥ���ճ�����"
   ClientHeight    =   9120
   ClientLeft      =   1305
   ClientTop       =   1185
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   5505
   Begin VB.Timer tmrPostOrdertest 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   855
      Top             =   7875
   End
   Begin VB.FileListBox File1 
      Height          =   2070
      Left            =   5670
      Pattern         =   "*.dat"
      TabIndex        =   77
      Top             =   6900
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Timer tmrPostOrder 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   270
      Top             =   7920
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<< ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   74
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Frame Frame5 
      Caption         =   "���ɱ༭"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5640
      TabIndex        =   49
      Top             =   3600
      Width           =   9735
      Begin VB.TextBox txt_fzhzzm 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   63
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt_dzhzzm 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   62
         Top             =   735
         Width           =   1215
      End
      Begin VB.TextBox txt_fjm 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2760
         TabIndex        =   61
         Top             =   375
         Width           =   1215
      End
      Begin VB.TextBox txt_djm 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2760
         TabIndex        =   60
         Top             =   750
         Width           =   1215
      End
      Begin VB.TextBox txt_fhdwmc 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4920
         TabIndex        =   59
         Top             =   375
         Width           =   3135
      End
      Begin VB.TextBox txt_shdwmc 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4920
         TabIndex        =   58
         Top             =   735
         Width           =   3135
      End
      Begin VB.TextBox txt_zcdd 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   57
         Top             =   1215
         Width           =   4335
      End
      Begin VB.TextBox txt_xcdd 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   960
         TabIndex        =   56
         Top             =   1590
         Width           =   4335
      End
      Begin VB.TextBox txt_qqcsMax 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6480
         TabIndex        =   55
         Top             =   1215
         Width           =   1575
      End
      Begin VB.TextBox txt_hzpm 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6480
         TabIndex        =   54
         Top             =   1590
         Width           =   1575
      End
      Begin VB.TextBox txt_dzyx 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2025
         TabIndex        =   53
         Top             =   2025
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txt_dztmism 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   52
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_fztmism 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3810
         TabIndex        =   51
         Top             =   2055
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txt_fzyx 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5355
         TabIndex        =   50
         Top             =   2055
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "��վ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   480
         TabIndex        =   73
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "��վ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   480
         TabIndex        =   72
         Top             =   735
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   71
         Top             =   390
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2280
         TabIndex        =   70
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4200
         TabIndex        =   69
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "�ջ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4200
         TabIndex        =   68
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "װ���ص�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   67
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "ж���ص�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   1590
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5520
         TabIndex        =   65
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5520
         TabIndex        =   64
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   5640
      TabIndex        =   30
      Top             =   240
      Width           =   9735
      Begin VB.TextBox txtExtcode 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8280
         TabIndex        =   76
         Top             =   2070
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txt_xqslh 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         TabIndex        =   40
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chk_ifzzjg 
         Caption         =   "װ�ؼӹ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3480
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cbo_qqcz 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Form1.frx":0000
         Left            =   840
         List            =   "Form1.frx":0037
         TabIndex        =   38
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txt_qqcs 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4200
         TabIndex        =   37
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txt_qqds 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6960
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_hqhw 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   35
         Top             =   1515
         Width           =   1695
      End
      Begin VB.CheckBox chk_dddxtz 
         Caption         =   "�ջ��˽��յ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txt_shdwdh 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4080
         TabIndex        =   33
         Top             =   2115
         Width           =   1935
      End
      Begin VB.TextBox txt_zcrq 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4200
         TabIndex        =   32
         Top             =   345
         Width           =   1935
      End
      Begin VB.TextBox txt_pzycfh 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6600
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   48
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label13 
         Caption         =   "װ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   46
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label8 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         Top             =   848
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6360
         TabIndex        =   44
         Top             =   885
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "������λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   43
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "�ջ����ֻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2880
         TabIndex        =   42
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "(�ֶ���д�밴""2014-01-05""�ĵĸ�ʽ��д)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   180
         Left            =   6240
         TabIndex        =   41
         Top             =   360
         Width           =   3420
      End
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "��ʼ�Զ��ύ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3795
      Picture         =   "Form1.frx":00CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7800
      Width           =   1605
   End
   Begin VB.CommandButton cmdDeAuto 
      Caption         =   "ֹͣ�Զ��ύ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      Picture         =   "Form1.frx":015C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7800
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "�Զ��ᱨ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   5295
      Begin VB.CheckBox chk_saveacc 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox txtUsername 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmd_profile 
         Caption         =   "���浱ǰ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3660
         Picture         =   "Form1.frx":01ED
         TabIndex        =   75
         Top             =   4860
         Width           =   1470
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�����ֶ��ύ!"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3630
         Picture         =   "Form1.frx":0874
         TabIndex        =   9
         Top             =   5565
         Width           =   1500
      End
      Begin VB.TextBox Txt_AllowAuto 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1425
         TabIndex        =   8
         Top             =   4290
         Width           =   2265
      End
      Begin VB.TextBox txt_outtime 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   26
         Text            =   "60"
         Top             =   4755
         Width           =   375
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   4935
         Begin VB.OptionButton Option2 
            Caption         =   "�ֶ���д"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   23
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton Option2 
            Caption         =   "�Զ���д(������1)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.ComboBox txt_orderlist 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   4935
      End
      Begin VB.CommandButton cmd_getorder 
         Caption         =   "��ȡԤ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4020
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmd_login 
         Caption         =   "��¼"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4020
         TabIndex        =   4
         Top             =   735
         Width           =   1095
      End
      Begin VB.TextBox txt_zcrqper 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   5
         Top             =   1665
         Width           =   1740
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2880
         TabIndex        =   17
         Top             =   1665
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "30���Ժ��1��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3585
         TabIndex        =   16
         Top             =   1680
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.TextBox txtPassWord 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   795
         Width           =   1740
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "�Զ��ύʱ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   4320
         Width           =   1170
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "���ʳ�ʱʱ��:      ��(ƽʱ���)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   4800
         Width           =   2790
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "δ��¼"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1455
         TabIndex        =   25
         Top             =   1245
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ��½״̬:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   1245
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��д����װ����Ϣ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   3480
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "ѡ��Ԥ����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "װ������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   1710
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�ʺ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   405
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ���ѱ������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.ComboBox txt_profile 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   4800
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   30
      X2              =   5505
      Y1              =   8460
      Y2              =   8460
   End
   Begin VB.Label lblInfo 
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   105
      TabIndex        =   29
      Top             =   8670
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LocalIP As String
Public http As WinHttp.WinHttpRequest
Public sen As String, sen2 As String, sen3 As String
Public vcodeIndex As Long
Public jsonorder As String, jsonorder2 As String, uuid As String
Public ISAUTO As Boolean, ISLOGIN As Boolean, ISOFFLINE As Boolean
Public JsonselIndex As Integer
Public city As String, testurl As String, testurl2 As String
Public yzmCode As String
Public heartline As Integer


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function LoadLibFromFile Lib "Sunday.dll" (ByVal FilePath As String, ByVal pass As String) As Long
Private Declare Function GetCodeFromBuffer Lib "Sunday.dll" (ByVal CdsFileIndex As Long, ByVal ImgBuffer As Long, ByVal ImgBufLen As Long, ByVal Vcode As String) As Boolean

Private Sub Form_Load()

    'city = "wulmq"
    city = "beij"
    'city = "taiy"
    
    If city = "beij" Then
        testurl = "_test1"  '��½��ַ
        testurl2 = "_test1" '"_test2"  '��ҳ��ַ
    Else
        esturl = ""
        testurl2 = ""
    End If
    JsonselIndex = -1
    Set http = New WinHttp.WinHttpRequest
    http.Option(4) = 13056
    http.Option(6) = False
    http.SetTimeouts 60000, 60000, 60000, 60000
    ISAUTO = False
    ISOFFLINE = False
    yzmCode = ""
    heartline = 0
    
    '�ύ����
    If Hour(Now()) >= 8 Then '8���Ժ�ڶ����ύ
        Txt_AllowAuto.Text = Format(DateAdd("d", 1, Now()), "yyyy-mm-dd 07:00:00")
    Else '7��֮ǰ�����ύ
        Txt_AllowAuto.Text = Format(Now(), "yyyy-mm-dd 07:00:00")
    End If
    
    'װ������
    txt_zcrqper.Text = Trim(Format(DateAdd("d", 31, Now()), "yyyy-mm-dd"))
    
    '�����˺�
    Call bindAccount
    
    '���ر���������ļ�
    File1.Path = App.Path & "\dat\"
    For i = 0 To File1.ListCount - 1
        If File1.List(i) <> "dat000.dat" Then
            txt_profile.AddItem Left(File1.List(i), Len(File1.List(i)) - 4)
        End If
    Next
    'txt_profile.ListIndex = txt_profile.ListCount - 1

    
    Call showinfo(3, "��ǰ�ʺ���δ��¼,���Ȳ��Ե�¼!")

End Sub



'���ض�����Ϣ
Private Sub txt_profile_click()
    
    Call showinfo(2, "�ù����ݲ�����,���޸�")
    Exit Sub

    If ISLOGIN = True Then
        Call showinfo(2, "��ǰ�ѵ�¼,�޷����ض�����Ϣ")
        Exit Sub
    End If
    
    Call loadProfile(txt_profile.List(txt_profile.ListIndex))
    
    Call showinfo(1, "������Ϣ�������!")
End Sub

'ѡ���û���
Private Sub txtUsername_Click()
    Call bindAccount(txtUsername.Text)
End Sub

'���¼
Private Sub cmd_login_Click()
    Dim funRe As String
    
    Dim username As String, password As String
    If txtUsername.Text = "" Then
        Call showinfo(2, "�������û���!")
        txtUsername.SetFocus
        Exit Sub
    End If
    username = Trim(txtUsername.Text)
    
    If txtPassWord.Text = "" Then
        Call showinfo(2, "����������!")
        txtPassWord.SetFocus
        Exit Sub
    End If
    password = Trim(txtPassWord.Text)
    
    
    Call showinfo(3, "��¼��,���Ե�....")
    cmd_login.Enabled = False
    
    funRe = intiAndLoginFull(username, password)
    
    If CheckFunRe(funRe, 1) <> 1 Then
        Call showinfo(2, "��¼ʧ��,����ԭ��:" & CheckFunRe(funRe, 2))
        cmd_login.Enabled = True
    Else
        Label6.ForeColor = &HD000&
        Label6.Caption = "�ѵ�¼(" & CheckFunRe(funRe, 2) & ")"
        ISLOGIN = True
        cmd_login.Enabled = True
        
        '�ɹ��Ժ��ٱ���
        If chk_saveacc.Value = 1 Then
            Call saveAccount(txtUsername.Text, txtPassWord.Text)
            Call showinfo(1, "��¼�ɹ�,�˺��������Զ�����!")
        Else
            Call showinfo(1, "��¼�ɹ�!")
        End If
    End If
End Sub

'���ȡԤ����
Private Sub cmd_getorder_Click()
    Dim funRe As String
    
    If ISLOGIN = False Then
        Call showinfo(2, "���ȵ�¼!")
        Exit Sub
    End If
    
    Call showinfo(3, "��ȡԤ������,���Ե�....")
    cmd_getorder.Enabled = False
    
    funRe = GetOrderNo()
    
    If CheckFunRe(funRe, 1) <> 1 Then
        Call showinfo(2, "��ȡʧ��,����ԭ��:" & CheckFunRe(funRe, 2))
        cmd_getorder.Enabled = True
    Else
        Call showinfo(1, "��ȡ�ɹ�,��ѡ��Ԥ����")
        cmd_getorder.Enabled = True
    End If
End Sub

'ѡ��Ԥ����
Private Sub txt_orderlist_Click()
    
    If ISOFFLINE = True Then Exit Sub

    Dim funRe As String
    
    If txt_orderlist.ListCount = 0 Then
        Call showinfo(2, "���Ȼ�ȡԤ����!")
        Exit Sub
    End If
    
    Call showinfo(3, "����Ԥ���Ż�ȡ������Ϣ��....")
    cmd_getorder.Enabled = False
    
    funRe = GetInfoByOrderNo(txt_orderlist.ListIndex)
    
    If CheckFunRe(funRe, 1) <> 1 Then
        Call showinfo(2, "��ȡʧ��,����ԭ��:" & CheckFunRe(funRe, 2))
        cmd_getorder.Enabled = True
    Else
        Call showinfo(1, "������д���,���Խ����ֶ����Զ��ύģʽ")
        cmd_getorder.Enabled = True
    End If
End Sub

'���Զ��ύ
Private Sub cmdAuto_Click()

    Dim offline As Integer

    '����7�㵽11��֮�� ʹ�����߶�����ʾ

    

    'offline = MsgBox("�Ƿ�Ҫʹ�����߶����ύ?", vbYesNo, "�Զ��ύ")
    
    'If offline = vbNo Then
    '    Exit Sub
    'End If

    If txtUsername.Text = "" Or txtPassWord.Text = "" Or txt_zcrqper.Text = "" Or txt_pzycfh.Text = "" Or Txt_AllowAuto.Text = "" Or txt_xqslh.Text = "" Then
        Call showinfo(2, "������д����ȫ,���ֶ���¼��ȡ������Ϣ���ٵ���Զ��ύ")
        Exit Sub
    End If
    
    If JsonselIndex = -1 Or jsonorder = "" Or jsonorder2 = "" Then
        Call showinfo(2, "������д����ȫ,��ѡ��Ԥ���Ż��ֶ���д������Ϣ���ٵ���Զ��ύ")
        Exit Sub
    End If

    ISAUTO = True
    Call showinfo(3, "�Զ��ύ������,Ϊ���������,�벻Ҫ���������ť")
    Call SavePage("[" & Now() & "]�Զ��ύ����...", "syslog")
    
    tmrPostOrdertest.Interval = 5000
    tmrPostOrdertest.Enabled = True
    
    Call lockAll
       
End Sub

'��ȡ���Զ��ύ
Private Sub cmdDeAuto_Click()
    ISAUTO = False
    Call showinfo(2, "�Զ��ύ�ر�")
    tmrPostOrdertest.Enabled = False
    
    Call unlockAll

End Sub

'�Զ��ύ����
Private Sub tmrPostOrder_Timer()

    On Error Resume Next
    DoEvents
    Dim funRe As String
    funRe = 0
    
    Call showinfo(3, "�Զ��ύ��,Ϊ���������,�벻Ҫ���������ť")
    
    tmpTime = DateDiff("s", Now(), Txt_AllowAuto)
    
    '��ǰ�����ӻ�ȡ��֤��
    If tmpTime > 300 Then
       Call showinfo(2, "δ���ύʱ��,ϵͳ������,����" & tmpTime \ 60 & "�ֿ�ʼ�ύ")
       Exit Sub
    ElseIf yzmCode = "" Then
        Call SavePage("[" & Now() & "]�Զ��ύ��ʼ����ʼ", "syslog")
        Do
            funRe = inti(txtUsername.Text)
            
            If CheckFunRe(funRe, 1) <> 1 Then
                Call SavePage("[" & Now() & "]��¼��ʼ��ʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
            End If
            
            Sleep (1000)
            
        Loop Until CheckFunRe(funRe, 1) = 1
    End If
   
    
    '��ǰ5�뿪ʼ�ύ
    If tmpTime > 0 Then
        Call showinfo(2, "δ���ύʱ��,ϵͳ������,����" & tmpTime \ 60 & "�ֿ�ʼ�ύ")
        Exit Sub
    End If
    
    Call SavePage("[" & Now() & "]�Զ��ύ��ʼ,��ʼ��¼", "syslog")
    
    '��½
    Do
        funRe = Login(txtUsername.Text, txtPassWord.Text)
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]��½ʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
            If CheckFunRe(funRe, 2) = "ϵͳά����" Then
                Exit Sub
            End If
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]��½�ɹ�,��ʼ�ύ", "syslog")
    
    http.SetTimeouts 180000, 180000, 180000, 180000
    
    '��½��ֱ���ύ,������鶩����
    Do
        funRe = PerPost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]Ԥ�ύʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
            
            If CheckFunRe(funRe, 2) = "������Ԥ�����ڷ�Χ" Or CheckFunRe(funRe, 2) = "δ�ҵ���Ӧ��������Ϣ" Then
            
                '��ȷʧ��
                Call SavePage("[" & Now() & "]" & CheckFunRe(funRe, 2) & ",�Զ��ύ�ر�", "syslog")
                ISAUTO = False
                Call showinfo(2, "��Ϣ��д��ʱ��ѡ�����,�Զ��ύ�ر�!")
                tmrPostOrder.Enabled = False
                
                Call unlockAll
                
                Exit Sub
            End If
            
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]Ԥ�ύ�ɹ�,uuid=" & uuid & ",��ʼ��ʽ�ύ", "syslog")
    '��ʽ�ύ
    Do
        funRe = RePost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]��ʽ�ύʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]�ύ�ɹ�,�Զ��ύ�ر�", "syslog")
    
    
    ISAUTO = False
    Call showinfo(1, "�ύ���,�Զ��ύ�ر�!")
    tmrPostOrder.Enabled = False
    
    Call unlockAll
    
End Sub

'�²����Զ��ύ����
Private Sub tmrPostOrdertest_Timer()
    On Error Resume Next
    DoEvents
    Dim funRe As String
    funRe = 0
    heartline = heartline + 1
    
    Call showinfo(3, "�Զ��ύ��,Ϊ���������,�벻Ҫ���������ť")
    
    tmpTime = DateDiff("s", Now(), Txt_AllowAuto)
    
    '��ǰ����Ӳ�����������
    If tmpTime > 300 Then
        Call showinfo(2, "δ���ύʱ��,ϵͳ������,����" & (tmpTime \ 60) + 1 & "�ֿ�ʼ�ύ")
        
        If heartline > 50 Then
            Call SavePage("[" & Now() & "]�������ӿ�ʼ" & sen2, "syslog")
           
            funRe = inti1(txtUsername.Text)
            
            If CheckFunRe(funRe, 1) <> 1 Then
                Call SavePage("[" & Now() & "]��������ʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
            End If
            
            heartline = 0
        End If
        
       Exit Sub
    End If
   
    
    '��ǰ5�뿪ʼ�ύ
    If tmpTime > 5 Then
        Call showinfo(2, "δ���ύʱ��,ϵͳ������,����" & tmpTime \ 60 & "�ֿ�ʼ�ύ")
        Exit Sub
    End If
    
    Call SavePage("[" & Now() & "]�Զ��ύ��ʼ,��ʼԤ�ύ", "syslog")
    
    '��½��ֱ���ύ,������鶩����
    Do
        funRe = PerPost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]Ԥ�ύʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
            
            If CheckFunRe(funRe, 2) = "������Ԥ�����ڷ�Χ" Or CheckFunRe(funRe, 2) = "δ�ҵ���Ӧ��������Ϣ" Then
            
                '��ȷʧ��
                Call SavePage("[" & Now() & "]" & CheckFunRe(funRe, 2) & ",�Զ��ύ�ر�", "syslog")
                ISAUTO = False
                Call showinfo(2, "��Ϣ��д��ʱ��ѡ�����,�Զ��ύ�ر�!")
                tmrPostOrdertest.Enabled = False
                
                Call unlockAll
                
                Exit Sub
            End If
            
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]Ԥ�ύ�ɹ�,uuid=" & uuid & ",��ʼ��ʽ�ύ", "syslog")
    '��ʽ�ύ
    Do
        funRe = RePost()
        
        If CheckFunRe(funRe, 1) <> 1 Then
            Call SavePage("[" & Now() & "]��ʽ�ύʧ��,����ԭ��:" & CheckFunRe(funRe, 2), "syslog")
        End If
        
        Sleep (1000)
        
    Loop Until CheckFunRe(funRe, 1) = 1
    
    Call SavePage("[" & Now() & "]�ύ�ɹ�,�Զ��ύ�ر�", "syslog")
    
    
    ISAUTO = False
    Call showinfo(1, "�ύ���,�Զ��ύ�ر�!")
    tmrPostOrdertest.Enabled = False
    
    Call unlockAll
End Sub

'���ֶ��ύ
Private Sub Command3_Click()
    On Error Resume Next
    
    If ISAUTO = True Then
        MsgBox "�Զ��ύ������,�޷����в���!"
        Exit Sub
    End If

    If txt_pzycfh.Text = "" Then Call showinfo(2, "���ϲ�����,����д���������ύ"): Exit Sub

    Call showinfo(3, "�ύ������,���𷴸����!")
    Command3.Enabled = False
    Dim surl As String, param As String
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_add"
    param = ""
    param = param & "currentPosition=" & "%E9%A2%84%E7%BA%A6%C2%A0%3E%3E%C2%A0%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    param = param & "&" & "djm=" & URLEncodeUTF8(txt_djm.Text)
    param = param & "&" & "dzhzzm=" & URLEncodeUTF8(txt_dzhzzm.Text)
    param = param & "&" & "dztmism=" & txt_dztmism.Text
    param = param & "&" & "dzyx=" & Replace(txt_dzyx.Text, " ", "+")
    param = param & "&" & "fhdwmc=" & URLEncodeUTF8(txt_fhdwmc.Text)
    param = param & "&" & "fjm=" & URLEncodeUTF8(txt_fjm.Text)
    param = param & "&" & "fzhzzm=" & URLEncodeUTF8(txt_fzhzzm.Text)
    param = param & "&" & "fztmism=" & txt_fztmism.Text
    param = param & "&" & "fzyx=" & Replace(txt_fzyx.Text, " ", "+")
    param = param & "&" & "hzpm=" & URLEncodeUTF8(txt_hzpm.Text)
    param = param & "&" & "keyword="
    param = param & "&" & "maxDate=" & Trim(txt_zcrq.Text) '& Format(DateAdd("m", 1, Now()) - 1, "yyyy-mm-dd")
    param = param & "&" & "minDate=" & Format(Now() + 3, "yyyy-mm-dd")
    param = param & "&" & "po.dddxtz=" & chk_dddxtz.Value
    param = param & "&" & "po.hqhw=" & txt_hqhw.Text
    param = param & "&" & "po.pzycfh=" & txt_pzycfh.Text
    param = param & "&" & "po.qqcs=" & txt_qqcs.Text
    param = param & "&" & "po.qqcz=" & Right(cbo_qqcz.Text, 1)
    param = param & "&" & "po.qqds=" & txt_qqds.Text
    param = param & "&" & "po.qqlx=0"
    param = param & "&" & "po.shdwdh=" & txt_shdwdh.Text
    param = param & "&" & "po.uuid=" '8ac086a9441480d4014419d6acbe0064"
    param = param & "&" & "po.xqslh=" & txt_xqslh.Text
    
    param = param & "&" & "po.zcrq=" & Trim(txt_zcrq.Text)
    
    param = param & "&" & "qqcsMax=" & txt_qqcsMax.Text
    param = param & "&" & "shdwmc=" & URLEncodeUTF8(txt_shdwmc.Text)
    param = param & "&" & "xcdd=" & URLEncodeUTF8(txt_xcdd.Text)
    param = param & "&" & "zcdd=" & URLEncodeUTF8(txt_zcdd.Text)
    
    
    Call SavePage("[" & Now() & ":step1]" & param & vbLf, "perpostdata")
    
    http.Open "POST", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        Call showinfo(2, "��ʱ1,�������ύ!")
        Command3.Enabled = True
        Exit Sub
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage("[" & Now() & ":step1]" & body1 & vbLf, "pergetdata")
    
    If InStr(body1, """success"":true") Then
        uuid = mySubstr(body1, "uuid"":""", """")
        
        param = "op=10&uuids=" & uuid & ",&mor_dzsw_security_info=mor_dzsw_security_disabled"
        Call SavePage("[" & Now() & ":step2]" & param & vbLf, "perpostdata")
        surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_operateZcrbjh"
        
        http.Open "POST", surl, False
        http.SetRequestHeader "Connection", "Keep-Alive"
        http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
        http.SetRequestHeader "Cache-Control", "no-cache"
        http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
        http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
        http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
        http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
        http.SetRequestHeader "Cookie", "CASTGC=" & sen3
        http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
        http.SetRequestHeader "Content-Length", Len(param)
        http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
        http.Send param
        
        If Err.Number <> 0 Then
            Err.Clear
            Call showinfo(2, "��ʱ2,�������ύ!")
            Command3.Enabled = True
            Exit Sub
        End If
        
        body2 = BytesToBstr(http.ResponseBody, 2)
        Call SavePage("[" & Now() & ":step2]" & body2 & vbLf, "pergetdata")
        
        If InStr(body2, """success"":true") Then
            Call showinfo(1, "�ֶ��ᱨ�ɹ�!")
            Command3.Enabled = True
            Exit Sub
        Else
            Call showinfo(2, "�ᱨʧ��,������־!")
            Command3.Enabled = True
            Exit Sub
        End If
    ElseIf InStr(body1, "������Ԥ�����ڷ�Χ") Then
        Call showinfo(2, "������Ԥ�����ڷ�Χ!")
        Command3.Enabled = True
        Exit Sub
    ElseIf InStr(body1, "δ�ҵ���Ӧ��������Ϣ") Then
        Call showinfo(2, "δ�ҵ���Ӧ��������Ϣ,�����¼����ѡԤ����!")
        Command3.Enabled = True
        Exit Sub
    Else
       Call showinfo(2, "Ԥ�ᱨʧ��,������־!")
       Command3.Enabled = True
       Exit Sub
    End If
    
End Sub


'��������
Private Sub cmd_profile_Click()
    Dim filen As String
    
    If ISLOGIN = False Then
        Call showinfo(2, "���ȵ�¼���ٱ��浱ǰ�û���������!")
        Exit Sub
    End If
    
    
    If txtUsername.Text = "" Then
        Call showinfo(2, "�������û���!")
        txtUsername.SetFocus
        Exit Sub
    End If

    
    If txtPassWord.Text = "" Then
        Call showinfo(2, "����������!")
        txtPassWord.SetFocus
        Exit Sub
    End If
    
    If txt_zcrqper.Text = "" Then
        Call showinfo(2, "������װ��ʱ��!")
        txt_zcrq.SetFocus
        Exit Sub
    End If

    
    If jsonorder = "" Then
        Call showinfo(2, "���Ȼ�ȡԤ����!")
        cmd_getorder.SetFocus
        Exit Sub
    End If
    
    If jsonorder2 = "" Or JsonselIndex = -1 Then
        Call showinfo(2, "��ѡ������Ԥ����!")
        txt_orderlist.SetFocus
        Exit Sub
    End If
    
    filen = "[" & txtUsername.Text & "]" & txt_zcrqper.Text & "_��������" & txt_qqcs & "_��վ��" & txt_dzhzzm & "_���" & txt_hzpm

    filen = InputBox("������д�ᱨ��Ϣ����Ϊ:", "��������", filen)
    
    filen = Replace(Replace(Replace(Replace(Replace(Replace(filen, ":", ""), "/", ""), "\", ""), "|", ""), """", ""), "?", "")
    
    Call saveProfile(filen)
    
    Call showinfo(1, "��ǰ�����ѱ���!")

End Sub

'**************************************************AUTOר�ú�����*********************************************************
Function inti(user As String) As String

    On Error Resume Next

    Dim ImgFile As String
    Dim Image() As Byte
    
   
    'ֱ����֤��
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/security/jcaptcha.jpg"
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        inti = "0|��ȡ��֤�볬ʱ002"
        Exit Function
    End If

    ImgFile = Fun_SaveImgToFile(http.ResponseBody, user & ".jpg", App.Path & "\")
    
    If Err.Number <> 0 Then
        Err.Clear
        inti = "0|��֤���ȡʧ��"
        Exit Function
    End If
    
    vcodeIndex = LoadLibFromFile("12306.lib", "123")
    
    If Err.Number <> 0 Then
        Err.Clear
        inti = "0|��֤��ʶ���������ʧ��"
        Exit Function
    End If


    If (vcodeIndex = -1) Then
        inti = "0|��֤��ʶ������ʧ��"
        Exit Function
    End If
    
    Dim Vcode As String
    Vcode = "      " '�����ȶ��������������ո񣬿ո�����Ҫ����֤���ַ�������1
   
    Call MyReadFile(ImgFile, Image)
     '�ڴ�ӿڵ�����֤��ͼ��ʶ��
    If (GetCodeFromBuffer(vcodeIndex, VarPtr(Image(0)), UBound(Image), Vcode)) Then
        txtExtcode.Text = Vcode
        yzmCode = Trim(txtExtcode.Text)
        
        head = http.GetAllResponseHeaders
        headers = Split(head, Chr(10))
        
        For ii = LBound(headers) To UBound(headers)
            If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
                p2 = InStr(headers(ii), ";")
                s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
                p2 = InStr(s, "=")
                s1 = Trim(Mid(s, 1, p2 - 1))
                s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
                        
                If s1 = "BIGipServerhyswpt_pool" Then
                    sen = s2
                End If
                
                If s1 = "DZSW_SESSIONID" Then
                    sen2 = s2
                End If
            End If
        Next
        
        inti = "1|ʶ��ɹ�"
    Else
        inti = "0|��֤��ʶ��ʧ��"
        Exit Function
    End If

End Function

'��������
Function inti1(user As String) As String

    On Error Resume Next
    
   
    'ֱ����֤��
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/security/jcaptcha.jpg"
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        inti1 = "0|�������ӳ�ʱ"
        Exit Function
    End If

        
    head = http.GetAllResponseHeaders
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
                    
            If s1 = "BIGipServerhyswpt_pool" Then
                sen = s2
            End If
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
        End If
    Next
    
    inti1 = "1| " & sen2

End Function
'�Զ���½
Function Login(user As String, pass As String) As String
    
    On Error Resume Next
    
    Dim username As String, password As String, extcode As String
    Dim param As String
    
    
    username = user
    password = pass
    extcode = Trim(txtExtcode.Text)
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/j_spring_security_check"
    param = "j_username=" & username & "&j_password=" & password & "&j_captcha=" & extcode & "&fromUrl=%2Flogin_bur.jsp"
    
    http.Open "POST", surl, False
    http.Option(WinHttpRequestOption_EnableRedirects) = 0
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"

    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        Login = "0|��¼��ʱ"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage("[" & Now() & "]httpStatus:" & http.Status & body1, "login")
    
    
    If InStr(body1, "Dzsw/home.jsp") > 0 Then
        Login = "1|��¼�ɹ�"
    Else
        Login = "0|�Զ���¼ʧ��"
        Exit Function
    End If
    
    
    '����cookie
    head = http.GetAllResponseHeaders
            
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
            
            If s1 = "CASTGC" Then
                sen3 = s2
            End If
        End If
    Next
    
    Exit Function
    
End Function

Function intiAndLoginFull(user As String, pass As String) As String
    
    On Error Resume Next

    Dim ImgFile As String
    Dim Image() As Byte
    
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|���糬ʱ001"
        Exit Function
    End If
    
    head = http.GetAllResponseHeaders
            
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
                    
            If s1 = "BIGipServerhyswpt_pool" Then
                sen = s2
            End If
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
        End If
    Next
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    If InStr(body1, "ϵͳ����ά����") > 0 Then
        intiAndLoginFull = "0|ϵͳά����"
        Exit Function
    End If
    
    '�ȶ�����ʾ��֤��
    ' src="/vcode.php?rnd=78475"/>
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/security/jcaptcha.jpg"
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|��ȡ��֤�볬ʱ002"
        Exit Function
    End If
    

    ImgFile = Fun_SaveImgToFile(http.ResponseBody, user & ".jpg", App.Path & "\")
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|��֤���ȡʧ��"
        Exit Function
    End If
    
    vcodeIndex = LoadLibFromFile("12306.lib", "123")
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|��֤��ʶ���������ʧ��"
        Exit Function
    End If


    If (vcodeIndex = -1) Then
        intiAndLoginFull = "0|��֤��ʶ������ʧ��"
        Exit Function
    End If
    
    Dim Vcode As String
    Vcode = "      " '�����ȶ��������������ո񣬿ո�����Ҫ����֤���ַ�������1
   
    Call MyReadFile(ImgFile, Image)
     '�ڴ�ӿڵ�����֤��ͼ��ʶ��
    If (GetCodeFromBuffer(vcodeIndex, VarPtr(Image(0)), UBound(Image), Vcode)) Then
        txtExtcode.Text = Vcode
    Else
        intiAndLoginFull = "0|��֤��ʶ��ʧ��"
        Exit Function
    End If
    
    Dim username As String, password As String, extcode As String
    Dim param As String
    
    
    username = user
    password = pass
    extcode = Trim(txtExtcode.Text)
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/j_spring_security_check"
    param = "j_username=" & username & "&j_password=" & password & "&j_captcha=" & extcode & "&fromUrl=%2Flogin_bur.jsp"
    
    http.Open "POST", surl, False
    http.Option(WinHttpRequestOption_EnableRedirects) = 1
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"

    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        intiAndLoginFull = "0|��¼��ʱ"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage(http.Status & body1, "login")
    
    
    If InStr(body1, "margin-left:50px;"">��ӭ����") > 0 Then
    'If InStr(body1, "Dzsw/home.jsp") > 0 Then
        intiAndLoginFull = "1|" & mySubstr(body1, ";white-space:nowrap;margin-left:5px;"">", "</span>")
    ElseIf InStr(body1, "ϵͳ����ά����") > 0 Then
        intiAndLoginFull = "0|ϵͳά����"
        Exit Function
    ElseIf InStr(body1, "��֤�����벻��ȷ") > 0 Then  '��֤�����벻��ȷ
        intiAndLoginFull = "0|��֤�����"
        Exit Function
    Else
        intiAndLoginFull = "0|��¼ʧ��,�����û���������"
        Exit Function
    End If
    
    
    '����cookie
    head = http.GetAllResponseHeaders
            
    headers = Split(head, Chr(10))
    
    For ii = LBound(headers) To UBound(headers)
        If Left(headers(ii), Len("Set-Cookie:")) = "Set-Cookie:" Then
            p2 = InStr(headers(ii), ";")
            s = Mid(headers(ii), Len("Set-Cookie:") + 1, p2 - Len("Set-Cookie:") - 1)
            p2 = InStr(s, "=")
            s1 = Trim(Mid(s, 1, p2 - 1))
            s2 = Trim(Mid(s, p2 + 1, Len(s) - p2))
            
            If s1 = "DZSW_SESSIONID" Then
                sen2 = s2
            End If
            
            If s1 = "CASTGC" Then
                sen3 = s2
            End If
        End If
    Next
    
    Exit Function
    
End Function

Function GetOrderNo() As String


    On Error Resume Next
    
    If ISLOGIN = False Then
       GetOrderNo = "0|���ȵ�¼"
       Exit Function
    End If
    
    Dim i As Integer
    Dim body1 As String, tmpStr As String
    
    If txt_zcrqper.Text = "" Then
        GetOrderNo = "0|��ѡ��װ������"
        Exit Function
    End If
    
    'https://frontier."& city &".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_getYsxq?q=%E7%8E%89%E7%B1%B3&limit=50&timestamp=1389019837982&zcrq=2014-01-08
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_getYsxq?q="
    surl = surl & "&limit=50&timestamp=1389019837982&zcrq="
    surl = surl & Trim(txt_zcrqper.Text)
    
    http.Open "GET", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "User-Agent", "Mozilla/4.0"
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send
    
    If Err.Number <> 0 Then
        Err.Clear
        GetOrderNo = "0|��ȡԤ���ų�ʱ"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage(body1, "jsonorder")
    
    jsonorder = body1
    
    If body1 <> "[]" Then
    
        txt_orderlist.Enabled = True
        txt_orderlist.Clear
        For i = 1 To lenJSON(body1)
            tmpStr = ""
            tmpStr = tmpStr & parseJSON(body1, "XQSLH", i)(0) & "("
            'tmpStr = tmpStr & parseJSON(body1, "FZHZZM", i)(0) & "|"
            'tmpStr = tmpStr & parseJSON(body1, "FHDWMC", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "DZHZZM", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "SHDWMC", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "HZPM", i)(0) & "|"
            tmpStr = tmpStr & parseJSON(body1, "CZ", i)(0) & "|"
    
            tmpStr = tmpStr & (CLng(parseJSON(body1, "PZCS", i)(0)) - CLng(parseJSON(body1, "JDZC4", i)(0)) - CLng(parseJSON(body1, "YPWZ", i)(0)) - CLng(parseJSON(body1, "YQWP", i)(0)) - CLng(parseJSON(body1, "FACS", i)(0))) & ")"
    
            txt_orderlist.AddItem tmpStr
        Next
        
        GetOrderNo = "1|��ȡԤ���ųɹ�"
        Exit Function
    Else
        txt_orderlist.Clear
        GetOrderNo = "0|û���ҵ��κ�Ԥ����"
        Exit Function
    End If
End Function


Function GetInfoByOrderNo(selIndex As Integer, Optional line As String = "online") As String
    On Error Resume Next

    Dim i As Integer, sycs As Long
    
    JsonselIndex = selIndex
    
    i = selIndex + 1
    
    If selIndex >= 0 Then
    
        sycs = (CLng(parseJSON(jsonorder, "PZCS", i)(0)) - CLng(parseJSON(jsonorder, "JDZC4", i)(0)) - CLng(parseJSON(jsonorder, "YPWZ", i)(0)) - CLng(parseJSON(jsonorder, "YQWP", i)(0)) - CLng(parseJSON(jsonorder, "FACS", i)(0)))
        
        txt_xqslh.Text = parseJSON(jsonorder, "XQSLH", i)(0)
        txt_fzhzzm.Text = parseJSON(jsonorder, "FZHZZM", i)(0)
        txt_fjm.Text = parseJSON(jsonorder, "FJQC", i)(0)
        txt_fhdwmc.Text = parseJSON(jsonorder, "FHDWMC", i)(0)
        txt_dzhzzm.Text = parseJSON(jsonorder, "DZHZZM", i)(0)
        txt_djm.Text = parseJSON(jsonorder, "DJQC", i)(0)
        txt_shdwmc.Text = parseJSON(jsonorder, "SHDWMC", i)(0)
        txt_hzpm.Text = parseJSON(jsonorder, "HZPM", i)(0)
        txt_hqhw.Text = parseJSON(jsonorder, "HQHW", i)(0)
        txt_zcrq.Text = Trim(txt_zcrqper.Text)
        
        txt_qqcs.Text = 1
        txt_qqds.Text = txt_qqcs.Text * 60
        txt_qqcsMax.Text = sycs
        
        txt_pzycfh.Text = parseJSON(jsonorder, "PZYCFH", i)(0)
        txt_dztmism.Text = parseJSON(jsonorder, "DZTMISM", i)(0)
        txt_fztmism.Text = parseJSON(jsonorder, "FZTMISM", i)(0)
        
        
        
        If (parseJSON(jsonorder, "IFZZJG", i)(0)) = 1 Then chk_ifzzjg.Value = 1
        
        
        s = parseJSON(jsonorder, "CZ", i)(0)
        For ii = 0 To cbo_qqcz.ListCount
            If InStr(cbo_qqcz.List(ii), s) Then
                cbo_qqcz.ListIndex = ii
                Exit For
            End If
        Next

        
        If line = "online" Then '���߻�ȡ ����ֱ�Ӽ����ڴ����jsonorder2
        
            surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_getZyxByPzycfh"
            param = "pzycfh=" & parseJSON(jsonorder, "PZYCFH", i)(0)
            http.Open "POST", surl, False
            http.SetRequestHeader "Connection", "Keep-Alive"
            http.SetRequestHeader "User-Agent", "Mozilla/4.0"
            http.SetRequestHeader "Cache-Control", "no-cache"
            http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
            http.SetRequestHeader "Accept", "application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*"
            http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
            http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
            http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/login_bur.jsp"
            http.SetRequestHeader "Content-Length", Len(param)
            http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            http.Send param
            
            If Err.Number <> 0 Then
                Err.Clear
                GetInfoByOrderNo = "0|����Ԥ���Ż�ȡ��ϸ��Ϣ��ʱ"
                Exit Function
            End If
            
            body1 = BytesToBstr(http.ResponseBody, 2)
            jsonorder2 = body1
            
            Call SavePage(body1, "jsonorder")
        
        End If
    
        txt_zcdd.Text = parseJSON(jsonorder2, "zcdd", 1)(0)
        txt_xcdd.Text = parseJSON(jsonorder2, "xcdd", 1)(0)
        txt_dzyx.Text = parseJSON(jsonorder2, "xcdddm", 1)(0)
        txt_fzyx.Text = parseJSON(jsonorder2, "zcdddm", 1)(0)
        
        If parseJSON(jsonorder2, "shdwdh", 1)(0) <> "" Then
            chk_dddxtz.Value = 1
            txt_shdwdh.Text = parseJSON(jsonorder2, "shdwdh", 1)(0)
        Else
           chk_dddxtz.Value = 0
           txt_shdwdh.Text = ""
        End If
            
        GetInfoByOrderNo = "1|������д���"
        
    End If

End Function

'Ԥ�ύ
Function PerPost() As String

    On Error Resume Next

    Dim surl As String, param As String
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_add"
    param = ""
    param = param & "currentPosition=" & "%E9%A2%84%E7%BA%A6%C2%A0%3E%3E%C2%A0%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    param = param & "&" & "djm=" & URLEncodeUTF8(txt_djm.Text)
    param = param & "&" & "dzhzzm=" & URLEncodeUTF8(txt_dzhzzm.Text)
    param = param & "&" & "dztmism=" & txt_dztmism.Text
    param = param & "&" & "dzyx=" & Replace(txt_dzyx.Text, " ", "+")
    param = param & "&" & "fhdwmc=" & URLEncodeUTF8(txt_fhdwmc.Text)
    param = param & "&" & "fjm=" & URLEncodeUTF8(txt_fjm.Text)
    param = param & "&" & "fzhzzm=" & URLEncodeUTF8(txt_fzhzzm.Text)
    param = param & "&" & "fztmism=" & txt_fztmism.Text
    param = param & "&" & "fzyx=" & Replace(txt_fzyx.Text, " ", "+")
    param = param & "&" & "hzpm=" & URLEncodeUTF8(txt_hzpm.Text)
    param = param & "&" & "keyword="
    param = param & "&" & "maxDate=" & Format(DateAdd("m", 1, Now()) - 1, "yyyy-mm-dd")
    param = param & "&" & "minDate=" & Format(Now(), "yyyy-mm-dd")
    param = param & "&" & "po.dddxtz=" & chk_dddxtz.Value
    param = param & "&" & "po.hqhw=" & txt_hqhw.Text
    param = param & "&" & "po.pzycfh=" & txt_pzycfh.Text
    param = param & "&" & "po.qqcs=" & txt_qqcs.Text
    param = param & "&" & "po.qqcz=" & Right(cbo_qqcz.Text, 1)
    param = param & "&" & "po.qqds=" & txt_qqds.Text
    param = param & "&" & "po.qqlx=0"
    param = param & "&" & "po.shdwdh=" & txt_shdwdh.Text
    param = param & "&" & "po.uuid="
    param = param & "&" & "po.xqslh=" & txt_xqslh.Text
    param = param & "&" & "po.zcrq=" & Trim(txt_zcrq.Text)
    param = param & "&" & "qqcsMax=" & txt_qqcsMax.Text
    param = param & "&" & "shdwmc=" & URLEncodeUTF8(txt_shdwmc.Text)
    param = param & "&" & "xcdd=" & URLEncodeUTF8(txt_xcdd.Text)
    param = param & "&" & "zcdd=" & URLEncodeUTF8(txt_zcdd.Text)
    
    
    Call SavePage("[" & Now() & ":step1]" & param & vbLf, "perpostdata")
    
    http.Open "POST", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        PerPost = "0|����Ԥ�ύ��ʱ"
        Exit Function
    End If
    
    body1 = BytesToBstr(http.ResponseBody, 2)
    
    Call SavePage("[" & Now() & ":step1]" & body1 & vbLf, "pergetdata")
    
    If InStr(body1, """success"":true") Then
        uuid = mySubstr(body1, "uuid"":""", """")
        If uuid <> "" Then
            PerPost = "1|Ԥ�ᱨ�ɹ�"
            Exit Function
        Else
            PerPost = "0|��ȡuuidʧ��"
            Exit Function
        End If
    ElseIf InStr(body1, "������Ԥ�����ڷ�Χ") Then
        PerPost = "0|������Ԥ�����ڷ�Χ"
        Exit Function
    ElseIf InStr(body1, "δ�ҵ���Ӧ��������Ϣ") Then
        PerPost = "0|δ�ҵ���Ӧ��������Ϣ"
        Exit Function
    Else
        PerPost = "0|Ԥ�ᱨʧ��"
        Exit Function
    End If


End Function


'��ʽ�ᱨ
Function RePost() As String

    On Error Resume Next

    Dim surl As String, param As String
    
    param = "op=10&uuids=" & uuid & ",&mor_dzsw_security_info=mor_dzsw_security_disabled"
    Call SavePage("[" & Now() & ":step2]" & param & vbLf, "perpostdata")
    
    surl = "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_operateZcrbjh"
    
    http.Open "POST", surl, False
    http.SetRequestHeader "Connection", "Keep-Alive"
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.SetRequestHeader "Host", "frontier." & city & ".12306.cn"
    http.SetRequestHeader "Accept", "application/json, text/javascript, */*"
    http.SetRequestHeader "Cookie", "BIGipServerhyswpt_pool=" & sen
    http.SetRequestHeader "Cookie", "DZSW_SESSIONID=" & sen2
    http.SetRequestHeader "Cookie", "CASTGC=" & sen3
    http.SetRequestHeader "Referer", "https://frontier." & city & ".12306.cn/gateway/hydzsw" & testurl2 & "/Dzsw/action/ZcrbjhAction_initAdd?currentPosition=%E9%A2%84%E7%BA%A6%26nbsp%3B%3E%3E%26nbsp%3B%E8%AE%A2%E7%A9%BA%E8%BD%A6"
    http.SetRequestHeader "Content-Length", Len(param)
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.Send param
    
    If Err.Number <> 0 Then
        Err.Clear
        RePost = "0|������ʽ�ᱨ��ʱ"
        Exit Function
    End If
    
    body2 = BytesToBstr(http.ResponseBody, 2)
    Call SavePage("[" & Now() & ":step2]" & body2 & vbLf, "pergetdata")
    
    If InStr(body2, """success"":true") Then
        RePost = "1|������ʽ�ᱨ�ɹ�"
        Exit Function
    Else
        RePost = "0|������ʽ�ᱨʧ��"
        Exit Function
    End If

End Function



'**************************************************����������*********************************************************
Private Sub Command4_Click()
    Form1.Width = 5595
End Sub


Private Sub Option2_Click(Index As Integer)
    If Index = 1 Then
        Form1.Width = 15705
    Else
        Form1.Width = 5595
    End If
End Sub



Private Sub txt_qqcs_Change()
    If txt_qqcs.Text <> "" And IsNumeric(txt_qqcs.Text) = True Then
        txt_qqds.Text = txt_qqcs.Text * 60
    End If
End Sub

Sub showinfo(Result As Integer, info As String)
    If Result = 1 Then   '�ɹ�
        lblInfo.ForeColor = &HD000&
        lblInfo.Caption = info
    ElseIf Result = 2 Then 'ʧ��
        lblInfo.ForeColor = &HFF&
        lblInfo.Caption = info
    ElseIf Result = 0 Then '������
        lblInfo.ForeColor = &HFFFF&
        lblInfo.Caption = info
    ElseIf Result = 3 Then '��ʾ��Ϣ
        lblInfo.ForeColor = &HC00000
        lblInfo.Caption = info
    End If
    
    Form1.Refresh
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        txt_zcrqper.Text = Trim(Format(Now(), "yyyy-mm-dd"))
    Else
        txt_zcrqper.Text = Trim(Format(DateAdd("d", 31, Now()), "yyyy-mm-dd"))
    End If
    
End Sub

'�����˺�
Sub saveAccount(user As String, pass As String)

    Dim tout As String, tin As String, flag As Boolean
    tou = ""
    tin = ""
    flag = False
    
    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(App.Path & "/dat/dat000.dat") = False Then
        Fso.CreateTextFile (App.Path & "/dat/dat000.dat")
    End If
    
    
    Open App.Path & "/dat/dat000.dat" For Input As #1
        Do While Not EOF(1)
            Line Input #1, tin
            If mySubstr(tin, "u=", ";") = user Then
               tout = tout & "u=" & user & ";p=" & pass & ";" & Chr(13) & Chr(10)
               flag = True
            Else
               If Len(tin) > 4 Then tout = tout & tin & vbCrLf
            End If
        Loop
    Close #1
    
    If flag = False Then
        tout = tout & "u=" & user & ";p=" & pass & ";" & vbCrLf
    End If
    
    Open App.Path & "/dat/dat000.dat" For Output As #1
        Print #1, tout;
    Close #1
    
End Sub


'��ȡ�˺�
Sub bindAccount(Optional user As String = "")

    Dim tout As String, tin As String, flag As Boolean
    tou = ""
    tin = ""
    flag = False
    
    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(App.Path & "/dat/dat000.dat") = False Then
        Exit Sub
    End If
    
    If user = "" Then
    
        Open App.Path & "/dat/dat000.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, tin
                If Len(tin) > 4 Then
                   txtUsername.AddItem (mySubstr(tin, "u=", ";"))
                End If
            Loop
        Close #1
        txtUsername.ListIndex = txtUsername.ListCount - 1
        
    Else
    
        Open App.Path & "/dat/dat000.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, tin
                If Len(tin) > 4 And mySubstr(tin, "u=", ";") = user Then
                   txtPassWord.Text = mySubstr(tin, "p=", ";")
                End If
            Loop
        Close #1
    
    End If
    
End Sub


Sub saveProfile(filename As String)

    Dim tout As String
    
    filename = App.Path & "/dat/" & filename & ".dat"

    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(filename) = False Then
        Fso.CreateTextFile (filename)
    End If
    
    tout = ""
    tout = tout & "user=" & Trim(txtUsername.Text) & "" & vbCrLf
    tout = tout & "pass=" & Trim(txtPassWord.Text) & "" & vbCrLf
    tout = tout & "comp=" & Mid(Label6.Caption, 5, Len(Label6.Caption) - 5) & "" & vbCrLf
    tout = tout & "zcrq=" & Trim(txt_zcrq.Text) & "" & vbCrLf
    tout = tout & "jsel=" & JsonselIndex & vbCrLf
    tout = tout & "jod1=" & jsonorder & "" & vbCrLf
    tout = tout & "jod2=" & jsonorder2 & "" & vbCrLf
   

    Open filename For Output As #1
        Print #1, tout;
    Close #1
    
End Sub


Sub loadProfile(filename As String)

    Dim tout As String, tin As String, tmpStr As String
    
    filename = App.Path & "/dat/" & filename & ".dat"

    Dim Fso As New Scripting.FileSystemObject
    
    If Fso.FileExists(filename) = False Then
        Call showinfo(2, "û���ҵ���Ӧ�Ķ����ļ�,����ʧ��!")
        Exit Sub
    End If
    
    

    Open filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, tin
            
            If Left(tin, 4) = "user" Then
                txtUsername.Text = Right(tin, Len(tin) - 5)
                ISOFFLINE = True
                
            ElseIf Left(tin, 4) = "pass" Then
                txtPassWord.Text = Right(tin, Len(tin) - 5)
                
            ElseIf Left(tin, 4) = "comp" Then
                Label6.ForeColor = RGB(0, 0, 255)
                Label6.Caption = "���߶���(" & Right(tin, Len(tin) - 5) & ")"
                
            ElseIf Left(tin, 4) = "zcrq" Then
                txt_zcrqper.Text = Right(tin, Len(tin) - 5)
                
            ElseIf Left(tin, 4) = "jsel" Then
                JsonselIndex = Right(tin, Len(tin) - 5)
                
            ElseIf Left(tin, 4) = "jod1" Then
    
                jsonorder = Right(tin, Len(tin) - 5)
                
                txt_orderlist.Clear
               
                tmpStr = ""
                tmpStr = tmpStr & parseJSON(jsonorder, "XQSLH", JsonselIndex + 1)(0) & "("
                tmpStr = tmpStr & parseJSON(jsonorder, "DZHZZM", JsonselIndex + 1)(0) & "|"
                tmpStr = tmpStr & parseJSON(jsonorder, "SHDWMC", JsonselIndex + 1)(0) & "|"
                tmpStr = tmpStr & parseJSON(jsonorder, "HZPM", JsonselIndex + 1)(0) & "|"
                tmpStr = tmpStr & parseJSON(jsonorder, "CZ", JsonselIndex + 1)(0) & "|"
    
                tmpStr = tmpStr & (CLng(parseJSON(jsonorder, "PZCS", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "JDZC4", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "YPWZ", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "YQWP", JsonselIndex + 1)(0)) - CLng(parseJSON(jsonorder, "FACS", JsonselIndex + 1)(0))) & ")"
            
                txt_orderlist.AddItem tmpStr
                txt_orderlist.Locked = True
                txt_orderlist.ListIndex = 0
                txt_orderlist.Enabled = False
                
            ElseIf Left(tin, 4) = "jod2" Then
            
                jsonorder2 = Right(tin, Len(tin) - 5)
                Call GetInfoByOrderNo(JsonselIndex, "offline")
                
                Option2(1).Value = True
                
            Else
                
            End If
        Loop
        
    Close #1
    
End Sub


Sub lockAll()
    cmdAuto.Enabled = False
    cmdDeAuto.Enabled = True
    
    cmd_login.Enabled = False
    cmd_getorder.Enabled = False
    
    txt_orderlist.Enabled = False
    txt_zcrqper.Enabled = False
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    Command3.Enabled = False
    Form1.Width = 5595
    Option2(0).Enabled = False
    Option2(1).Enabled = False
    
    cmd_profile.Enabled = False
    txt_profile.Enabled = False
    
    Txt_AllowAuto.Enabled = False
    
    
End Sub

Sub unlockAll()

    cmdAuto.Enabled = True
    cmdDeAuto.Enabled = False
    
    cmd_login.Enabled = True
    cmd_getorder.Enabled = True
    
    txt_orderlist.Enabled = True
    txt_zcrqper.Enabled = True
    Option1(0).Enabled = True
    Option1(1).Enabled = True
    Command3.Enabled = True
    Form1.Width = 5595
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    
    cmd_profile.Enabled = True
    txt_profile.Enabled = True
    
    Txt_AllowAuto.Enabled = True
    
End Sub




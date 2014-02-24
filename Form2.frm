VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3105
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   3105
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   480
      Left            =   1920
      TabIndex        =   2
      Top             =   2010
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   1560
      Width           =   2700
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   255
      ScaleHeight     =   1035
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   345
      Width           =   2715
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

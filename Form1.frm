VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7320
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   990
      TabIndex        =   4
      Top             =   225
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   2970
      TabIndex        =   3
      Top             =   225
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   4815
      TabIndex        =   2
      Top             =   225
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3210
      Left            =   45
      TabIndex        =   0
      Top             =   1170
      Width           =   7215
      Begin VB.TextBox Text1 
         Height          =   2985
         Left            =   45
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   7125
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'开发软件时候,把这个modal装入程序中.然后加入如下代码:
Private Sub Form_Load()
    Call ResizeInit(Me) '在程序装入时必须加入
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me) '确保窗体改变时控件随之改变
End Sub

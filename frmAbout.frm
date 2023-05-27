VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于..."
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3645
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdReadMe 
      Caption         =   "Read Me"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblExeName 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0152
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdReadMe_Click()
    Shell "cmd /c notepad """ & App.Path & "\ReadMe.txt""", vbHide
End Sub

Private Sub Form_Load()
    Caption = "关于" + App.EXEName
    lblExeName = App.ProductName & " Ver " & App.Major & "." & App.Minor & "." & App.Revision
    lblInfo = "Made by Dangfer" & vbCrLf & "弱者随手写之，君用而笑之。"
End Sub


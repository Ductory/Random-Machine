VERSION 5.00
Begin VB.UserControl CheckList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   HasDC           =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   1695
   ToolboxBitmap   =   "CheckList.ctx":0000
   Begin VB.VScrollBar vscBar 
      Height          =   2415
      Left            =   1440
      Max             =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox chkItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "New Item"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu mnuPop 
      Caption         =   "POP"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAll 
         Caption         =   "全选"
      End
      Begin VB.Menu mnuDeselectAll 
         Caption         =   "取消全选"
      End
      Begin VB.Menu mnuReverseSelect 
         Caption         =   "反选"
      End
   End
End
Attribute VB_Name = "CheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CheckBox Ver 0.0.0
Option Explicit

Dim itemsCount As Long '记录Items总数
Dim itemsShow As Long '记录可视的Items个数


Public Sub AddItem(ByVal newItem As String)
    itemsCount = itemsCount + 1
    '原先无Items
    If itemsCount = 1 Then
        chkItem(0).Caption = newItem
        chkItem(0).Visible = True
        Exit Sub
    End If
    '原先已有
    Load chkItem(itemsCount - 1)
    With chkItem(itemsCount - 1)
        .Caption = newItem
        .Top = chkItem(itemsCount - 2).Top + 240
        .Visible = True
    End With
    '更新ScrollBar
    If itemsCount - itemsShow >= 0 Then
        vscBar.Max = itemsCount - itemsShow
    Else
        vscBar.Max = 0
    End If
End Sub

Public Sub Clear()
    Dim i As Long
    chkItem(0).Visible = False
    For i = 1 To itemsCount - 1
        Unload chkItem(i)
    Next
    itemsCount = 0
End Sub

Public Property Get Count() As Long
    Count = itemsCount
End Property

Public Property Get Checked(ByVal Index As Long) As Boolean
    Checked = chkItem(Index).Value
End Property

Public Property Let Checked(ByVal Index As Long, pChecked As Boolean)
    chkItem(Index).Value = -pChecked
End Property

Public Property Get Item(ByVal Index As Long) As String
    Item = chkItem(Index).Caption
End Property

Public Property Let Item(ByVal Index As Long, strCaption As String)
    chkItem(Index).Caption = strCaption
End Property


'Private Sub chkItem_Click(Index As Integer)
'    bReset = True
'    '不建议在用户控件中使用全局变量，因为这会影响控件的封装性，建议写成
'    'Public Event Change(ByVal Index As Long) '在通用区声明一个事件
'    'RaiseEvent Change(Index) '激发事件
'End Sub

Private Sub chkItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then PopupMenu mnuPop
End Sub

Private Sub mnuDeselectAll_Click()
    Dim i As Long
    For i = 0 To itemsCount - 1
        chkItem(i).Value = 0
    Next
End Sub

Private Sub mnuReverseSelect_Click()
    Dim i As Long
    For i = 0 To itemsCount - 1
        chkItem(i).Value = 1 - chkItem(i).Value
    Next
End Sub

Private Sub mnuSelectAll_Click()
    Dim i As Long
    For i = 0 To itemsCount - 1
        chkItem(i).Value = 1
    Next
End Sub

Private Sub UserControl_Resize()
    Dim i As Long
    itemsShow = (Height - 20) \ 240
    vscBar.Left = Width - 255
    vscBar.Height = Height
    chkItem(0).Width = Width - 255
End Sub

Private Sub vscBar_Change()
    vscBar_Scroll
End Sub

Private Sub vscBar_Scroll()
    On Error Resume Next
    Dim i As Long
    For i = 0 To itemsShow
        chkItem(vscBar.Value + i).Top = i * 240 + 20
        chkItem(vscBar.Value + i).ZOrder
    Next
End Sub

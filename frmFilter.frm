VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "筛选器"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin RandomMachine.CheckList cklKey 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1815
      _extentx        =   3201
      _extenty        =   2355
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "筛选"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtMatch 
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.OptionButton optKey 
      Caption         =   "包含"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.OptionButton optKey 
      Caption         =   "存在"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   300
      ItemData        =   "frmFilter.frx":0000
      Left            =   120
      List            =   "frmFilter.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblStringMatch 
      Caption         =   "字符串匹配"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblNumberMatch 
      Caption         =   "数值匹配"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'hWndInsertAfter 参数可选值:
Private Const HWND_TOPMOST = -1 ' {在前面, 位于任何顶部窗口的前面}
'wFlags 参数可选值:
Private Const SWP_NOSIZE = 1 ' {忽略 cx、cy, 保持大小}
Private Const SWP_NOMOVE = 2 ' {忽略 X、Y, 不改变位置}

Dim KeyStyle As Long

Private Sub cmbFilter_Click() '该写法仅为了压缩代码量
    Dim i As Long
    i = cmbFilter.ListIndex
    lblStringMatch.Visible = (i = 0) 'Value
    lblNumberMatch.Visible = (i = 2) 'Weight
    txtMatch.Visible = (i <> 1) 'Value 或 Weight
    optKey(0).Visible = (i = 1): optKey(1).Visible = (i = 1): cklKey.Visible = (i = 1) 'Key
    cmdRun.Top = 1080 - 1080 * (i = 1) '位置调整
End Sub

'筛选器
Private Sub cmdRun_Click()
    On Error GoTo Err
    Dim matchKey() As String
    Dim i As Long, j As Long, k As Long
    Dim KeyTable As String, sKey() As String
    '添加项
    Dim tmpCheckList As CheckList
    Set tmpCheckList = frmMain.cklList
    For i = 0 To UBound(drawList)
        If tmpCheckList.Checked(i) Then '已勾选
            drawList(j) = i: j = j + 1 '索引加入，设置抽取项
        End If
    Next
    
    Select Case cmbFilter.ListIndex
    Case 0 'Value
        Dim s As String
        s = txtMatch.Text
        For i = 0 To UBound(drawList)
            '若不匹配，则取消选中
            If Not items(drawList(i)).Value Like s Then tmpCheckList.Checked(drawList(i)) = False
        Next
    Case 1 'Key
        '关键字加入
        j = 0
        If cklKey.Count = 0 Then Exit Sub '若未选择关键字则退出
        ReDim matchKey(cklKey.Count - 1)
        For i = 0 To UBound(matchKey)
            If cklKey.Checked(i) Then '将选中的Key入表
                matchKey(j) = cklKey.Item(i) '索引加入
                KeyTable = KeyTable & "," & matchKey(j) '构建关键字表
                j = j + 1
            End If
        Next
        ReDim Preserve matchKey(j - 1)
        KeyTable = KeyTable & ","

        For i = 0 To tmpCheckList.Count - 1
            If Not tmpCheckList.Checked(i) Then GoTo Continue '获取下一个选中项
            '关键字样式
            '0存在 1包含
            If KeyStyle = 0 Then
                sKey = Split(items(i).Key, ",") '取出关键字
                '将matchKey与Item的关键字逐一匹配
                For k = 0 To UBound(sKey)
                    If InStr(KeyTable, "," & sKey(k) & ",") Then GoTo Continue '存在
                Next
                tmpCheckList.Checked(i) = False '未匹配到结果
            Else
                For j = 0 To UBound(matchKey)
                    '将items().Key转为 ",K1,K2,...,Kn," 的形式
                    '将 Ki 转为 ",Ki," 的形式
                    If InStr("," & items(i).Key & ",", "," & matchKey(j) & ",") = 0 Then tmpCheckList.Checked(i) = False: Exit For '未匹配到结果
                Next
            End If
Continue:
        Next
    Case 2
        For i = 0 To UBound(drawList)
            If Not vbs.Eval(items(drawList(i)).Weight & txtMatch.Text) Then tmpCheckList.Checked(drawList(i)) = False
        Next
    End Select
Err:
End Sub

Private Sub Form_Load()
    '列表初始化
    cmbFilter.AddItem "值"
    cmbFilter.AddItem "关键字"
    cmbFilter.AddItem "权"
    '读取配置
    cmbFilter.ListIndex = getIniInt(StrPtr("Config"), StrPtr("Filter"), 0&, lpConfig)
    KeyStyle = getIniInt(StrPtr("Config"), StrPtr("KeyFilter"), 0&, lpConfig)
    optKey(KeyStyle).Value = True
    '顶层窗口
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1 '防止窗体被删除
    Me.Hide
End Sub

Private Sub optKey_Click(Index As Integer)
    KeyStyle = Index
End Sub

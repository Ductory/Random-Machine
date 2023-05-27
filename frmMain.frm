VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Machine"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6975
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6975
   StartUpPosition =   2  '屏幕中心
   Begin RandomMachine.CheckList cklList 
      Height          =   3495
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   6165
   End
   Begin VB.Frame frmDraw 
      Caption         =   "抽取界面"
      Height          =   2055
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdDraw 
         Caption         =   "抽取"
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdDrawBatch 
         Caption         =   "批量抽取"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.Frame frmDrawList 
         Height          =   975
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2895
         Begin VB.Label lblShow 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "幼圆"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.OptionButton optDraw 
         Caption         =   "抽取随机数"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optDraw 
         Caption         =   "抽取名单"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame frmDrawRandNum 
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2895
         Begin VB.CheckBox chkSort 
            Caption         =   "升序"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chkOnly 
            Caption         =   "唯一"
            Height          =   255
            Left            =   840
            TabIndex        =   18
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtH 
            Height          =   270
            Left            =   2040
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtL 
            Height          =   270
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblH 
            Caption         =   "上界"
            Height          =   255
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblL 
            Caption         =   "下界"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.ListBox lstShow 
      Appearance      =   0  'Flat
      Height          =   3450
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame frmSetting 
      Caption         =   "设置"
      Height          =   1575
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
      Begin VB.CheckBox chkAddWeight 
         Caption         =   "权增"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "右键设置权增值"
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame frmBatch 
         Height          =   975
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   2055
         Begin VB.TextBox txtGroupName 
            Height          =   270
            Left            =   600
            TabIndex        =   13
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtBatch 
            Height          =   270
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblGroupName 
            Caption         =   "组名"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblBatch 
            Caption         =   "批量抽取数量"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkPowerWeight 
         Caption         =   "幂权"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "右键设置底数"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkAutoRemove 
         Caption         =   "自动屏蔽"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkWeight 
         Caption         =   "允许加权"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblListShow 
      Caption         =   "抽取一览"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblPreList 
      Caption         =   "列表预览"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开文件(O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistoryP 
         Caption         =   "历史记录"
         Begin VB.Menu mnuHistory 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuHistory 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuHistory 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuHistory 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuHistory 
            Caption         =   ""
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuExtend 
      Caption         =   "扩展"
      Begin VB.Menu mnuEx 
         Caption         =   "扩展名"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助"
      Begin VB.Menu mnuAbout 
         Caption         =   "关于..."
      End
   End
   Begin VB.Menu mnuListFile 
      Caption         =   "文件名"
      Visible         =   0   'False
      Begin VB.Menu mnuFileInfo 
         Caption         =   "文件信息"
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "筛选器"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuLstShow 
      Caption         =   "POP_LSTSHOW"
      Visible         =   0   'False
      Begin VB.Menu mnuShowClear 
         Caption         =   "清空"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "导出"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text '不区分大小写
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ALIVE = &H103

Private Type EXTEND
    Command As String
    ExitCode As Boolean
    FlagEx As Boolean
    Startup As Boolean
End Type

Dim Extends() As EXTEND '扩展程序
Dim cdlFile As New CommonDialog
Dim batch As Long '批处理
Dim history(1 To 5) As String '历史记录
Dim drawFlag As Boolean '抽取按钮 抽取(T)/停止(F) 状态
Dim KeyStyle As Byte '存在/包含

Dim powerWeight As Single, addWeight As Long '幂权，权增
Public FlagEx As Boolean '扩展标志
Public RMBuff As String '脚本缓冲区
'抽取主过程
Public Function draw() As Long
    Dim i As Long, r As Long, tmp As Long
    '若抽取列表中无项，或都是权重为0的项，则退出
    If drawCount = 0 Or (bWei And wSUM = 0) Then
        MsgBox$ "列表已经抽取完毕!", vbExclamation, "警告"
        drawFlag = False
        draw = -1
        Exit Function
    End If
    If chkWeight.Value Then '加权
        If chkPowerWeight.Value Then '幂权
            Do
                i = Int(Rnd() * drawCount)
            Loop While powerWeight ^ (items(drawList(i)).Weight - wMIN) <= Rnd()
        Else '普通权
            If 2 * drawCount * wMAX / wSUM <= (drawCount + 1) / 2 Then '使用二步加权抽取
                Do
                    i = Int(Rnd() * drawCount) '均分抽
                Loop While items(drawList(i)).Weight / wMAX <= Rnd()
            Else '遍历抽取
                r = Int(Rnd() * wSUM)
                tmp = 0
                For i = 0 To drawCount - 1
                    tmp = tmp + items(drawList(i)).Weight
                    If r < tmp Then Exit For
                Next
            End If
        End If
    Else '普通抽取
        i = Int(Rnd() * drawCount)
    End If
    draw = i
End Function
'更新抽取项
Public Sub reset()
    Dim i As Long, j As Long
    wSUM = 0
    For i = 0 To UBound(items)
        '加入
        If cklList.Checked(i) Then
            If bWei Then '加权
                wSUM = wSUM + items(i).Weight
            End If
            drawList(j) = i
            j = j + 1
        End If
    Next
    drawCount = j
End Sub

'读取文件
Private Sub OpenFile(FileName As String)
    Dim sLines() As String, sec() As String, i As Long, j As Long, b() As Byte
    Dim KeyTable As String
    '显示文件名
    mnuListFile.Caption = Mid$(FileName, InStrRev(FileName, "\") + 1)
    mnuListFile.Visible = True
    '置初值
    bKey = False: bWei = False
    wSUM = 0: wMAX = 0: wMIN = &H7FFFFFFF
    '清空
    lstShow.Clear
    cklList.Clear
    frmFilter.cklKey.Clear
    '读取文件
    Open FileName For Binary As #1
        b = Input(LOF(1), #1)
        '读取预抽取头
        Do While b(i) <> vbKeyReturn
            Select Case b(i)
            Case vbKeyK: bKey = True
            Case vbKeyW: bWei = True
            End Select
            i = i + 1
        Loop
        '读取items
        sLines = Split(b, vbCrLf)
        For i = 1 To UBound(sLines) '排除第一行的预抽取头
            If Len(sLines(i)) Then '项非空
                sec = Split(sLines(i))
                i = i - 1 '先-1后+1降低取址时间开支
                ReDim Preserve items(i)
                '判断Key
                If bKey Then
                    If Len(sec(1)) Then KeyTable = KeyTable + sec(1) + "," '若存在关键字则加入
                    items(i).Key = sec(1)
                End If
                '处理Weight
                If bWei Then
                    items(i).Weight = sec(1 - bKey) '若有Key，则为sec(2)，否则为sec(1)
                    wSUM = wSUM + items(i).Weight
                    If items(i).Weight < wMIN Then
                        wMIN = items(i).Weight
                    ElseIf items(i).Weight > wMAX Then
                        wMAX = items(i).Weight
                    End If
                End If
                '写item
                items(i).Value = sec(0)
                cklList.AddItem sec(0): cklList.Checked(i) = True
                i = i + 1
            End If
        Next
        drawCount = UBound(items) + 1
        ReDim drawList(drawCount - 1)
        reset
        
        '添加关键字到筛选器
        Dim tmpCheckList As CheckList
        Set tmpCheckList = frmFilter.cklKey
        tmpCheckList.Clear
        If bKey Then
            Dim k As Long
            '将KeyTable转为 ,K1,,K2,,K3,,...,,Kn, 的形式
            KeyTable = "," & Replace$(KeyTable, ",", ",,")
            KeyTable = Left$(KeyTable, Len(KeyTable) - 1) '去掉最后一个,
            i = 1
            Do
                k = InStr(i + 2, KeyTable, ",,") '找到下一个 ",Ki,"
                If k = 0 Then Exit Do '没有找到，说明已遍历完毕
                '取 ",Ki," 并将表中其他所有 ",Ki," 替换为""，实现去重
                KeyTable = Mid$(KeyTable, 1, k) + Replace$(KeyTable, Mid$(KeyTable, i, k - i + 1), "", k + 1)
                i = k + 1 '迭代
            Loop
            '分割Key，先去掉两旁的","，然后按",,"作为分隔符进行分割
            sec = Split(Mid$(KeyTable, 2, Len(KeyTable) - 2), ",,")
            For i = 0 To UBound(sec)
                tmpCheckList.AddItem sec(i) '加入关键字的CheckList
            Next
        End If
        Set tmpCheckList = Nothing '清除对象引用
        'chkWeight有效判断
        If bWei Then
            chkWeight.Enabled = True
        Else
            chkWeight.Value = 0: chkWeight.Enabled = False
        End If
    Close #1
End Sub
'更新历史记录
Private Sub UpdateHistory(ByVal FileName As String)
    Dim i As Long, j As Long, s As String
    For i = 1 To 5
        If FileName = history(i) Then
            '如果已存在该记录，则将其放到history(1)
            s = history(i)
            For j = i To 2 Step -1
                history(j) = history(j - 1)
            Next
            history(1) = s
            GoTo UpdateMenu '更新菜单
        End If
    Next
    '如果不存在，则新建后放到history(1)
    If Len(Dir(FileName)) <> 0 And Len(FileName) = 0 Then Exit Sub '不存在文件则退出
    For j = 4 To 1 Step -1 '历史记录迭代
        history(j + 1) = history(j)
    Next
    history(1) = FileName
UpdateMenu:
    For i = 1 To 5 '更新菜单
        If Len(history(i)) Then
            mnuHistory(i).Caption = history(i)
            mnuHistory(i).Enabled = True
        Else
            mnuHistory(i).Caption = "（无）"
            mnuHistory(i).Enabled = False
        End If
    Next
End Sub
'扩展程序
Private Function ShellEx(ByVal PID As Long, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
    Dim hProcess As Long, ExitCode As Long
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, PID)
    Do
        ShellEx = GetExitCodeProcess(hProcess, ExitCode)
        DoEvents
    Loop While ExitCode = STILL_ALIVE
    CloseHandle hProcess
End Function


'幂权
Private Sub chkPowerWeight_Click()
    chkAddWeight.Value = chkPowerWeight.Value
End Sub

Private Sub chkPowerWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then powerWeight = Val(InputBox$("请输入底数", "设置幂权", powerWeight))
    If powerWeight = 0 Then MsgBox$ "底数不能为0！", vbCritical, "Error": chkPowerWeight.Value = 0
End Sub
'权增
Private Sub chkAddWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then addWeight = Val(InputBox$("请输入权增值", "设置权增", addWeight))
End Sub

'窗体加载
Private Sub Form_Load()
    Dim i As Long, tmp As Long, s As String
    Dim lpTemp As Long, ApplicationName As String
    Dim buff As String
    
    Randomize '随机种子发生器
    'vbs初始化
    vbs.Language = "VBScript"
    vbs.AddObject "RM", Me, True
    '读取配置文件
    '注意：以下的指针操作为危险操作，因为存在对临时字符串进行取址，这可能导致不可预测的结果
    '获取Config路径指针
    ConfigName = App.Path & "\Config.ini"
    lpConfig = StrPtr(ConfigName)
    '缓冲区大小
    ApplicationName = "Config": lpTemp = StrPtr(ApplicationName)
    buffSize = getIniInt(lpTemp, StrPtr("BufferSize"), 256&, lpConfig)
    buff = Space$(buffSize)
    '抽取方式
    optDraw(getIniInt(lpTemp, StrPtr("Draw"), 0&, lpConfig)).Value = True '设置OptionButton
    '全局CheckBox
    chkWeight.Value = getIniInt(lpTemp, StrPtr("WeightAllow"), 0&, lpConfig)
    chkAutoRemove.Value = getIniInt(lpTemp, StrPtr("RepeatAllow"), 0&, lpConfig)
    '批量处理
    batch = getIniInt(lpTemp, StrPtr("BatchNum"), 1&, lpConfig)
    '抽取随机数
    L = getIniInt(lpTemp, StrPtr("NumL"), 0&, lpConfig)
    H = getIniInt(lpTemp, StrPtr("numH"), 0&, lpConfig)
    chkOnly.Value = getIniInt(lpTemp, StrPtr("NumOnly"), 0&, lpConfig)
    chkSort.Value = getIniInt(lpTemp, StrPtr("NumSort"), 0&, lpConfig)
    '幂权
    '将Long读成Single
    tmp = getIniInt(lpTemp, StrPtr("PowerWeightNum"), 0&, lpConfig)
    CopyMemory powerWeight, tmp, 4
    chkPowerWeight.Value = getIniInt(lpTemp, StrPtr("PowerWeight"), 0&, lpConfig)
    '权增
    addWeight = getIniInt(lpTemp, StrPtr("AddWeightNum"), 0&, lpConfig)
    chkAddWeight.Value = getIniInt(lpTemp, StrPtr("AddWeight"), 0&, lpConfig)
    '扩展程序
    i = 0
    ApplicationName = "Extend": lpTemp = StrPtr(ApplicationName)
    Do
        '获取命令
        s = Left$(buff, getIniStr(lpTemp, StrPtr("E" & i + 1), 0&, StrPtr(buff), buffSize&, lpConfig))
        If Len(s) = 0 Then Exit Do '未找到扩展
        ReDim Preserve Extends(i)
        tmp = InStr(s, ")") + 1 '找到扩展标题开始位置
        If i > 0 Then Load mnuEx(i) '更新菜单
        mnuEx(i).Caption = Mid$(s, tmp, InStr(s, ":") - tmp) '将标题添加到菜单
        With Extends(i) '处理其他参数
            .Command = Mid$(s, InStr(s, ":") + 1)
            .ExitCode = (InStr(s, "E") <> 0)
            .FlagEx = (InStr(s, "F") <> 0)
            .Startup = (InStr(s, "S") <> 0)
        End With
        If Extends(i).Startup Then mnuEx_Click CInt(i)
        i = i + 1
    Loop
    '历史记录
    ApplicationName = "History": lpTemp = StrPtr(ApplicationName)
    For i = 1 To 5
        history(i) = Left$(buff, getIniStr(lpTemp, StrPtr("H" & i), 0&, StrPtr(buff), 256&, lpConfig))
    Next

    txtBatch.Text = batch
    txtL.Text = L: txtH.Text = H
    '启动筛选器
    Load frmFilter
    '打开最近文件
    If Not FlagEx Then mnuHistory_Click 1
    '如果打开脚本
    If Len(Command$) Then
        Open Command$ For Output As #2
            s = Input(LOF(2), #2)
        Close #2
        vbs.ExecuteStatement s
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim lpTemp As Long, ApplicationName As String
    ApplicationName = "Config": lpTemp = StrPtr(ApplicationName)
    setIniStr lpTemp, StrPtr("Draw"), StrPtr(CStr(optDraw(0).Value + 1)), lpConfig
    setIniStr lpTemp, StrPtr("WeightAllow"), StrPtr(CStr(chkWeight.Value)), lpConfig
    setIniStr lpTemp, StrPtr("RepeatAllow"), StrPtr(CStr(chkAutoRemove.Value)), lpConfig
    setIniStr lpTemp, StrPtr("BatchNum"), StrPtr(CStr(batch)), lpConfig
    setIniStr lpTemp, StrPtr("NumL"), StrPtr(CStr(L)), lpConfig
    setIniStr lpTemp, StrPtr("NumH"), StrPtr(CStr(H)), lpConfig
    setIniStr lpTemp, StrPtr("NumOnly"), StrPtr(CStr(chkOnly.Value)), lpConfig
    setIniStr lpTemp, StrPtr("NumSort"), StrPtr(CStr(chkSort.Value)), lpConfig
    setIniStr lpTemp, StrPtr("KeyFilter"), StrPtr(CStr(KeyStyle)), lpConfig
    setIniStr lpTemp, StrPtr("Filter"), StrPtr(CStr(frmFilter.cmbFilter.ListIndex)), lpConfig
    CopyMemory i, powerWeight, 4&
    setIniStr lpTemp, StrPtr("PowerWeightNum"), StrPtr(CStr(i)), lpConfig
    setIniStr lpTemp, StrPtr("PowerWeight"), StrPtr(CStr(chkPowerWeight.Value)), lpConfig
    setIniStr lpTemp, StrPtr("AddWeightNum"), StrPtr(CStr(addWeight)), lpConfig
    setIniStr lpTemp, StrPtr("AddWeight"), StrPtr(CStr(chkAddWeight.Value)), lpConfig

    ApplicationName = "History": lpTemp = StrPtr(ApplicationName)
    For i = 1 To 5
        setIniStr lpTemp, StrPtr("H" & i), StrPtr(history(i)), lpConfig
    Next i
    
    Open history(1) For Output As #1
        Print #1, IIf(bKey, "K", Empty) & IIf(bWei, "W", Empty)
        For i = 0 To UBound(items)
            Print #1, items(i).Value & IIf(bKey, " " & items(i).Key, Empty) & IIf(bWei, " " & items(i).Weight, Empty)
        Next
    Close #1
    End
End Sub
'抽取
Private Sub cmdDraw_Click()
    Dim i As Long
    reset
    If optDraw(0).Value Then
        drawFlag = Not drawFlag
        If drawFlag Then
            cmdDrawBatch.Enabled = False
            cmdDraw.Caption = "停止"
            Do
                i = draw()
                If i = -1 Then Exit Do '引发了异常
                lblShow.Caption = items(drawList(i)).Value
                DoEvents
            Loop While drawFlag
        
            If i <> -1 Then '没有异常引发
                lstShow.AddItem "抽取 " & items(drawList(i)).Value
                '权增
                If chkAddWeight.Value Then
                    items(drawList(i)).Weight = items(drawList(i)).Weight + addWeight
                    If items(drawList(i)).Weight > wMAX Then wMAX = items(drawList(i)).Weight
                End If
                '自动屏蔽已抽
                If chkAutoRemove.Value = 1 Then
                    '修正索引
                    cklList.Checked(drawList(i)) = False
                    If bWei Then wSUM = wSUM - items(drawList(i)).Weight
                    drawCount = drawCount - 1
                    drawList(i) = drawList(drawCount)
                End If
            End If
        End If
        If Not drawFlag Then '抽取时，drawFlag = true
            cmdDrawBatch.Enabled = True
            cmdDraw.Caption = "抽取"
        End If
    Else
        L = Val(txtL.Text): H = Val(txtH.Text)
        i = Int((H - L + 1) * Rnd) + L
        MsgBox$ i, vbInformation, "抽取结果"
        lstShow.AddItem "随机数 " & vbCrLf & i
    End If
End Sub

Private Sub cmdDrawBatch_Click()
    Dim i As Long, j As Long, tmp As Long, initCount As Long, s As String
    reset
    If optDraw(0).Value Then
        cmdDrawBatch.Enabled = False
        cmdDraw.Enabled = False
    
        s = txtGroupName.Text & " " & vbCrLf
        If Len(s) = 3 Then s = "未命名组" & s '组名为空
        initCount = drawCount - 1
        For i = 1 To batch
            j = draw()
            If j = -1 Then Exit For '引发了异常
            '权增
            If chkAddWeight.Value Then
                items(drawList(j)).Weight = items(drawList(j)).Weight + addWeight
                If items(drawList(j)).Weight > wMAX Then wMAX = items(drawList(j)).Weight
            End If
            s = s & items(drawList(j)).Value & " "
            '自动屏蔽已抽（没有异常引发时）
            If chkAutoRemove.Value Then
                cklList.Checked(drawList(j)) = False
                If bWei Then wSUM = wSUM - items(drawList(j)).Weight
                drawCount = drawCount - 1
                drawList(j) = drawList(drawCount)
            Else
                tmp = drawList(j) '交换drawList中第j个和最后1个
                drawList(j) = drawList(drawCount - 1)
                drawList(drawCount - 1) = tmp
                If bWei Then wSUM = wSUM - items(tmp).Weight
                drawCount = drawCount - 1
            End If
        Next
        If chkAutoRemove.Value = 0 Then '不自动屏蔽
            If bWei Then '重新加回SUM
                For i = drawCount To initCount
                    wSUM = wSUM + items(drawList(i)).Weight
                Next
            End If
            drawCount = initCount
        End If
        If i <> 1 Then lstShow.AddItem s: MsgBox$ s '没有抽出0个
        cmdDrawBatch.Enabled = True
        cmdDraw.Enabled = True
    Else
        L = Val(txtL.Text): H = Val(txtH.Text)
        If batch > H - L + 1 Then MsgBox "区间过小!", vbCritical, "Error"
        Dim result() As Long, hash() As Long
        ReDim result(batch - 1): ReDim hash(L To H)
        
        For i = 0 To batch - 1
            j = Int((H - L + 1) * Rnd) + L
            If chkOnly.Value Then
                Do While hash(j)
                    j = Int((H - L + 1) * Rnd) + L
                Loop
            End If
            result(i) = j
            hash(j) = hash(j) + 1
        Next
        If chkSort.Value Then
            For i = L To H
                For j = 1 To hash(i)
                    result(tmp) = i: tmp = tmp + 1
                Next
            Next
        End If
        For i = 0 To batch - 1
            s = s & result(i) & " "
        Next
        lstShow.AddItem "批量随机数 " & vbCrLf & s
        MsgBox$ s, vbInformation, "抽取结果"
    End If
End Sub

'展示列表
Private Sub lstShow_DblClick()
    MsgBox$ lstShow.List(lstShow.ListIndex)
End Sub

Private Sub lstShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then PopupMenu mnuLstShow
End Sub
'扩展
Private Sub mnuEx_Click(Index As Integer)
    mnuEx(Index).Checked = Not mnuEx(Index).Checked
    If mnuEx(Index).Checked Then
        If Extends(Index).FlagEx Then FlagEx = True
        Dim s As String
        '替换系统变量
        s = Extends(Index).Command
        s = Replace$(s, "%HWND%", Me.hwnd)
        mnuEx(Index).Tag = Shell(App.Path & "\Ex\" & s)
        If Extends(Index).ExitCode Then
            MsgBox$ "扩展程序""" & mnuEx(Index).Caption & """已终止运行" & vbCrLf _
            & "退出代码: " & ShellEx(mnuEx(Index).Tag) '运行扩展程序
        End If
    Else
        Shell "cmd /c taskkill /f /pid " & mnuEx(Index).Tag, vbHide
    End If
End Sub

'导出
Private Sub mnuExport_Click()
    If cdlFile.showSave(hwnd) Then
        MsgBox cdlFile.FilePath
        Open cdlFile.FilePath For Output As #1
            Dim s As String, i As Long
            For i = 0 To lstShow.ListCount - 1
                s = s & lstShow.List(i) & vbCrLf
            Next
            Print #1, s
        Close #1
        Shell "cmd /c notepad " & cdlFile.FilePath, vbHide
    Else
        MsgBox$ "无法导出!", vbCritical, "Error"
    End If
End Sub
'筛选器
Private Sub mnuFilter_Click()
    frmFilter.Show
End Sub
'文件信息
Private Sub mnuFileInfo_Click()
    Dim s As String
    s = "关键字：" & vbTab & IIf(bKey, "是", "否") & vbCrLf & _
        "权重：" & vbTab & IIf(bKey, "是", "否") & vbCrLf & _
        "项数：" & vbTab & UBound(items) + 1
    MsgBox$ s, vbInformation, "文件信息"
End Sub


'清空展示列表
Private Sub mnuShowClear_Click()
    lstShow.Clear
End Sub
'打开文件
Private Sub mnuOpen_Click()
    cdlFile.showOpen Me.hwnd
    '由于输入的路径不一定有效，所以要加上路径判断
    '由于Dir("")为查找当前目录下一个文件，所以要加上空字符串判断
    If Len(Dir(cdlFile.FilePath)) <> 0 And Len(cdlFile.FilePath) <> 0 Then
        OpenFile cdlFile.FilePath
        UpdateHistory cdlFile.FilePath
    Else
        MsgBox$ "文件不存在!", vbCritical, "Error"
    End If
End Sub
'打开历史记录
Private Sub mnuHistory_Click(Index As Integer)
    If Len(history(Index)) <> 0 Then '文件存在
        If Len(Dir(history(Index))) <> 0 Then
            OpenFile history(Index)
        Else
            MsgBox "文件不存在!", vbCritical, "Error"
            history(Index) = Empty
        End If
    End If
    UpdateHistory history(Index)
End Sub
'关于
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub
'抽取模式 0:名单 1:随机数
Private Sub optDraw_Click(Index As Integer)
    frmDrawList.Visible = (Index = 0)
End Sub

'批量抽取设置
Private Sub txtBatch_Change()
    batch = Val(txtBatch.Text) '加Val()防止为空时报错
End Sub

'RM脚本
Public Function extract(ByVal times As Long, ByVal bWeight As Boolean, Optional groupName As String = "未命名组") As String
    extract = IIf(Len(groupName), groupName, "未命名组") & vbCrLf
    Dim i As Long, j As Long, tmp As Long
    
    For i = 1 To times
        j = draw()
        If j = -1 Then Exit For '引发了异常
        '修正SUM
        If bWei Then wSUM = wSUM - items(drawList(j)).Weight
        extract = extract & items(drawList(j)).Value & " "
        '自动屏蔽已抽（没有异常引发时）
        drawList(j) = drawList(drawCount - 1)
    Next
    RMBuff = RMBuff & extract & vbCrLf
End Function

Public Function export(exportPath As String)
    If Len(exportPath) = 0 Then MsgBox$ "未指定导出路径！", vbCritical, "脚本错误"
    Open exportPath For Output As #2
        Print #2, RMBuff
    Close #2
End Function

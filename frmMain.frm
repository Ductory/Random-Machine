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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "��ȡ����"
      Height          =   2055
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdDraw 
         Caption         =   "��ȡ"
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdDrawBatch 
         Caption         =   "������ȡ"
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
               Name            =   "��Բ"
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
         Caption         =   "��ȡ�����"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optDraw 
         Caption         =   "��ȡ����"
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
            Caption         =   "����"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chkOnly 
            Caption         =   "Ψһ"
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
            Caption         =   "�Ͻ�"
            Height          =   255
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblL 
            Caption         =   "�½�"
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
      Caption         =   "����"
      Height          =   1575
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
      Begin VB.CheckBox chkAddWeight 
         Caption         =   "Ȩ��"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "�Ҽ�����Ȩ��ֵ"
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
            Caption         =   "����"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblBatch 
            Caption         =   "������ȡ����"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkPowerWeight 
         Caption         =   "��Ȩ"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "�Ҽ����õ���"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkAutoRemove 
         Caption         =   "�Զ�����"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkWeight 
         Caption         =   "�����Ȩ"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblListShow 
      Caption         =   "��ȡһ��"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblPreList 
      Caption         =   "�б�Ԥ��"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�"
      Begin VB.Menu mnuOpen 
         Caption         =   "���ļ�(O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistoryP 
         Caption         =   "��ʷ��¼"
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
      Caption         =   "��չ"
      Begin VB.Menu mnuEx 
         Caption         =   "��չ��"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����"
      Begin VB.Menu mnuAbout 
         Caption         =   "����..."
      End
   End
   Begin VB.Menu mnuListFile 
      Caption         =   "�ļ���"
      Visible         =   0   'False
      Begin VB.Menu mnuFileInfo 
         Caption         =   "�ļ���Ϣ"
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "ɸѡ��"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuLstShow 
      Caption         =   "POP_LSTSHOW"
      Visible         =   0   'False
      Begin VB.Menu mnuShowClear 
         Caption         =   "���"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text '�����ִ�Сд
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

Dim Extends() As EXTEND '��չ����
Dim cdlFile As New CommonDialog
Dim batch As Long '������
Dim history(1 To 5) As String '��ʷ��¼
Dim drawFlag As Boolean '��ȡ��ť ��ȡ(T)/ֹͣ(F) ״̬
Dim KeyStyle As Byte '����/����

Dim powerWeight As Single, addWeight As Long '��Ȩ��Ȩ��
Public FlagEx As Boolean '��չ��־
Public RMBuff As String '�ű�������
'��ȡ������
Public Function draw() As Long
    Dim i As Long, r As Long, tmp As Long
    '����ȡ�б����������Ȩ��Ϊ0������˳�
    If drawCount = 0 Or (bWei And wSUM = 0) Then
        MsgBox$ "�б��Ѿ���ȡ���!", vbExclamation, "����"
        drawFlag = False
        draw = -1
        Exit Function
    End If
    If chkWeight.Value Then '��Ȩ
        If chkPowerWeight.Value Then '��Ȩ
            Do
                i = Int(Rnd() * drawCount)
            Loop While powerWeight ^ (items(drawList(i)).Weight - wMIN) <= Rnd()
        Else '��ͨȨ
            If 2 * drawCount * wMAX / wSUM <= (drawCount + 1) / 2 Then 'ʹ�ö�����Ȩ��ȡ
                Do
                    i = Int(Rnd() * drawCount) '���ֳ�
                Loop While items(drawList(i)).Weight / wMAX <= Rnd()
            Else '������ȡ
                r = Int(Rnd() * wSUM)
                tmp = 0
                For i = 0 To drawCount - 1
                    tmp = tmp + items(drawList(i)).Weight
                    If r < tmp Then Exit For
                Next
            End If
        End If
    Else '��ͨ��ȡ
        i = Int(Rnd() * drawCount)
    End If
    draw = i
End Function
'���³�ȡ��
Public Sub reset()
    Dim i As Long, j As Long
    wSUM = 0
    For i = 0 To UBound(items)
        '����
        If cklList.Checked(i) Then
            If bWei Then '��Ȩ
                wSUM = wSUM + items(i).Weight
            End If
            drawList(j) = i
            j = j + 1
        End If
    Next
    drawCount = j
End Sub

'��ȡ�ļ�
Private Sub OpenFile(FileName As String)
    Dim sLines() As String, sec() As String, i As Long, j As Long, b() As Byte
    Dim KeyTable As String
    '��ʾ�ļ���
    mnuListFile.Caption = Mid$(FileName, InStrRev(FileName, "\") + 1)
    mnuListFile.Visible = True
    '�ó�ֵ
    bKey = False: bWei = False
    wSUM = 0: wMAX = 0: wMIN = &H7FFFFFFF
    '���
    lstShow.Clear
    cklList.Clear
    frmFilter.cklKey.Clear
    '��ȡ�ļ�
    Open FileName For Binary As #1
        b = Input(LOF(1), #1)
        '��ȡԤ��ȡͷ
        Do While b(i) <> vbKeyReturn
            Select Case b(i)
            Case vbKeyK: bKey = True
            Case vbKeyW: bWei = True
            End Select
            i = i + 1
        Loop
        '��ȡitems
        sLines = Split(b, vbCrLf)
        For i = 1 To UBound(sLines) '�ų���һ�е�Ԥ��ȡͷ
            If Len(sLines(i)) Then '��ǿ�
                sec = Split(sLines(i))
                i = i - 1 '��-1��+1����ȡַʱ�俪֧
                ReDim Preserve items(i)
                '�ж�Key
                If bKey Then
                    If Len(sec(1)) Then KeyTable = KeyTable + sec(1) + "," '�����ڹؼ��������
                    items(i).Key = sec(1)
                End If
                '����Weight
                If bWei Then
                    items(i).Weight = sec(1 - bKey) '����Key����Ϊsec(2)������Ϊsec(1)
                    wSUM = wSUM + items(i).Weight
                    If items(i).Weight < wMIN Then
                        wMIN = items(i).Weight
                    ElseIf items(i).Weight > wMAX Then
                        wMAX = items(i).Weight
                    End If
                End If
                'дitem
                items(i).Value = sec(0)
                cklList.AddItem sec(0): cklList.Checked(i) = True
                i = i + 1
            End If
        Next
        drawCount = UBound(items) + 1
        ReDim drawList(drawCount - 1)
        reset
        
        '��ӹؼ��ֵ�ɸѡ��
        Dim tmpCheckList As CheckList
        Set tmpCheckList = frmFilter.cklKey
        tmpCheckList.Clear
        If bKey Then
            Dim k As Long
            '��KeyTableתΪ ,K1,,K2,,K3,,...,,Kn, ����ʽ
            KeyTable = "," & Replace$(KeyTable, ",", ",,")
            KeyTable = Left$(KeyTable, Len(KeyTable) - 1) 'ȥ�����һ��,
            i = 1
            Do
                k = InStr(i + 2, KeyTable, ",,") '�ҵ���һ�� ",Ki,"
                If k = 0 Then Exit Do 'û���ҵ���˵���ѱ������
                'ȡ ",Ki," ���������������� ",Ki," �滻Ϊ""��ʵ��ȥ��
                KeyTable = Mid$(KeyTable, 1, k) + Replace$(KeyTable, Mid$(KeyTable, i, k - i + 1), "", k + 1)
                i = k + 1 '����
            Loop
            '�ָ�Key����ȥ�����Ե�","��Ȼ��",,"��Ϊ�ָ������зָ�
            sec = Split(Mid$(KeyTable, 2, Len(KeyTable) - 2), ",,")
            For i = 0 To UBound(sec)
                tmpCheckList.AddItem sec(i) '����ؼ��ֵ�CheckList
            Next
        End If
        Set tmpCheckList = Nothing '�����������
        'chkWeight��Ч�ж�
        If bWei Then
            chkWeight.Enabled = True
        Else
            chkWeight.Value = 0: chkWeight.Enabled = False
        End If
    Close #1
End Sub
'������ʷ��¼
Private Sub UpdateHistory(ByVal FileName As String)
    Dim i As Long, j As Long, s As String
    For i = 1 To 5
        If FileName = history(i) Then
            '����Ѵ��ڸü�¼������ŵ�history(1)
            s = history(i)
            For j = i To 2 Step -1
                history(j) = history(j - 1)
            Next
            history(1) = s
            GoTo UpdateMenu '���²˵�
        End If
    Next
    '��������ڣ����½���ŵ�history(1)
    If Len(Dir(FileName)) <> 0 And Len(FileName) = 0 Then Exit Sub '�������ļ����˳�
    For j = 4 To 1 Step -1 '��ʷ��¼����
        history(j + 1) = history(j)
    Next
    history(1) = FileName
UpdateMenu:
    For i = 1 To 5 '���²˵�
        If Len(history(i)) Then
            mnuHistory(i).Caption = history(i)
            mnuHistory(i).Enabled = True
        Else
            mnuHistory(i).Caption = "���ޣ�"
            mnuHistory(i).Enabled = False
        End If
    Next
End Sub
'��չ����
Private Function ShellEx(ByVal PID As Long, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
    Dim hProcess As Long, ExitCode As Long
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, PID)
    Do
        ShellEx = GetExitCodeProcess(hProcess, ExitCode)
        DoEvents
    Loop While ExitCode = STILL_ALIVE
    CloseHandle hProcess
End Function


'��Ȩ
Private Sub chkPowerWeight_Click()
    chkAddWeight.Value = chkPowerWeight.Value
End Sub

Private Sub chkPowerWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then powerWeight = Val(InputBox$("���������", "������Ȩ", powerWeight))
    If powerWeight = 0 Then MsgBox$ "��������Ϊ0��", vbCritical, "Error": chkPowerWeight.Value = 0
End Sub
'Ȩ��
Private Sub chkAddWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then addWeight = Val(InputBox$("������Ȩ��ֵ", "����Ȩ��", addWeight))
End Sub

'�������
Private Sub Form_Load()
    Dim i As Long, tmp As Long, s As String
    Dim lpTemp As Long, ApplicationName As String
    Dim buff As String
    
    Randomize '������ӷ�����
    'vbs��ʼ��
    vbs.Language = "VBScript"
    vbs.AddObject "RM", Me, True
    '��ȡ�����ļ�
    'ע�⣺���µ�ָ�����ΪΣ�ղ�������Ϊ���ڶ���ʱ�ַ�������ȡַ������ܵ��²���Ԥ��Ľ��
    '��ȡConfig·��ָ��
    ConfigName = App.Path & "\Config.ini"
    lpConfig = StrPtr(ConfigName)
    '��������С
    ApplicationName = "Config": lpTemp = StrPtr(ApplicationName)
    buffSize = getIniInt(lpTemp, StrPtr("BufferSize"), 256&, lpConfig)
    buff = Space$(buffSize)
    '��ȡ��ʽ
    optDraw(getIniInt(lpTemp, StrPtr("Draw"), 0&, lpConfig)).Value = True '����OptionButton
    'ȫ��CheckBox
    chkWeight.Value = getIniInt(lpTemp, StrPtr("WeightAllow"), 0&, lpConfig)
    chkAutoRemove.Value = getIniInt(lpTemp, StrPtr("RepeatAllow"), 0&, lpConfig)
    '��������
    batch = getIniInt(lpTemp, StrPtr("BatchNum"), 1&, lpConfig)
    '��ȡ�����
    L = getIniInt(lpTemp, StrPtr("NumL"), 0&, lpConfig)
    H = getIniInt(lpTemp, StrPtr("numH"), 0&, lpConfig)
    chkOnly.Value = getIniInt(lpTemp, StrPtr("NumOnly"), 0&, lpConfig)
    chkSort.Value = getIniInt(lpTemp, StrPtr("NumSort"), 0&, lpConfig)
    '��Ȩ
    '��Long����Single
    tmp = getIniInt(lpTemp, StrPtr("PowerWeightNum"), 0&, lpConfig)
    CopyMemory powerWeight, tmp, 4
    chkPowerWeight.Value = getIniInt(lpTemp, StrPtr("PowerWeight"), 0&, lpConfig)
    'Ȩ��
    addWeight = getIniInt(lpTemp, StrPtr("AddWeightNum"), 0&, lpConfig)
    chkAddWeight.Value = getIniInt(lpTemp, StrPtr("AddWeight"), 0&, lpConfig)
    '��չ����
    i = 0
    ApplicationName = "Extend": lpTemp = StrPtr(ApplicationName)
    Do
        '��ȡ����
        s = Left$(buff, getIniStr(lpTemp, StrPtr("E" & i + 1), 0&, StrPtr(buff), buffSize&, lpConfig))
        If Len(s) = 0 Then Exit Do 'δ�ҵ���չ
        ReDim Preserve Extends(i)
        tmp = InStr(s, ")") + 1 '�ҵ���չ���⿪ʼλ��
        If i > 0 Then Load mnuEx(i) '���²˵�
        mnuEx(i).Caption = Mid$(s, tmp, InStr(s, ":") - tmp) '��������ӵ��˵�
        With Extends(i) '������������
            .Command = Mid$(s, InStr(s, ":") + 1)
            .ExitCode = (InStr(s, "E") <> 0)
            .FlagEx = (InStr(s, "F") <> 0)
            .Startup = (InStr(s, "S") <> 0)
        End With
        If Extends(i).Startup Then mnuEx_Click CInt(i)
        i = i + 1
    Loop
    '��ʷ��¼
    ApplicationName = "History": lpTemp = StrPtr(ApplicationName)
    For i = 1 To 5
        history(i) = Left$(buff, getIniStr(lpTemp, StrPtr("H" & i), 0&, StrPtr(buff), 256&, lpConfig))
    Next

    txtBatch.Text = batch
    txtL.Text = L: txtH.Text = H
    '����ɸѡ��
    Load frmFilter
    '������ļ�
    If Not FlagEx Then mnuHistory_Click 1
    '����򿪽ű�
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
'��ȡ
Private Sub cmdDraw_Click()
    Dim i As Long
    reset
    If optDraw(0).Value Then
        drawFlag = Not drawFlag
        If drawFlag Then
            cmdDrawBatch.Enabled = False
            cmdDraw.Caption = "ֹͣ"
            Do
                i = draw()
                If i = -1 Then Exit Do '�������쳣
                lblShow.Caption = items(drawList(i)).Value
                DoEvents
            Loop While drawFlag
        
            If i <> -1 Then 'û���쳣����
                lstShow.AddItem "��ȡ " & items(drawList(i)).Value
                'Ȩ��
                If chkAddWeight.Value Then
                    items(drawList(i)).Weight = items(drawList(i)).Weight + addWeight
                    If items(drawList(i)).Weight > wMAX Then wMAX = items(drawList(i)).Weight
                End If
                '�Զ������ѳ�
                If chkAutoRemove.Value = 1 Then
                    '��������
                    cklList.Checked(drawList(i)) = False
                    If bWei Then wSUM = wSUM - items(drawList(i)).Weight
                    drawCount = drawCount - 1
                    drawList(i) = drawList(drawCount)
                End If
            End If
        End If
        If Not drawFlag Then '��ȡʱ��drawFlag = true
            cmdDrawBatch.Enabled = True
            cmdDraw.Caption = "��ȡ"
        End If
    Else
        L = Val(txtL.Text): H = Val(txtH.Text)
        i = Int((H - L + 1) * Rnd) + L
        MsgBox$ i, vbInformation, "��ȡ���"
        lstShow.AddItem "����� " & vbCrLf & i
    End If
End Sub

Private Sub cmdDrawBatch_Click()
    Dim i As Long, j As Long, tmp As Long, initCount As Long, s As String
    reset
    If optDraw(0).Value Then
        cmdDrawBatch.Enabled = False
        cmdDraw.Enabled = False
    
        s = txtGroupName.Text & " " & vbCrLf
        If Len(s) = 3 Then s = "δ������" & s '����Ϊ��
        initCount = drawCount - 1
        For i = 1 To batch
            j = draw()
            If j = -1 Then Exit For '�������쳣
            'Ȩ��
            If chkAddWeight.Value Then
                items(drawList(j)).Weight = items(drawList(j)).Weight + addWeight
                If items(drawList(j)).Weight > wMAX Then wMAX = items(drawList(j)).Weight
            End If
            s = s & items(drawList(j)).Value & " "
            '�Զ������ѳ飨û���쳣����ʱ��
            If chkAutoRemove.Value Then
                cklList.Checked(drawList(j)) = False
                If bWei Then wSUM = wSUM - items(drawList(j)).Weight
                drawCount = drawCount - 1
                drawList(j) = drawList(drawCount)
            Else
                tmp = drawList(j) '����drawList�е�j�������1��
                drawList(j) = drawList(drawCount - 1)
                drawList(drawCount - 1) = tmp
                If bWei Then wSUM = wSUM - items(tmp).Weight
                drawCount = drawCount - 1
            End If
        Next
        If chkAutoRemove.Value = 0 Then '���Զ�����
            If bWei Then '���¼ӻ�SUM
                For i = drawCount To initCount
                    wSUM = wSUM + items(drawList(i)).Weight
                Next
            End If
            drawCount = initCount
        End If
        If i <> 1 Then lstShow.AddItem s: MsgBox$ s 'û�г��0��
        cmdDrawBatch.Enabled = True
        cmdDraw.Enabled = True
    Else
        L = Val(txtL.Text): H = Val(txtH.Text)
        If batch > H - L + 1 Then MsgBox "�����С!", vbCritical, "Error"
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
        lstShow.AddItem "��������� " & vbCrLf & s
        MsgBox$ s, vbInformation, "��ȡ���"
    End If
End Sub

'չʾ�б�
Private Sub lstShow_DblClick()
    MsgBox$ lstShow.List(lstShow.ListIndex)
End Sub

Private Sub lstShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then PopupMenu mnuLstShow
End Sub
'��չ
Private Sub mnuEx_Click(Index As Integer)
    mnuEx(Index).Checked = Not mnuEx(Index).Checked
    If mnuEx(Index).Checked Then
        If Extends(Index).FlagEx Then FlagEx = True
        Dim s As String
        '�滻ϵͳ����
        s = Extends(Index).Command
        s = Replace$(s, "%HWND%", Me.hwnd)
        mnuEx(Index).Tag = Shell(App.Path & "\Ex\" & s)
        If Extends(Index).ExitCode Then
            MsgBox$ "��չ����""" & mnuEx(Index).Caption & """����ֹ����" & vbCrLf _
            & "�˳�����: " & ShellEx(mnuEx(Index).Tag) '������չ����
        End If
    Else
        Shell "cmd /c taskkill /f /pid " & mnuEx(Index).Tag, vbHide
    End If
End Sub

'����
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
        MsgBox$ "�޷�����!", vbCritical, "Error"
    End If
End Sub
'ɸѡ��
Private Sub mnuFilter_Click()
    frmFilter.Show
End Sub
'�ļ���Ϣ
Private Sub mnuFileInfo_Click()
    Dim s As String
    s = "�ؼ��֣�" & vbTab & IIf(bKey, "��", "��") & vbCrLf & _
        "Ȩ�أ�" & vbTab & IIf(bKey, "��", "��") & vbCrLf & _
        "������" & vbTab & UBound(items) + 1
    MsgBox$ s, vbInformation, "�ļ���Ϣ"
End Sub


'���չʾ�б�
Private Sub mnuShowClear_Click()
    lstShow.Clear
End Sub
'���ļ�
Private Sub mnuOpen_Click()
    cdlFile.showOpen Me.hwnd
    '���������·����һ����Ч������Ҫ����·���ж�
    '����Dir("")Ϊ���ҵ�ǰĿ¼��һ���ļ�������Ҫ���Ͽ��ַ����ж�
    If Len(Dir(cdlFile.FilePath)) <> 0 And Len(cdlFile.FilePath) <> 0 Then
        OpenFile cdlFile.FilePath
        UpdateHistory cdlFile.FilePath
    Else
        MsgBox$ "�ļ�������!", vbCritical, "Error"
    End If
End Sub
'����ʷ��¼
Private Sub mnuHistory_Click(Index As Integer)
    If Len(history(Index)) <> 0 Then '�ļ�����
        If Len(Dir(history(Index))) <> 0 Then
            OpenFile history(Index)
        Else
            MsgBox "�ļ�������!", vbCritical, "Error"
            history(Index) = Empty
        End If
    End If
    UpdateHistory history(Index)
End Sub
'����
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub
'��ȡģʽ 0:���� 1:�����
Private Sub optDraw_Click(Index As Integer)
    frmDrawList.Visible = (Index = 0)
End Sub

'������ȡ����
Private Sub txtBatch_Change()
    batch = Val(txtBatch.Text) '��Val()��ֹΪ��ʱ����
End Sub

'RM�ű�
Public Function extract(ByVal times As Long, ByVal bWeight As Boolean, Optional groupName As String = "δ������") As String
    extract = IIf(Len(groupName), groupName, "δ������") & vbCrLf
    Dim i As Long, j As Long, tmp As Long
    
    For i = 1 To times
        j = draw()
        If j = -1 Then Exit For '�������쳣
        '����SUM
        If bWei Then wSUM = wSUM - items(drawList(j)).Weight
        extract = extract & items(drawList(j)).Value & " "
        '�Զ������ѳ飨û���쳣����ʱ��
        drawList(j) = drawList(drawCount - 1)
    Next
    RMBuff = RMBuff & extract & vbCrLf
End Function

Public Function export(exportPath As String)
    If Len(exportPath) = 0 Then MsgBox$ "δָ������·����", vbCritical, "�ű�����"
    Open exportPath For Output As #2
        Print #2, RMBuff
    Close #2
End Function

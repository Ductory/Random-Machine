VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ɸѡ��"
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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "ɸѡ"
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
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.OptionButton optKey 
      Caption         =   "����"
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
      Caption         =   "�ַ���ƥ��"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblNumberMatch 
      Caption         =   "��ֵƥ��"
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

'hWndInsertAfter ������ѡֵ:
Private Const HWND_TOPMOST = -1 ' {��ǰ��, λ���κζ������ڵ�ǰ��}
'wFlags ������ѡֵ:
Private Const SWP_NOSIZE = 1 ' {���� cx��cy, ���ִ�С}
Private Const SWP_NOMOVE = 2 ' {���� X��Y, ���ı�λ��}

Dim KeyStyle As Long

Private Sub cmbFilter_Click() '��д����Ϊ��ѹ��������
    Dim i As Long
    i = cmbFilter.ListIndex
    lblStringMatch.Visible = (i = 0) 'Value
    lblNumberMatch.Visible = (i = 2) 'Weight
    txtMatch.Visible = (i <> 1) 'Value �� Weight
    optKey(0).Visible = (i = 1): optKey(1).Visible = (i = 1): cklKey.Visible = (i = 1) 'Key
    cmdRun.Top = 1080 - 1080 * (i = 1) 'λ�õ���
End Sub

'ɸѡ��
Private Sub cmdRun_Click()
    On Error GoTo Err
    Dim matchKey() As String
    Dim i As Long, j As Long, k As Long
    Dim KeyTable As String, sKey() As String
    '�����
    Dim tmpCheckList As CheckList
    Set tmpCheckList = frmMain.cklList
    For i = 0 To UBound(drawList)
        If tmpCheckList.Checked(i) Then '�ѹ�ѡ
            drawList(j) = i: j = j + 1 '�������룬���ó�ȡ��
        End If
    Next
    
    Select Case cmbFilter.ListIndex
    Case 0 'Value
        Dim s As String
        s = txtMatch.Text
        For i = 0 To UBound(drawList)
            '����ƥ�䣬��ȡ��ѡ��
            If Not items(drawList(i)).Value Like s Then tmpCheckList.Checked(drawList(i)) = False
        Next
    Case 1 'Key
        '�ؼ��ּ���
        j = 0
        If cklKey.Count = 0 Then Exit Sub '��δѡ��ؼ������˳�
        ReDim matchKey(cklKey.Count - 1)
        For i = 0 To UBound(matchKey)
            If cklKey.Checked(i) Then '��ѡ�е�Key���
                matchKey(j) = cklKey.Item(i) '��������
                KeyTable = KeyTable & "," & matchKey(j) '�����ؼ��ֱ�
                j = j + 1
            End If
        Next
        ReDim Preserve matchKey(j - 1)
        KeyTable = KeyTable & ","

        For i = 0 To tmpCheckList.Count - 1
            If Not tmpCheckList.Checked(i) Then GoTo Continue '��ȡ��һ��ѡ����
            '�ؼ�����ʽ
            '0���� 1����
            If KeyStyle = 0 Then
                sKey = Split(items(i).Key, ",") 'ȡ���ؼ���
                '��matchKey��Item�Ĺؼ�����һƥ��
                For k = 0 To UBound(sKey)
                    If InStr(KeyTable, "," & sKey(k) & ",") Then GoTo Continue '����
                Next
                tmpCheckList.Checked(i) = False 'δƥ�䵽���
            Else
                For j = 0 To UBound(matchKey)
                    '��items().KeyתΪ ",K1,K2,...,Kn," ����ʽ
                    '�� Ki תΪ ",Ki," ����ʽ
                    If InStr("," & items(i).Key & ",", "," & matchKey(j) & ",") = 0 Then tmpCheckList.Checked(i) = False: Exit For 'δƥ�䵽���
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
    '�б��ʼ��
    cmbFilter.AddItem "ֵ"
    cmbFilter.AddItem "�ؼ���"
    cmbFilter.AddItem "Ȩ"
    '��ȡ����
    cmbFilter.ListIndex = getIniInt(StrPtr("Config"), StrPtr("Filter"), 0&, lpConfig)
    KeyStyle = getIniInt(StrPtr("Config"), StrPtr("KeyFilter"), 0&, lpConfig)
    optKey(KeyStyle).Value = True
    '���㴰��
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1 '��ֹ���屻ɾ��
    Me.Hide
End Sub

Private Sub optKey_Click(Index As Integer)
    KeyStyle = Index
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (ByVal pOpenfilename As Any) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (ByVal pOpenfilename As Any) As Long

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Dim strFilePath As String
Dim FileName As OPENFILENAME

Public Function showOpen(ByVal hwnd As Long, Optional FilterContain As String) As Boolean
    'Set   the   parent   window
    FileName.hwndOwner = hwnd
    'Select   a   filter
    FileName.lpstrFilter = Replace$(FilterContain, "|", vbNullChar) & vbNullChar & vbNullChar
    'Set   the   title
    FileName.lpstrTitle = "选择一个列表文件"
    If GetOpenFileName(VarPtr&(FileName)) Then
        showOpen = True
        strFilePath = Trim$(FileName.lpstrFile)
    End If
End Function

Public Function showSave(ByVal hwnd As Long, Optional FilterContain As String) As Boolean
    'Set   the   parent   window
    FileName.hwndOwner = hwnd
    'Select   a   filter
    FileName.lpstrFilter = Replace$(FilterContain, "|", vbNullChar) & vbNullChar & vbNullChar
    'Set   the   title
    FileName.lpstrTitle = "另存为..."
    If GetSaveFileName(VarPtr&(FileName)) Then
        showSave = True
        strFilePath = Trim$(FileName.lpstrFile)
    End If
End Function

Property Get FilePath() As String
    FilePath = strFilePath
End Property

Private Sub Class_Initialize()
    FileName.lStructSize = Len(FileName)
    'Set   the   application's   instance
    FileName.hInstance = App.hInstance
    'create   a   buffer   for   the   file
    FileName.lpstrFile = Space$(buffSize)
    'set   the   maximum   length   of   a   returned   file
    FileName.nMaxFile = buffSize
    'Create   a   buffer   for   the   file   title
    FileName.lpstrFileTitle = Space$(buffSize)
    'Set   the   maximum   length   of   a   returned   file   title
    FileName.nMaxFileTitle = buffSize
    'Set   the   initial   directory
    FileName.lpstrInitialDir = ""
    'No   flags
    FileName.flags = 0
End Sub

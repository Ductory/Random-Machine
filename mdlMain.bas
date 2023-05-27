Attribute VB_Name = "mdlMain"
Option Explicit

Declare Function getIniInt Lib "kernel32" Alias "GetPrivateProfileIntW" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal nDefault As Any, ByVal lpFileName As Any) As Long
Declare Function getIniStr Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As Any, ByVal nSize As Any, ByVal lpFileName As Any) As Long
Declare Function setIniStr Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As Any) As Long

Public Type Item
    Value As String
    Key As String
    Weight As Long
End Type

Public items() As Item '���б�
Public drawList() As Long '��ȡ������
Public drawCount As Long  '��

Public ConfigName As String
Public lpConfig As Long
Public buffSize As Long
Public L As Long, H As Long
Public wMAX As Long, wMIN As Long, wSUM As Long '��Ȩ��Ϣ
Public bKey As Boolean, bWei As Boolean 'Ԥ��ȡͷ
'Public bReset As Boolean

Public vbs As New MSScriptControl.ScriptControl


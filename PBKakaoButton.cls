VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBKakaoButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public n        As String
Public t    As String
Public u1      As String
Public u2      As String

Public Sub fromDictionary(dic As Variant)
    n = dic.Item("n")
    t = dic.Item("t")
    u1 = dic.Item("u1")
    u2 = dic.Item("u2")
End Sub

Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")

    If n <> "" Then toDictionary.Add "n", n
    If t <> "" Then toDictionary.Add "t", t
    If u1 <> "" Then toDictionary.Add "u1", u1
    If u2 <> "" Then toDictionary.Add "u2", u2
End Function





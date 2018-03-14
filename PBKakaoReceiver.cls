VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBKakaoReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rcv As String
Public rcvnm As String
Public msg As String
Public altmsg As String

Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")

    If rcv <> "" Then toDictionary.Add "rcv", rcv
    If rcvnm <> "" Then toDictionary.Add "rcvnm", rcvnm
    If msg <> "" Then toDictionary.Add "msg", msg
    If altmsg <> "" Then toDictionary.Add "altmsg", altmsg
    
End Function


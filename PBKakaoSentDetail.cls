VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBKakaoSentDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public state As Integer
Public result As Integer
Public altSubject As String
Public altContent As String
Public altContentType As Integer
Public altResult As Integer
Public sendDT As String
Public resultDT As String
Public contentType As String
Public receiveNum As String
Public receiveName As String
Public content As String
Public altSendDT As String
Public altResultDT As String
Public receiptNum As String
Public requestNum As String
Public interOPRefKey As String


Public Sub fromDictionary(dic As Variant)

    state = dic.Item("state")
    result = dic.Item("result")
    altSubject = dic.Item("altSubject")
    altContent = dic.Item("altContent")
    altContentType = dic.Item("altContentType")
    altResult = dic.Item("altResult")
    sendDT = dic.Item("sendDT")
    resultDT = dic.Item("resultDT")
    contentType = dic.Item("contentType")
    receiveNum = dic.Item("receiveNum")
    receiveName = dic.Item("receiveName")
    content = dic.Item("content")
    altSendDT = dic.Item("altSendDT")
    altResultDT = dic.Item("altResultDT")
    receiptNum = dic.Item("receiptNum")
    requestNum = dic.Item("requestNum")
    interOPRefKey = dic.Item("interOPRefKey")
End Sub



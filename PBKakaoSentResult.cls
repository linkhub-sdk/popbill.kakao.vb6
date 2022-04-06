VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBKakaoSentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public contentType As String
Public templateCode As String
Public plusFriendID As String
Public sendNum As String
Public altSubject As String
Public altContent As String
Public altSendType As String
Public reserveDT As String
Public adsYN As Boolean
Public imageURL As String
Public sendCnt As String
Public successCnt As String
Public failCnt As String
Public altCnt As String
Public cancelCnt As String
Public msgs As Collection
Public btns As Collection

Public Sub fromDictionary(dic As Variant)

    contentType = dic.Item("contentType")
    templateCode = dic.Item("templateCode")
    plusFriendID = dic.Item("plusFriendID")
    sendNum = dic.Item("sendNum")
    altSubject = dic.Item("altSubject")
    altContent = dic.Item("altContent")
    altSendType = dic.Item("altSendType")
    reserveDT = dic.Item("reserveDT")
    adsYN = dic.Item("adsYN")
    imageURL = dic.Item("imageURL")
    sendCnt = dic.Item("sendCnt")
    successCnt = dic.Item("successCnt")
    failCnt = dic.Item("failCnt")
    altCnt = dic.Item("altCnt")
    cancelCnt = dic.Item("cancelCnt")
    
    Dim tCollection As Collection
    Set tCollection = dic.Item("msgs")
    
    If (tCollection Is Nothing) = False Then
        Set msgs = New Collection
        
        Dim t As Variant
        Dim searchList As PBKakaoSentDetail

        For Each t In tCollection
            Set searchList = New PBKakaoSentDetail
            searchList.fromDictionary t
            msgs.Add searchList
        Next
    End If
    
    If dic.Exists("btns") Then
        Set tCollection = dic.Item("btns")
        
        If (tCollection Is Nothing) = False Then
            Set btns = New Collection
            
            Dim buttonInfo As PBKakaoButton
    
            For Each t In tCollection
                Set buttonInfo = New PBKakaoButton
                buttonInfo.fromDictionary t
                btns.Add buttonInfo
            Next
            
        End If
    End If
    
    
End Sub





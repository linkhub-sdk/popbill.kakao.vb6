VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBATSTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public templateCode            As String
Public templateName           As String
Public template               As String
Public plusFriendID            As String

Public btns As Collection


Public Sub fromDictionary(dic As Variant)
    templateCode = dic.Item("templateCode")
    templateName = dic.Item("templateName")
    template = dic.Item("template")
    plusFriendID = dic.Item("plusFriendID")
    
    Dim tCollection As Collection
    
    If dic.Exists("btns") Then
    
        Set tCollection = dic.Item("btns")
                
        If (tCollection Is Nothing) = False Then
            Set btns = New Collection
            
            Dim T As Variant
            Dim btnInfo As PBKakaoButton
    
            For Each T In tCollection
                Set btnInfo = New PBKakaoButton
                
                btnInfo.fromDictionary T
                
                btns.Add btnInfo
            Next
        
        End If
    End If



    
    
    
End Sub




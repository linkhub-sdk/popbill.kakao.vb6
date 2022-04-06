VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBPlusFriend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public plusFriendID     As String
Public plusFriendName   As String
Public regDT            As String
Public state            As String
Public stateDT          As String

Public Sub fromDictionary(dic As Variant)
    plusFriendID = dic.Item("plusFriendID")
    plusFriendName = dic.Item("plusFriendName")
    regDT = dic.Item("regDT")
    state = dic.Item("state")
    stateDT = dic.Item("stateDT")
End Sub



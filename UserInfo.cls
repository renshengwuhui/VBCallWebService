VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:人生无悔-张建国
'Date:2018-04-19
'Comment:用户

Option Explicit

Private reUserCode As String
Private reUserPassword As String


Public Property Get UserCode() As String
    UserCode = reUserCode
End Property
Public Property Let UserCode(ByVal newValue As String)
    reUserCode = newValue
End Property


Public Property Get UserPassword() As String
    UserPassword = reUserPassword
End Property
Public Property Let UserPassword(ByVal newValue As String)
    reUserPassword = newValue
End Property

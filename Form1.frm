VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WebService���ò���"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   4695
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
   End
   Begin VB.TextBox txtReturnCode 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   5535
   End
   Begin VB.CommandButton cmdTestWebService 
      Caption         =   "����WebService"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "���"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:�����޻�-�Ž���
'Date:2018-04-19
'Comment:����WebService�ӿڵ���

Option Explicit

Dim webServiceServer As String
Dim webServiceAddress As String
Dim webServiceMethod As String
Dim responseText As String
Dim session As String

Private Sub GetPOById()
    Dim postUrl As String
    Dim request As WinHttpRequest
    Dim sc As Object
    Dim json As Object
    '��ȡ����
    webServiceAddress = "PurcharseOrder.asmx/GetPOById" 'ҵ��ӿ�
    postUrl = webServiceServer + webServiceAddress
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open webServiceMethod, postUrl, True 'ͬ����������
    request.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    request.SetRequestHeader "Cookie", session
    request.Send "{" & Chr(34) & "Id" & Chr(34) & ":" & Chr(34) & "1" & Chr(34) & "}"
    request.WaitForResponse '�ȴ�����
    responseText = request.responseText '����WebService���ص���JSON�ַ���
    responseText = Replace(responseText, "<?xml version=" + Chr(34) + "1.0" + Chr(34) + " encoding=" + Chr(34) + "utf-8" + Chr(34) + "?>", "")
    responseText = Replace(responseText, "<string xmlns=" + Chr(34) + "http://tempuri.org/" + Chr(34) + ">", "")
    responseText = Replace(responseText, "</string>", "")
    '��������
    Set sc = CreateObject("msscriptcontrol.scriptcontrol")
    sc.language = "javascript"
    Set json = sc.eval("data=" & responseText & ";")
    txtReturnCode.Text = json.ReturnCode '����WebService���ص���JSON�ַ����к���ReturnCode���Key
    txtResult.Text = responseText
End Sub

Private Sub LoginWebService()
    Dim postUrl As String
    Dim request As WinHttpRequest
    Dim instance As New UserInfo
    webServiceAddress = "System.asmx/Login" '��¼�ӿ�
    postUrl = webServiceServer + webServiceAddress
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open webServiceMethod, postUrl, True 'ͬ����������
    request.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    instance.UserCode = "zhangjianguo"
    instance.UserPassword = "123456789"
    request.Send "{" & Chr(34) & "UserCode" & Chr(34) & ":" & Chr(34) & instance.UserCode & Chr(34) & "," & _
        Chr(34) & "UserPassword" & Chr(34) & ":" & Chr(34) & instance.UserPassword & Chr(34) & "}"
    request.WaitForResponse '�ȴ�����
    responseText = request.responseText
    session = request.GetResponseHeader("Set-Cookie")
End Sub


Private Sub cmdTestWebService_Click()
    Call LoginWebService
    Call GetPOById
End Sub

Private Sub Form_Load()
    webServiceServer = "http://zhangjg-pc:8999/"
    webServiceMethod = "POST"
End Sub

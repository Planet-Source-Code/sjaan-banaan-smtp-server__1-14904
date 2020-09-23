VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2160
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Mail.frx":0000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MAILFROM As String
Private MAILTO As String
Private MAILDATA As String

Private DOMAIN As String
Private IP As String

Private DataOk As Boolean
Private DataSend As String
Private DataError As Boolean

Private Function IsHome(MailAddr, winsock As winsock) As Boolean
Dim MAILT As String

MAILT = Right(MailAddr, Len(MailAddr) - InStr(MailAddr, "@"))
IsHome = ((MAILT = DOMAIN) Or (MAILT = IP))
End Function

Private Sub Storemail(Username, MyData)
Dim x, i As Long
x = Dir(App.Path & "\mails\" & Username & Right("00" & i, 2) & ".txt")
Do Until x = ""
i = i + 1
x = Dir(App.Path & "\mails\" & Username & Right("00" & i, 2) & ".txt")
Loop
x = Left(Username, InStr(Username, "@") - 1) & Right("00" & CStr(i), 2)
Open App.Path & "\mails\" & x & ".txt" For Output As #1
Print #1, MyData
Close #1
End Sub

Private Sub SendMail(MailtoAddr, Maildatato, Mailfromdata)
Dim Username As String
Dim useraddr As String
Dim x As String
x = MailtoAddr
If x = "" Then Exit Sub
Username = Left(x, InStr(x, "@") - 1)
useraddr = Right(x, Len(x) - InStr(x, "@"))
DataOk = False
DataError = False
useraddr = IIf(UCase(Left(useraddr, 5)) <> "MAIL.", "mail." & useraddr, useraddr)
Winsock2.Close
Do Until Winsock2.State = sckClosed
DoEvents
Loop

Winsock2.Connect useraddr, 25

Do Until DataOk
If DataError Then Exit Sub
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "220" Then Winsock2.SendData "HELO " & useraddr & vbNewLine
Do Until DataOk
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "250" Then Winsock2.SendData "MAIL FROM: " & Mailfromdata & vbNewLine
Do Until DataOk
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "250" Then Winsock2.SendData "RCPT TO: " & x & vbNewLine
Do Until DataOk
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "250" Then Winsock2.SendData "DATA" & vbNewLine
Do Until DataOk
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "354" Then Winsock2.SendData Maildatato & vbNewLine & "." & vbNewLine
Do Until DataOk
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "250" Then Winsock2.SendData "QUIT" & vbNewLine
Do Until DataOk
DoEvents
Loop
DataOk = False
If Left(DataSend, 3) = "221" Then Winsock2.Close

End Sub

Private Sub Add(text, send As Boolean)
If send Then
Text1.SelColor = vbRed
Text1.SelText = Text1.SelText & text & vbNewLine
If Winsock1.State = 7 Then Winsock1.SendData text & vbNewLine
Else
Text1.SelColor = vbBlue
Text1.SelText = Text1.SelText & text
End If
End Sub


Private Sub Form_Load()
Winsock1.LocalPort = 25
Winsock1.Listen
If Dir(App.Path & "\mails\", vbDirectory) = "" Then MkDir App.Path & "\mails"
IP = Winsock1.LocalIP
StartIt
DOMAIN = NameByAddr(Winsock1.LocalIP)
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopIt
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
Me.Caption = Winsock1.RemoteHostIP
Text1.text = ""
Add "220 mail server ready...", True
MAILFROM = ""
MAILTO = ""
MAILDATA = ""
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim VDATA As String
Dim Thing As String
Dim MAILTO1 As String
Dim MAILT


Winsock1.GetData VDATA
Add VDATA, False
If (Left(VDATA, 4) = "HELO") Or (Left(VDATA, 4) = "EHLO") Then

    Add "250 " & Winsock1.LocalHostName & "  Hello " & NameByAddr(Winsock1.RemoteHostIP) & " [""" & Winsock1.RemoteHostIP & """], pleased to meet you", True
    
ElseIf Left(VDATA, 10) = "MAIL FROM:" Then

    MAILFROM = Replace(Replace(Right(VDATA, Len(VDATA) - 11), Chr(10), ""), Chr(13), "")
    RemoveStuff MAILFROM
    If CheckAddress(MAILFROM) Then
    Add "250 " & MAILFROM & " Sender ok", True
    Else
    Add "500 Invalid email address '" & MAILFROM & "'", True
    End If
    
ElseIf Left(VDATA, 8) = "RCPT TO:" Then
    
    MAILTO1 = Replace(Replace(Right(VDATA, Len(VDATA) - 9), Chr(10), ""), Chr(13), "")
    RemoveStuff MAILTO1
    
    If CheckAddress(MAILTO1) Then
    MAILTO = MAILTO & MAILTO1 & ";"
    Add "250 " & MAILTO1 & " RCPT ok", True
    Else
    Add "500 invalid rcpt '" & MAILTO1 & "'", True
    End If
    
ElseIf Left(VDATA, 4) = "DATA" Then
    If MAILTO = "" Then
    Add "500 no rcpt specified", True
    Else
    Add "354 Enter mail, end with ""."" on a line by itself", True
    End If
    
ElseIf Right(VDATA, 5) = vbNewLine & "." & vbNewLine Then
    MAILDATA = Left(VDATA, Len(VDATA) - 5)
    
    If MAILTO <> "" Then
    Add "250 Mail ok, ready for delivery", True
    Else
    Add "500 no rcpt specified", True
    End If
    
ElseIf Left(VDATA, 4) = "QUIT" Then

    Add "221 mail server closing connection", True
    For Each MAILT In Split(MAILTO, ";")
    If Not IsHome(MAILT, Winsock1) Then
    SendMail MAILT, MAILDATA, MAILFROM
    Else
     Storemail MAILT, MAILDATA
    End If
    Next
ElseIf Left(VDATA, 4) = "RSET" Then
    Winsock1.Close
End If
End Sub

Private Function CheckAddress(strAddress)
Dim x As String, temp
Dim b As String, a As String
x = strAddress
If InStr(x, "@") = 0 Then GoTo wrong
temp = Split(x, "@")
If UBound(temp) <> 1 Then GoTo wrong
If (temp(0) = "") Or (temp(1) = "") Then GoTo wrong
temp = Split(temp(1), ".")
If UBound(temp) < 1 Then GoTo wrong
If (temp(0) = "") Or (temp(1) = "") Then GoTo wrong
CheckAddress = True
Exit Function

wrong:
CheckAddress = False
Exit Function
End Function
Private Sub RemoveStuff(String1)
String1 = Replace(String1, "<", "")
String1 = Replace(String1, ">", "")
String1 = Replace(String1, "[", "")
String1 = Replace(String1, "]", "")
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
End Sub


Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Winsock2.GetData DataSend
DataOk = True
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock2.Close
DataError = True
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIRC - Michael Leaney"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
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
Private WithEvents Irc As CLIRC
Attribute Irc.VB_VarHelpID = -1
Option Compare Text
Option Explicit



'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 Thankyou
 






Function RandomString(strLength As Integer)
    Dim newString As String
    Dim tmpNum As Integer
    
    For tmpNum = 1 To strLength
        newString = newString + Chr(Int(Rnd * 70) + 32)
    Next tmpNum
    
    RandomString = newString
End Function

Private Sub Form_Load()
'Code by Michael Leaney _
 _
Please ask for authors permission before _
using code for any purpose what so ever _
many long hard hours of work have been _
put into this code, so please pay respect _
 _
 leahcimic@hotmail.com _
 _
 Thankyou

Set Irc = New CLIRC
Randomize Timer
Irc.SocketHandle = Form1.Winsock1
Irc.UserFullName = "Mics BOT"
Irc.UserHostName = "HOEST"
Irc.UserName = "546"
Irc.UserNick = "Sentinel" + Trim(Str(Int(Rnd * 11000)))
Irc.UserServName = "BLAH"
Irc.Connect "comcen.au.austnet.org"




End Sub


Private Sub Irc_Action(nick As String, User As String, Host As String, Target As String, Message As String)
If Message = "slaps " + Irc.UserNick + " around a bit with a large trout" Then
    Irc.Action Target, "Shoves the trout up " + nick + "'s bum"
End If
End Sub

Private Sub Irc_BAN(srcNick As String, srcUser As String, srcHost As String, Chan As String, tarNick As String, TarUser As String, TarHost As String)

    If srcNick = Irc.UserNick Then Exit Sub
    If Irc.UserNick Like tarNick And Irc.UserName Like TarUser And Irc.UserHostName Like TarHost Then
        Irc.Kick Chan, srcNick, "Do not abuse thet power of @"
        Irc.UNBAN Chan, tarNick + "!" + TarUser + "@" + TarHost
        Irc.Notice srcNick, "Hello " + srcNick + ", your BAN has affected me on " + Chan + "! Please fix! BAN MASK to remove is " + tarNick + "!" + TarUser + "@" + TarHost
    End If
End Sub

Private Sub Irc_Connected()
Irc.Join "#CLIRCDevelopment"

End Sub



Private Sub Irc_DEOP(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String)

If srcNick = Irc.UserNick Then Exit Sub
Irc.BAN Chan, "*!*@" + srcHost
Irc.OP Target, Target
Irc.Kick Chan, srcNick, "Go away!"

End Sub

Private Sub Irc_Join(nick As String, User As String, Host As String, Channel As String)
    Irc.Notice nick, "Hey " + nick + ", welcome to " + Channel
End Sub

Private Sub Irc_Kick(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String, Message As String)
    Irc.Join Chan
    Irc.Notice srcNick, "I'm not a football!"
End Sub

Private Sub Irc_Notice(nick As String, User As String, Host As String, Target As String, Message As String)
   If Target = Irc.UserNick And nick <> Irc.UserNick Then
        Irc.Notice nick, "Hello " + nick + ", no humans here ;)"
    End If
End Sub

Private Sub Irc_OP(srcNick As String, srcUser As String, srcHost As String, Chan As String, TargetNick As String)
If TargetNick = Irc.UserNick Then
'    Irc.DEOP Chan, TargetNick
    Irc.Notice srcNick, "Thanks for the +o, " + srcNick + "!"
End If
End Sub

Private Sub Irc_PrivateMessage(nick As String, User As String, Host As String, Target As String, Message As String)
    If Mid(Message, 1, 5) = "VOICE" Then
        Irc.MODELIST "+v", Mid(Message, 7), Irc.CLChan.NickListing(Mid(Message, 7), "[!@,+]*")
    End If
    If Mid(Message, 1, 5) = "BANAL" Then
        Irc.MODELIST "+b", Mid(Message, 7), Irc.ConvNToMask(Irc.CLChan.NickList(Mid(Message, 7)), 1)
    End If
    If Mid(Message, 1, 5) = "KIKAL" Then
        Irc.KICKLIST Mid(Message, 7), Irc.CLChan.NickList(Mid(Message, 7)), "Kick all request"
    End If
    If Mid(Message, 1, 5) = "OPALL" Then
        Irc.MODELIST "+o", Mid(Message, 7), Irc.CLChan.NickNOOPList(Mid(Message, 7))
    End If
    If Mid(Message, 1, 5) = "DOPAL" Then
        Irc.MODELIST "-o", Mid(Message, 7), Irc.CLChan.NickOPList(Mid(Message, 7))
    End If
    If Mid(Message, 1, 4) = "JOIN" Then
        Irc.Join Mid(Message, 6)
        Exit Sub
    End If

    
End Sub


Private Sub Irc_UNBAN(srcNick As String, srcUser As String, srcHost As String, Chan As String, tarNick As String, TarUser As String, TarHost As String)
    If Irc.UserNick Like tarNick And Irc.UserName Like TarUser And Irc.UserHostName Like TarHost Then
        Irc.Notice srcNick, "Thankyou " + srcNick + "!"
    End If
End Sub

Private Sub Winsock1_Connect()
    Irc.Connected
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim tmpData As String
    Winsock1.GetData tmpData
    Irc.RecvData tmpData
End Sub



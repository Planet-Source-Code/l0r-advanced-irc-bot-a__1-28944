VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtQuit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "I r da bawt"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Settings"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox IRCServer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Finished"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox AV 
      Appearance      =   0  'Flat
      Caption         =   "AutoVoice"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtProxyPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "8080"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtProxy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "206.112.72.3"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtNick 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "l0r[]"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "6667"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "l0r"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "l0r"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "irc.dal.net"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Sock1 
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
'Option Explicit
Dim TheData As String: Dim Chr1 As String: Dim SpyChan As String
Dim WatchChan As String: Dim Spy As Boolean: Dim Password As String
Dim i As Integer: Dim UserFlag As Integer: Dim tmpData As String

Dim section1 As String
Dim section2 As String
Dim section3 As String
Dim section4 As String
Dim section5 As String

Dim LoadedList As Boolean

Dim nick As String
Dim chan As String
Dim comm As String
Dim host As String
Dim ip As String
Dim xComm As String

Dim OldNick As String: Dim NewNick As String

Private Sub Combo1_Click()

On Error Resume Next

    INISetup App.Path & "\settings.ini"

    If LoadedList = True Then

        xListItem = Combo1.List(Combo1.ListIndex)

        txtServer.Text = Split(xListItem, ", ")(0)
        txtPort.Text = Split(xListItem, ", ")(1)

        Write_Ini "Config", "Server", txtServer.Text
        Write_Ini "Config", "Port", txtPort.Text

    End If

End Sub

Private Sub Command1_Click()
    
    Log "xIRC Bawt v1.0 BETA by l0r" & vbCrLf & "Connecting..."
    iConnect
    
End Sub
Private Sub Command2_Click()
    QUIT txtQuit.Text
End Sub

Private Sub Command4_Click()

    INISetup App.Path & "\settings.ini"

    Write_Ini "Config", "Server", txtServer.Text
    Write_Ini "Config", "Port", txtPort.Text
    Write_Ini "Config", "Nick", txtNick.Text
    Write_Ini "Config", "User", txtUser.Text
    Write_Ini "Config", "Email", txtEmail.Text
    Write_Ini "Config", "Pass", txtPassword.Text
    Write_Ini "Config", "Quit", txtQuit.Text

    Write_Ini "Proxy", "Server", txtProxy.Text
    Write_Ini "Proxy", "Port", txtProxyPort.Text

    Write_Ini "Automation", "Voice", AV.Value

End Sub

Private Sub Command5_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()

    Me.Visible = False
    Password = txtPassword.Text
    Form3.Show
    Load Form2
    LoadLists
    xMenu Form3
    Command1_Click
    INISetup App.Path & "\settings.ini"
    RefreshServers
    Combo1.Text = "Select a server..."
    ReadINI

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Shell_NotifyIcon NIM_DELETE, nid
    End

End Sub

Private Sub Sock1_Close()

Log "+CONNECTION LOST: " & Date & " - " & Time & " / " & IRCServer.Text
IRCServer.Text = ""
Form3.Timer2.Enabled = True

End Sub

Private Sub Sock1_Connect()

    'This line is for porting through a proxy:
    'Sock1.SendData "CONNECT " & txtServer & ":" & txtPort & " HTTP/1.0" & vbCrLf & vbCrLf

    Form3.Timer2.Enabled = False

    Log "+CONNECTED: " & Date & " - " & Time

    Register

    Form1.IRCServer.Text = Sock1.RemoteHost & "@" & Sock1.RemoteHostIP

End Sub

Private Sub Sock1_DataArrival(ByVal bytesTotal As Long)

Sock1.GetData TheData

Log TheData

If Split(TheData, " ")(1) = "311" Then
    ParseWhois TheData
End If

If Mid(TheData, 1, 4) = "PING" Then Sock1.SendData "PONG " & Mid(TheData, 6, Len(TheData) - 6) & vbCrLf

        section1 = Split(TheData, " ")(0)
        If Len(section1) > 0 Then
            nick = Split(section1, "!")(0): nick = Mid(nick, 2, Len(nick))
        End If

    Select Case UCase(Split(TheData, " ")(1))

        Case "NICK"
            OldNick = nick
            NewNick = Mid(Split(TheData, " ")(2), 2, Len(Split(TheData, " ")(2)) - 1)

        Case "JOIN"

            chan = Split(TheData, " ")(2): chan = Mid(chan, 2, Len(chan) - 3)
            host = Split(section1, "!")(1)
            
            ip = Split(host, "@")(1)

            If IsInList(Form4.Shitlist, nick) = 1 Then
                Send Kick(chan, nick, "lamer.")
                Send Ban(chan, ip)
            End If

            If IsInList(Form4.Shitlist, ip) = 1 Then
                Send Kick(chan, nick, "lamer.")
                Send Ban(chan, ip)
            End If

            If IsInList(Form4.AOP, nick) = 1 Then
                Send OP(chan, nick)
            End If

            If IsInList(Form4.AOP, ip) = 1 Then
                Send OP(chan, nick)
            End If

            If AV.Value = 1 Then
               Send Voice(chan, nick)
            End If

            If Spy = True Then
                If chan = SpyChan Then
                    Send Msg(WatchChan, "(" & chan & ") " & "- [" & host & "] " & nick & " has joined the channel.")
                End If
            End If

        Case "NOTICE"

        Case "PRIVMSG"

            section2 = Split(TheData, " ")(1)
            section3 = Split(TheData, " ")(2)
            section4 = Split(TheData, ":")(2)

            If Not section4 = "" Then
                comm = Mid(section4, 1, Len(section4) - 2)
            End If

            If Mid(comm, 1, 5) = Chr(1) & "PING" Then
                Send Msg(nick, Chr(1) & "PONG" & Mid(comm, 6, 10) & Chr(1))
            End If

            host = Split(section1, "!")(1)
            chan = section3

            If Spy = True Then
                If chan = SpyChan Then
                    Send Msg(WatchChan, "(" & chan & ")" & " <" & nick & "> " & comm)
                End If
            End If

            UserFlag = 0
            
            'This is optimized to do the least amount of processing possible. I think.
            
            If IsInList(Form4.Admins, nick) = 1 Then
                UserFlag = 3
                GoTo Validate
            End If

            If IsInList(Form4.Opers, nick) = 1 Then
                UserFlag = 2
                GoTo Validate
            End If

            If IsInList(Form4.Users, nick) = 1 Then
                UserFlag = 1
                GoTo Validate
            End If

            Debug.Print UserFlag
Validate:
            If Len(comm) > 0 Then
                If UserFlag > 1 Then

                    Select Case LCase(Split(comm, " ")(0))

                        Case "%stat"
                            Send "NICK l0r[" & Split(comm, " ")(1) & "]"

                        Case "%join"
                            Send "JOIN " & Split(comm, " ")(1)

                        Case "%part"
                            Send "PART " & Split(comm, " ")(1)

                        Case "%say"
                            Dim tmpData1 As String
                            Dim tmpData2 As String
                            tmpData1 = Split(comm, " ")(1)
                            tmpData2 = Split(comm, " ;")(1)
                            Send Msg(tmpData1, tmpData2)

                        Case "%spy"
                            Send Msg(chan, "ok")
                            SpyChan = Split(comm, " ")(1)
                            Spy = True
                            WatchChan = chan
                            Send "JOIN " & SpyChan

                        Case "%spy_off"
                            Send "PART " & SpyChan & " :oh ya i forgot."
                            Send Msg(WatchChan, "done")
                            Spy = False

                        Case "%nick"
                            Send "NICK " & Split(comm, " ")(1)
                            txtNick.Text = Split(comm, " ")(1)

                        Case "%quit"
                            QUIT txtQuit.Text

                        Case "%op"
                            Send Msg("chanserv@services.dal.net", "OP " & chan & " " & txtNick)

                        Case "%identify"
                            Send Msg("nickserv@services.dal.net", "identify " & Password)
                        
                        Case "%register"
                            Send "NICK " & Split(comm, " ")(1)
                            Send Msg("nickserv@services.dal.net", "register " & Password & " a@t3n.org")
                            DoEvents
                            Send Msg(chan, "ok done, " & Split(comm, " ")(1) & " is yours.")
                            DoEvents
                            Send "NICK " & txtNick.Text

                        Case "%ban"
                            xComm = Split(comm, " ")(1)
                            Send Ban(chan, xComm)

                        Case "%unban"
                            xComm = Split(comm, " ")(1)
                            Send unBan(chan, xComm)
 
                        Case "%adduser"
                            xComm = Split(comm, " ")(1)
                            If IsInList(Form4.Users, xComm) = 0 Then
                                Form4.Users.AddItem xComm
                                SaveList Form4.Users, "users.ini"
                                Send Msg(chan, "added, " & xComm)
                            Else
                                Send Msg(chan, xComm & " already exists as a user.")
                            End If

                        Case "%shitlist"
                            xComm = Split(comm, " ")(1)
                            If IsInList(Form4.Shitlist, xComm) = 0 Then
                                Form4.Shitlist.AddItem xComm
                                SaveList Form4.Shitlist, "shitlist.ini"
                                Send Msg(chan, "shitlisted, " & Split(comm, " ")(1))
                            Else
                                Send Msg(chan, xComm & " is already shitlisted.")
                            End If

                        Case "%aop"
                            xComm = Split(comm, " ")(1)
                            If IsInList(Form4.AOP, xComm) = 0 Then
                                Form4.AOP.AddItem xComm
                                SaveList Form4.AOP, "AOP.ini"
                                Send Msg(chan, "auto oped, " & xComm)
                            Else
                                Send Msg(chan, xComm & " already exists as an aop.")
                            End If

                        Case "%del_aop"
                            xComm = Split(comm, " ")(1)
                            SaveList Form4.AOP, "AOP.ini"
                            DeleteItem Form4.AOP, xComm
                            Send Msg(chan, "deleted op, " & xComm)

                        Case "%del_shitlist"
                            xComm = Split(comm, " ")(1)
                            DeleteItem Form4.Shitlist, xComm
                            SaveList Form4.Shitlist, "shitlist.ini"
                            Send Msg(chan, "de-shitlisted, " & xComm)

                        Case "%del_user"
                            xComm = Split(comm, " ")(1)
                            DeleteItem Form4.Users, xComm
                            SaveList Form4.Users, "users.ini"
                            Send Msg(chan, "deleted user, " & xComm)

                        Case "%addoper"
                            If IsInList(Form4.Admins, nick) Then
                                xComm = Split(comm, " ")(1)
                                If IsInList(Form4.Opers, xComm) = 0 Then
                                    Form4.Opers.AddItem xComm
                                    SaveList Form4.Opers, "opers.ini"
                                    Send Msg(chan, "added opper, " & xComm)
                                Else
                                    Send Msg(chan, xComm & " already exists as an operator.")
                                End If
                            Else
                                Send Msg(chan, "you don't have access to add opers.")
                            End If

                        Case "%del_oper"
                            If IsInList(Form4.Admins, nick) Then
                                xComm = Split(comm, " ")(1)
                                DeleteItem Form4.Opers, xComm
                                SaveList Form4.Users, "opers.ini"
                                Send Msg(chan, "deleted oper, " & xComm)
                            Else
                                Send Msg(chan, "you don't have access to delete opers.")
                            End If

                    End Select
                End If
                
                If UserFlag > 0 Then
                    Select Case LCase(comm)

                        Case "%time"
                            Send Msg(chan, nick & ", it's " & Time & ".")

                        Case "%date"
                            Sock1.SendData "PRIVMSG " & chan & " :" & nick & ", it's " & Date & "." & vbCrLf

                        Case "%check"
                            Send Msg(chan, "Chan: " & chan)
                            Send Msg(chan, "Nick: " & nick)
                            Send Msg(chan, "Host: " & host)
                            Send Msg(chan, "Comm: " & comm)
                            Send Msg(chan, "Time: " & Time & " | Date: " & Date)

                        Case "%opme"
                            Send OP(chan, nick)
                            Send Msg(chan, "granted.")

                    End Select

                If UCase(Split(comm, " ")(0)) = "%CONVERT" Then

                    Select Case UCase(Split(comm, " ")(1))

                        Case "HEX2ASC;"
                            Chr1 = Split(comm, "; ")(1)
                            Send Msg(chan, nick & ": " & Val("&H" & Chr1))

                        Case "ASC2HEX;"
                            Chr1 = Split(comm, "; ")(1)
                            If IsNumeric(Chr1) = True Then tmpData = Hex(Chr1)
                            If IsNumeric(Chr1) = False Then tmpData = Hex(Asc(Chr1))
                            Send Msg(chan, nick & ": " & tmpData)

                        Case "ASC;"
                            Chr1 = Mid(Split(comm, "; ")(1), 1, 2)
                            Send Msg(chan, nick & ": " & Asc(Chr1))

                        Case "CHR;"
                            Chr1 = Split(comm, "; ")(1)
                            Send Msg(chan, nick & ": " & Chr(Val(Chr1)))

                        Case "ASC2BIN;"
                            Chr1 = Split(comm, "; ")(1)
                            Send Msg(chan, nick & ": " & Binary(Chr1))

                        Case "BIN2ASC;"
                            Chr1 = Split(comm, "; ")(1)
                            Send Msg(chan, nick & ": " & Ascii(Chr1))

                        Case "ANSI_BOX;"
                            Send Msg(chan, nick & ": " & Box(Val(Split(comm, "; ")(1)), Val(Split(comm, "; ")(2)), Val(Split(comm, "; ")(3)), Val(Split(comm, "; ")(4)), Val(Split(comm, "; ")(5))))

                        Case "ELITE;"
                            'Send Msg(chan, nick, Elite(Split(comm, "; ")(1)))

                    End Select
                End If
            End If
        End If
    End Select

End Sub

Private Sub Sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Sock1.Close
End Sub

Public Sub RefreshServers()

On Error Resume Next
    
    INISetup App.Path & "\servers.ini"

    For i = 0 To Read_Ini("count", "val")

        DoEvents
        LineInfo = Read_Ini("servers", "n" & i)
        iserver = Split(LineInfo, "SERVER:")(1)
        iserver = Split(iserver, ":")(0)
        iport = Split(LineInfo, "SERVER:")(1)
        iport = Split(iport, "GROUP")(0)
        iport = Split(iport, ":")(1)
        iport = Split(iport, "-")(1)

        Combo1.AddItem iserver & ", " & Mid(iport, 1, 4)

    Next i

LoadedList = True

End Sub

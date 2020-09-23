Attribute VB_Name = "Module1"
'Option Explicit
Dim BinChr As String: Dim xChr As String: Dim xResult As String: Dim Z As Integer: Dim NewChr As Integer: Dim xB As Integer: Dim xB2 As Integer
Dim TmpStr As String: Dim NewBin As String: Dim C As Integer: Dim DD As Integer: Dim EE As Integer
Dim BinArray(7) As String: Dim User As String: Dim Password As String: Dim i As Integer 'Declarations
Dim Letters As String: Dim comm As String: Dim DataBuild As String

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const HWND_TOPMOST = -1

Public nid As NOTIFYICONDATA

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Global isWhois As Boolean
Global LogFile As String

Public Function xMenu(FormName As Form)
    
Dim hSysMenu As Long, nCnt As Long

    hSysMenu = GetSystemMenu(FormName.hwnd, False)

    If hSysMenu Then
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
            DrawMenuBar FormName.hwnd
        End If
    End If

End Function

Public Function Ascii(BinText As String) 'Convert Binary to ASCII

    For xB = 1 To Len(BinText)
        BinChr = Mid(BinText, xB, 8)
        Z = 128: NewChr = 0

        For xB2 = 1 To 8
            xChr = Mid(BinChr, xB2, 1)
            If xChr = "1" Then
                NewChr = NewChr + Z
                Z = Z / 2
            Else
                Z = Z / 2
            End If
        Next xB2

        xResult = xResult & Chr(NewChr)
        xB = xB + 7

    Next xB

    Ascii = xResult: xResult = ""

End Function

Public Function Binary(AscText As String) 'Convert ASCII to Binary
    
    For C = 1 To Len(AscText)
    
        DD = Asc(Mid(AscText, C, 1)): BinArray(7) = DD Mod 2
        
        DD = DD \ 2: BinArray(6) = DD Mod 2
        DD = DD \ 2: BinArray(5) = DD Mod 2
        DD = DD \ 2: BinArray(4) = DD Mod 2
        DD = DD \ 2: BinArray(3) = DD Mod 2
        DD = DD \ 2: BinArray(2) = DD Mod 2
        DD = DD \ 2: BinArray(1) = DD Mod 2
        DD = DD \ 2: BinArray(0) = DD Mod 2

        For EE = 0 To UBound(BinArray)
            TmpStr = TmpStr + BinArray(EE)
        Next EE

        NewBin = NewBin + TmpStr: TmpStr = ""

    Next C

    Binary = NewBin: NewBin = ""

End Function

Public Function Box(xPos As Integer, yPos As Integer, Width As Integer, Height As Integer, Border As Integer)

'Nifty function i made for creating ansi instructions for GUI, good for telnet servers.

Dim TopLeft As String: Dim TopRight As String
Dim BottomLeft As String: Dim BottomRight As String
Dim AcrossTop As String: Dim AcrossBottom As String
Dim DownLeft As String: Dim DownRight As String
Dim TheBoxData

If Border = 1 Then
    TopLeft = Chr(218): TopRight = Chr(191)
    BottomLeft = Chr(192): BottomRight = Chr(217)
    AcrossTop = Chr(196): AcrossBottom = Chr(196)
    DownLeft = Chr(179): DownRight = Chr(179)
End If

If Border = 2 Then
    TopLeft = Chr(201): TopRight = Chr(187)
    BottomLeft = Chr(200): BottomRight = Chr(188)
    AcrossTop = Chr(205): AcrossBottom = Chr(205)
    DownLeft = Chr(186): DownRight = Chr(186)
End If

If Border = 3 Then
    TopLeft = Chr(220): TopRight = Chr(220)
    BottomLeft = Chr(223): BottomRight = Chr(223)
    AcrossTop = Chr(220): AcrossBottom = Chr(223)
    DownLeft = Chr(221): DownRight = Chr(222)
End If

If Border = 4 Then
    TopLeft = Chr(219): TopRight = Chr(219)
    BottomLeft = Chr(219): BottomRight = Chr(219)
    AcrossTop = Chr(219): AcrossBottom = Chr(219)
    DownLeft = Chr(219): DownRight = Chr(219)
End If

    TheBoxData = TheBoxData & Mv(xPos, yPos) & TopLeft

        For i = 1 To Width - 2
            TheBoxData = TheBoxData & AcrossTop
            DoEvents
        Next

    TheBoxData = TheBoxData & Mv(xPos, Width + yPos - 1) & TopRight

    TheBoxData = TheBoxData & Mv(Height + xPos - 1, yPos) & BottomLeft

        For i = 1 To Width - 2
            TheBoxData = TheBoxData & AcrossBottom
            DoEvents
        Next

    TheBoxData = TheBoxData & Mv(Height + xPos - 1, yPos + Width - 1) & BottomRight

        For i = 1 To Height - 2
            TheBoxData = TheBoxData & Mv(xPos + i, yPos) & DownLeft
            DoEvents
        Next

        For i = 1 To Height - 2
            TheBoxData = TheBoxData & Mv(xPos + i, Width + yPos - 1) & DownRight
            DoEvents
        Next

Box = TheBoxData

End Function

Function Mv(xPos As Integer, yPos As Integer)

Mv = "[" & xPos & ";" & yPos & "H" 'ANSI code to move the cursors position.

End Function

Public Function l33t(ASCII_Code As String) 'A really stupid non-randomized not-so-slite speaker.

    Select Case ASCII_Code

        Case "a"
            l33t = Chr(229)
        
        Case "b"
            l33t = Chr(223)

        Case "c"
            l33t = Chr(231)
            
        Case "d"
            l33t = "Ð"
            
        Case "e"
            l33t = Chr(235)
            
        Case "i"
            l33t = Chr(238)

        Case "l"
            l33t = Chr(163)

        Case "n"
            l33t = Chr(241)
            
        Case "o"
            l33t = Chr(248)
             
        Case "u"
            l33t = Chr(249)
             
        Case "x"
            l33t = Chr(215)
             
        Case "A"
            l33t = Chr(197)
        
        Case "C"
            l33t = Chr(199)
                     
        Case "E"
            l33t = Chr(203)

        Case "N"
            l33t = Chr(209)

        Case "Y"
            l33t = Chr(165)
            
        Case "U"
            l33t = Chr(220)

        Case "I"
            l33t = Chr(207)
            
        Case "O"
            l33t = Chr(216)

        Case "?"
            l33t = Chr(191)
        
        Case "!"
            l33t = Chr(161)

        Case "0"
            l33t = Chr(186)

        Case "1"
            l33t = Chr(185)

        Case "2"
            l33t = Chr(178)

        Case "3"
            l33t = Chr(179)
        
        Case Else
            l33t = LCase(ASCII_Code)
        
    End Select

End Function

Public Function Elite(Text As String) 'Second routine for eliterizing

For i = 1 To Len(Text)
    Letters = Letters & l33t(Mid(Split(comm, "; ")(1), i, 1))
Next

Letters = Replace(Letters, "ØØ", "ÒÖ")
Letters = Replace(Letters, "øø", "òö")
Letters = Replace(Letters, "ÏÏ", "ÌÏ")
Letters = Replace(Letters, "îî", "ìï")
Letters = Replace(Letters, "ËË", "ÈË")
Letters = Replace(Letters, "ëë", "èë")
Letters = Replace(Letters, "ÅË", Chr(198))
Letters = Replace(Letters, "åë", Chr(230))
Letters = Replace(Letters, "îrç", "ì®©")
Letters = Replace(Letters, "ÏrÇ", "ì®©")

End Function

Public Function LoadLists() 'Load the user lists

LoadList Form4.AOP, "aop.ini"
LoadList Form4.Admins, "admins.ini"
LoadList Form4.Opers, "opers.ini"
LoadList Form4.Users, "users.ini"

End Function

Public Function Msg(xChannel As String, Message As String)

Msg = "PRIVMSG " & xChannel & " :" & Message

End Function

Public Function OP(xChannel As String, Nickname As String)

OP = "MODE " & xChannel & " +o " & Nickname

End Function

Public Function DEOP(xChannel As String, Nickname As String)

DEOP = "MODE " & xChannel & " -o" & Nickname

End Function

Public Function Voice(xChannel As String, xNickname As String)

Send "MODE " & xChannel & " +v " & xNickname

End Function

Public Function Ban(xChannel As String, NicknameOrHost As String)

Ban = "MODE " & xChannel & " +b "

    If IsNumeric(Mid(NicknameOrHost, 1, 2)) = True Then
        Ban = Ban & "*!*@" & NicknameOrHost
    Else
        Ban = Ban & NicknameOrHost
    End If

End Function

Public Function unBan(Channel As String, Nickname As String)

unBan = "MODE " & Channel & " -b "

    If IsNumeric(Mid(Nickname, 1, 2)) = True Then
        unBan = unBan & "*!*@" & Nickname
    Else
        unBan = unBan & Nickname
    End If

End Function

Public Function Kick(xChannel As String, Nickname As String, Reason As String)

Kick = "KICK " & xChannel & " " & Nickname & " :" & Reason

End Function

Public Function Send(Data As String)

If Len(Data) < 500 Then
    Form1.Sock1.SendData Data & vbCrLf
End If

End Function

Public Function Register()

    Send "USER " & Form1.txtUser & " ? * :" & Form1.txtEmail
    Send "NICK " & Form1.txtNick
    Send Msg("nickserv@services.dal.net", "identify " & Password)
    Send "JOIN #decode"

End Function

Public Function QUIT(QuitMessage As String)

        DoEvents
    Send "QUIT :" & QuitMessage
        DoEvents
    Form1.Sock1.Close

Log "+CONNECTION LOST: " & Date & " - " & Time & " / " & Form1.IRCServer.Text
Form1.IRCServer.Text = ""

End Function

Public Function IsInList(xControl As Control, SearchString As String)

    For i = 0 To xControl.ListCount

        If UCase(SearchString) = UCase(xControl.List(i)) Then
            IsInList = 1
            Exit For
        Else
            IsInList = 0
        End If

    Next i

End Function

Public Function SaveList(xList As Control, xFile As String)

Open App.Path & "\lists\" & xFile For Output As #1
    For i = 0 To xList.ListCount
        Print #1, xList.List(i)
    Next i

Close #1

End Function

Public Function LoadList(xControl As Control, ListFile As String)

xControl.Clear

Open App.Path & "\lists\" & ListFile For Input As #1

    Do While Not EOF(1)
        Line Input #1, User
        xControl.AddItem User
    Loop
Close #1

End Function

Public Function DeleteItem(xList As Control, xItem As String)

    For i = 0 To xList.ListCount
        If xItem = xList.List(i) Then xList.RemoveItem i
    Next i

End Function

Public Function Log(xData As String)

    If Form3.LogToScreen = True Then

        DataBuild = Form3.txtInData.Text

        If Form3.TimeStampOn = True Then
            DataBuild = DataBuild & "[" & Time & "] "
        End If

        Form3.txtInData.Text = DataBuild & xData & vbCrLf

    End If

    If Form3.LogToFile = True Then

        DataBuild = ""

        If Form3.TimeStampOn = True Then
            DataBuild = "[" & Time & "] "
        End If

    If Len(LogFile) > 0 Then
        Open LogFile For Append As #1
            Print #1, DataBuild & xData
        Close #1
    End If

    End If

End Function

Public Function ParseWhois(DataToParse As String)

    wSection1 = Split(DataToParse, vbCrLf)(0)
    wSection2 = Split(DataToParse, vbCrLf)(1)
    wSection3 = Split(DataToParse, vbCrLf)(2)
    wSection4 = Split(DataToParse, vbCrLf)(3)

    wNick = Split(wSection1, " ")(3)
    wUser = Split(wSection1, " ")(4)
    wHost = Split(wSection1, " ")(5)
    wName = Mid(Split(wSection1, " ")(7), 2, Len(Split(wSection1, " ")(7)))
    
    If Not Split(wSection2, " ")(1) = "319" Then
        wChan = Split(wSection2, ":")(2)
    End If

    wServer = Split(wSection3, " ")(4)
    wQuote = Split(wSection3, " :")(1)

If Len(wSection4) > 0 Then
    If Not Split(wSection4, " :")(1) = "End of /WHOIS list." Then
        wIdentified = Split(wSection4, " :")(1)
    End If
End If

    MsgBox "Nick: " & wNick & vbCrLf & _
           "User: " & wUser & vbCrLf & _
           "Host: " & wHost & vbCrLf & _
           "Name: " & wName & vbCrLf & _
           "Channels: " & wChan & vbCrLf & _
           "Server: " & wServer & " - " & wQuote & vbCrLf & _
           wNick & " " & wIdentified

End Function

Public Function ReadINI()

With Form3.txtInData

    INISetup App.Path & "\settings.ini"

    .FontName = Read_Ini("Visual", "FontName")
    .FontSize = Read_Ini("Visual", "FontSize")
    .FontBold = Read_Ini("Visual", "FontBold")
    .FontItalic = Read_Ini("Visual", "FontItalic")
    .FontStrikethru = Read_Ini("Visual", "FontStrike")
    .FontUnderline = Read_Ini("Visual", "FontUnder")

End With

With Form3.txtConsole

    INISetup App.Path & "\settings.ini"

    .FontName = Read_Ini("Visual", "cFontName")
    .FontSize = Read_Ini("Visual", "cFontSize")
    .FontBold = Read_Ini("Visual", "cFontBold")
    .FontItalic = Read_Ini("Visual", "cFontItalic")
    .FontStrikethru = Read_Ini("Visual", "cFontStrike")
    .FontUnderline = Read_Ini("Visual", "cFontUnder")

End With

If Read_Ini("Visual", "TimeStamp") = True Then
    Form3.TimeStampOn.Checked = True
Else
    Form3.TimeStampOn.Checked = False
    Form3.TimeStampOff.Checked = True
End If

If Read_Ini("Logging", "LogToFile") = True Then
    Form3.LogToFile.Checked = True
Else
    Form3.LogToFile.Checked = False
End If

If Read_Ini("Logging", "LogToScreen") = True Then
    Form3.LogToScreen.Checked = True
Else
    Form3.LogToScreen.Checked = False
End If

Form1.txtServer.Text = Read_Ini("Config", "Server")
Form1.txtPort.Text = Read_Ini("Config", "Port")
Form1.txtNick.Text = Read_Ini("Config", "Nick")
Form1.txtUser.Text = Read_Ini("Config", "User")
Form1.txtEmail.Text = Read_Ini("Config", "Email")
Form1.txtPassword.Text = Read_Ini("Config", "Pass")
Form1.txtQuit.Text = Read_Ini("Config", "Quit")

Form1.txtProxy.Text = Read_Ini("Proxy", "Server")
Form1.txtPort.Text = Read_Ini("Proxy", "Port")

Form1.AV.Value = Read_Ini("Automation", "Voice")

End Function

Public Sub iConnect()
    
    Form1.Sock1.Close
    Form1.Sock1.Connect Form1.txtServer, Form1.txtPort 'or for proxy: Sock1.Connect txtProxy, txtProxyPort

    Form3.Timer2.Enabled = True

End Sub

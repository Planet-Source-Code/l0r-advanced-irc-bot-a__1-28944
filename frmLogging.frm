VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000040&
   Caption         =   "Console - []"
   ClientHeight    =   4245
   ClientLeft      =   165
   ClientTop       =   720
   ClientWidth     =   9615
   LinkTopic       =   "Form3"
   ScaleHeight     =   4245
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Buffer 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   600
      Top             =   240
   End
   Begin MSComDlg.CommonDialog CommonD 
      Left            =   8880
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin VB.TextBox txtInData 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
   Begin VB.TextBox txtConsole 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   9615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Maximize 
         Caption         =   "Maximize"
      End
      Begin VB.Menu Minimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu Hide 
         Caption         =   "Hide"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "Settings"
      Begin VB.Menu MainSettings 
         Caption         =   "Main..."
      End
      Begin VB.Menu Lists 
         Caption         =   "Lists..."
      End
   End
   Begin VB.Menu Logging 
      Caption         =   "Logging"
      Begin VB.Menu TimeStamp 
         Caption         =   "Time Stamp"
         Begin VB.Menu TimeStampOn 
            Caption         =   "On"
            Checked         =   -1  'True
         End
         Begin VB.Menu TimeStampOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu LogTo 
         Caption         =   "Log To..."
         Begin VB.Menu LogToFile 
            Caption         =   "File..."
            Checked         =   -1  'True
         End
         Begin VB.Menu LogToScreen 
            Caption         =   "Screen"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu Display 
      Caption         =   "Display"
      Begin VB.Menu Font 
         Caption         =   "Font"
         Begin VB.Menu FontConsole 
            Caption         =   "Console..."
         End
         Begin VB.Menu FontInput 
            Caption         =   "Input..."
         End
      End
      Begin VB.Menu Color 
         Caption         =   "Color"
         Begin VB.Menu ColorConsole 
            Caption         =   "Console"
            Begin VB.Menu ConsoleForeColor 
               Caption         =   "ForeColor..."
            End
            Begin VB.Menu ConsoleBackColor 
               Caption         =   "BackColor..."
            End
         End
         Begin VB.Menu ColorInput 
            Caption         =   "Input"
            Begin VB.Menu InputForeColor 
               Caption         =   "ForeColor..."
            End
            Begin VB.Menu InputBackColor 
               Caption         =   "BackColor..."
            End
         End
         Begin VB.Menu Stripe 
            Caption         =   "Stripe"
            Begin VB.Menu StripeForeColor 
               Caption         =   "ForeColor..."
            End
            Begin VB.Menu StripBackColor 
               Caption         =   "BackColor..."
            End
         End
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Begin VB.Menu Whois 
         Caption         =   "Whois..."
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetryNum As Integer
Dim i As String

Private Sub Exit_Click()
    End
End Sub

Private Sub FontConsole_Click()

With CommonD
    .Flags = cdlCFScreenFonts
    .ShowFont
    txtInData.Font = .FontName
    txtInData.FontSize = .FontSize
    txtInData.FontBold = .FontBold
    txtInData.FontItalic = .FontItalic
    txtInData.FontStrikethru = .FontStrikethru
    txtInData.FontUnderline = .FontUnderline

    INISetup App.Path & "\settings.ini"

    Write_Ini "Visual", "FontName", .FontName
    Write_Ini "Visual", "FontSize", .FontSize
    Write_Ini "Visual", "FontBold", .FontBold
    Write_Ini "Visual", "FontItalic", .FontItalic
    Write_Ini "Visual", "FontStrike", .FontStrikethru
    Write_Ini "Visual", "FontUnder", .FontUnderline
End With

End Sub

Private Sub FontInput_Click()

With CommonD
    .Flags = cdlCFScreenFonts
    .ShowFont
    txtConsole.Font = .FontName
    txtConsole.FontSize = .FontSize
    txtConsole.FontBold = .FontBold
    txtConsole.FontItalic = .FontItalic
    txtConsole.FontStrikethru = .FontStrikethru
    txtConsole.FontUnderline = .FontUnderline

    INISetup App.Path & "\settings.ini"

    Write_Ini "Visual", "cFontName", .FontName
    Write_Ini "Visual", "cFontSize", .FontSize
    Write_Ini "Visual", "cFontBold", .FontBold
    Write_Ini "Visual", "cFontItalic", .FontItalic
    Write_Ini "Visual", "cFontStrike", .FontStrikethru
    Write_Ini "Visual", "cFontUnder", .FontUnderline
End With

End Sub

Private Sub Form_Load()
    INISetup App.Path & "\settings.ini"
    LogFile = Read_Ini("Logging", "LogFile")
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    txtInData.Move txtInData.Left, txtInData.Top, Me.ScaleWidth - 10, Me.ScaleHeight - Shape1.Height - txtConsole.Height
    Shape1.Move 1, Shape1.Top, Me.ScaleWidth, Shape1.Height
    txtConsole.Move 1, txtInData.Height + Shape1.Height - 20, Me.ScaleWidth, txtConsole.Height

    If WindowState = vbMinimized Then
        Me.Hide
        Me.Refresh

    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = Me.Caption & vbNullChar
    End With

        Shell_NotifyIcon NIM_ADD, nid
    Else
        Shell_NotifyIcon NIM_DELETE, nid

    End If

End Sub

Private Sub Hide_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Lists_Click()

    If Form4.Visible = True Then
        Form4.Visible = False
    Else
        Form4.Visible = True
    End If

End Sub

Private Sub LogToFile_Click()

INISetup App.Path & "\settings.ini"

If LogToFile.Checked = True Then
    LogToFile.Checked = False
Else
    LogToFile.Checked = True
    CommonD.DialogTitle = "Log to file..."
    CommonD.Filter = "All files | *.*"
    CommonD.ShowSave
    LogFile = CommonD.FileName
    If CommonD.FileName = "" Then LogFile = App.Path & "\logs\main.log"
End If

Write_Ini "Logging", "LogToFile", LogToFile.Checked
Write_Ini "Logging", "LogFile", LogFile

End Sub

Private Sub LogToScreen_Click()

INISetup App.Path & "\settings.ini"

If LogToScreen.Checked = True Then
    LogToScreen.Checked = False
Else
    LogToScreen.Checked = True
End If

Write_Ini "Logging", "LogToScreen", LogToScreen.Checked

End Sub

Private Sub MainSettings_Click()

    If Form1.Visible = True Then
        Form1.Visible = False
    Else
        Form1.Visible = True
    End If

End Sub

Private Sub Maximize_Click()
    Me.WindowState = vbMaximized
End Sub

Private Sub Minimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Timer1_Timer()

Me.Caption = "Console - [" & Form1.IRCServer.Text & "] - " & Time & " - " & Date

'If i = "11" Then i = "0"
'nid.hIcon = App.Path & "\icons\" & "a" & i & ".ico"
'Val(i) = Val(i) + 1

End Sub

Private Sub Timer2_Timer()

RetryNum = RetryNum + 1

Log "Connection timeout."
Log "Connection retry #" & RetryNum

iConnect

End Sub

Private Sub TimeStampOff_Click()

INISetup App.Path & "\settings.ini"

If TimeStampOff.Checked = False Then
    TimeStampOff.Checked = True
    TimeStampOn.Checked = False
End If

Write_Ini "Visual", "TimeStamp", TimeStampOn.Checked

End Sub

Private Sub TimeStampOn_Click()

INISetup App.Path & "\settings.ini"

If TimeStampOn.Checked = False Then
    TimeStampOn.Checked = True
    TimeStampOff.Checked = False
End If

Write_Ini "Visual", "TimeStamp", TimeStampOn.Checked

End Sub

Private Sub txtConsole_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Form1.Sock1.State = sckConnected Then
        Send txtConsole.Text
        Log "-> " & txtConsole.Text
        Buffer.AddItem txtConsole.Text
        txtConsole.Text = ""
        If Buffer.ListCount > 60 Then Buffer.Clear
    End If
End If

End Sub

Private Sub txtInData_Change()
    txtInData.SelStart = Len(txtInData.Text)
    If Len(txtInData.Text) > 40000 Then txtInData.Text = ""
    DoEvents
End Sub

Private Sub txtconsole_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

Select Case KeyCode

    Case 38
        Buffer.ListIndex = Buffer.ListIndex + 1
        txtConsole.Text = Buffer.List(Buffer.ListIndex)
        txtConsole.SelStart = Len(txtConsole.Text)

    Case 40
        Buffer.ListIndex = Buffer.ListIndex - 1
        txtConsole.Text = Buffer.List(Buffer.ListIndex)
        txtConsole.SelStart = Len(txtConsole.Text)

End Select

End Sub

Private Sub txtInData_KeyPress(KeyAscii As Integer)

txtConsole.SetFocus
txtConsole.Text = txtConsole.Text & Chr(KeyAscii)
txtConsole.SelStart = Len(txtConsole.Text)

End Sub

Private Sub txtInData_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Clipboard.SetText txtInData.SelText
txtConsole.SetFocus

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
Dim Sys As Long

    Sys = x / Screen.TwipsPerPixelX

    Select Case Sys

    Case WM_LBUTTONDOWN:
        Me.WindowState = vbNormal
        Me.Visible = True

    End Select

End Sub

Private Sub Whois_Click()

RetVal = InputBox("Nickname:", "Enter nickname...", ".")

Send "WHOIS " & RetVal

isWhois = True

End Sub

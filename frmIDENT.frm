VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   630
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrIDENT 
      Left            =   720
      Top             =   120
   End
   Begin MSWinsockLib.Winsock SockIDENT 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Ez-IDENT v1.0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IDENTname As String
Dim IDENTdata As String
 
Private Sub Form_Load()

    IDENTname = "l0r"
    SockIDENT.LocalPort = 113
    SockIDENT.Listen
    tmrIDENT.Interval = 250

End Sub

Private Sub SockIDENT_DataArrival(ByVal bytesTotal As Long)

    SockIDENT.GetData IDENTdata

    Log IDENTdata

End Sub

Private Sub tmrIDENT_Timer()

    If SockIDENT.State <> 2 And SockIDENT.State <> 7 Then
        SockIDENT.Close
        SockIDENT.Listen
    End If

End Sub

Private Sub SockIDENT_ConnectionRequest(ByVal requestID As Long)

    SockIDENT.Close
    SockIDENT.Accept requestID
    SockIDENT.SendData "113, 133:USERID:WIN32:" & IDENTname

    Log "-> IDENT REQUEST: " & SockIDENT.RemoteHostIP & " / " & SockIDENT.RemoteHost

End Sub

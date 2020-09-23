VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Manager - [ Users ]"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   390
      Left            =   2400
      ScaleHeight     =   330
      ScaleWidth      =   1155
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
      Begin VB.CommandButton Command3 
         Caption         =   "Finished"
         Height          =   330
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.ListBox Opers 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox Users 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   390
      Left            =   2400
      ScaleHeight     =   330
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   330
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   390
      Left            =   2400
      ScaleHeight     =   330
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.ListBox Admins 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox Shitlist 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   390
      Left            =   2400
      ScaleHeight     =   330
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   600
      Width           =   1215
      Begin VB.CommandButton Command4 
         Caption         =   "Refresh"
         Height          =   330
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.ListBox AOP 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Opers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Admins"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AutoOps"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Shitlist"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As String

Private Sub Command1_Click()

RetVal = InputBox("Add item: ", "Enter item to add...")

If AOP.Visible = True Then
    AOP.AddItem RetVal
    SaveList AOP, "aop.ini"
End If

If Admins.Visible = True Then
    Admins.AddItem RetVal
    SaveList Admins, "admins.ini"
End If

If Shitlist.Visible = True Then
    Shitlist.AddItem RetVal
    SaveList Shitlist, "shitlist.ini"
End If

If Users.Visible = True Then
    Users.AddItem RetVal
    SaveList Users, "users.ini"
End If

If Opers.Visible = True Then
    Opers.AddItem RetVal
    SaveList Opers, "opers.ini"
End If

End Sub

Private Sub Command2_Click()

If AOP.Visible = True Then
    AOP.RemoveItem AOP.ListIndex
    SaveList AOP, "aop.ini"
End If

If Admins.Visible = True Then
    Admins.RemoveItem Admins.ListIndex
    SaveList Admins, "admins.ini"
End If

If Shitlist.Visible = True Then
    Shitlist.RemoveItem Shitlist.ListIndex
    SaveList Shitlist, "shitlist.ini"
End If

If Users.Visible = True Then
    Users.RemoveItem Users.ListIndex
    SaveList Users, "users.ini"
End If

If Opers.Visible = True Then
    Opers.RemoveItem Opers.ListIndex
    SaveList Opers, "opers.ini"
End If

End Sub

Private Sub Command3_Click()
    Me.Hide
End Sub

Private Sub Command4_Click()
    LoadLists
End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Index

        Case 1
            Admins.Visible = False
            AOP.Visible = False
            Users.Visible = True
            Shitlist.Visible = False
            Opers.Visible = False
            Me.Caption = "List Manager [ Users ]"

        Case 2
            Admins.Visible = False
            AOP.Visible = False
            Users.Visible = False
            Shitlist.Visible = False
            Opers.Visible = True
            Me.Caption = "List Manager [ Opers ]"

        Case 3
            Admins.Visible = True
            AOP.Visible = False
            Users.Visible = False
            Shitlist.Visible = False
            Opers.Visible = False
            Me.Caption = "List Manager [ Admins ]"

        Case 4
            Admins.Visible = False
            AOP.Visible = True
            Users.Visible = False
            Shitlist.Visible = False
            Opers.Visible = False
            Me.Caption = "List Manager [ AutoOps ]"

        Case 5
            Admins.Visible = False
            AOP.Visible = False
            Users.Visible = False
            Shitlist.Visible = True
            Opers.Visible = False
            Me.Caption = "List Manager [ Shitlist ]"

    End Select

End Sub

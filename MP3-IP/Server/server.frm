VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form server 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3-IP-Server"
   ClientHeight    =   3120
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton mm_choice 
      Caption         =   "CD Audio"
      Height          =   372
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1200
      Width           =   3195
   End
   Begin VB.OptionButton mm_choice 
      Caption         =   "Other Device"
      Height          =   312
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Top             =   2520
      Width           =   3195
   End
   Begin VB.OptionButton mm_choice 
      Caption         =   "Wave Audio"
      Height          =   312
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   2100
      Width           =   3195
   End
   Begin VB.OptionButton mm_choice 
      Caption         =   "Midi Sequence"
      Height          =   312
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1680
      Width           =   3195
   End
   Begin VB.Frame mmcontrols_f 
      Caption         =   "Multi-Media Controls"
      Height          =   2892
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   3432
      Begin VB.CommandButton cmd_open 
         Height          =   552
         Left            =   2820
         Picture         =   "server.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   552
      End
      Begin MCI.MMControl mmcontrol 
         Height          =   552
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2676
         _ExtentX        =   4710
         _ExtentY        =   979
         _Version        =   393216
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         RecordVisible   =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
   End
   Begin VB.Frame network_f 
      Caption         =   "Network Status"
      Height          =   2892
      Left            =   3540
      TabIndex        =   0
      Top             =   120
      Width           =   3552
      Begin MSComctlLib.TreeView network 
         Height          =   2472
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   3252
         _ExtentX        =   5741
         _ExtentY        =   4366
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList client_icons 
      Left            =   420
      Top             =   60
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "server.frx":074C
            Key             =   "client_guest"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "server.frx":0B9E
            Key             =   "client_admin"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "server.frx":0FF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock mp3ip_server 
      Index           =   0
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   60
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
End
Attribute VB_Name = "server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intMax As Long
'*********************
Const guest = 1
Const admin = 2
'*********************

Private Sub cmd_open_Click()
    MMControl.Command = "Close"
    commond_ctrl
End Sub

Private Sub Form_Load()
    intMax = 0
    mp3ip_server(0).LocalPort = 1001
    mp3ip_server(0).Listen
    network.ImageList = client_icons
    mm_choice(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl.Command = "Close"
End Sub

Private Sub mm_choice_Click(Index As Integer)
    MMControl.Command = "Close"
    Select Case Index
        Case Is = 0
            MMControl.DeviceType = "CDAudio"
            cmd_open.Enabled = False
        Case Is = 1
            MMControl.DeviceType = "Sequencer"
            cmd_open.Enabled = True
            cmd_open.ToolTipText = "Open a Midi Sequence"
            CommonDialog1.Filter = "Midi Sequence (*.mid)|*.mid|"
        Case Is = 2
            MMControl.DeviceType = "Waveaudio"
            cmd_open.Enabled = True
            cmd_open.ToolTipText = "Open a Wave Audio file"
            CommonDialog1.Filter = "Wave Audio (*.wav)|*.wav|"
        Case Is = 3
            MMControl.DeviceType = "Other"
            cmd_open.Enabled = True
            cmd_open.ToolTipText = "Open Other"
            CommonDialog1.Filter = "All Files (*.*)|*.*|"
    End Select
    MMControl.Command = "Open"
End Sub

Private Sub mp3ip_server_Close(Index As Integer)
    delfrom_network mp3ip_server(Index).RemoteHostIP, CLng(Index)
End Sub

Private Sub mp3ip_server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        intMax = intMax + 1
        Load mp3ip_server(intMax)
        mp3ip_server(intMax).LocalPort = 0
        mp3ip_server(intMax).Accept requestID
        addto_network mp3ip_server(intMax).RemoteHostIP, admin, intMax
    End If
End Sub

Private Sub mp3ip_server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    mp3ip_server(Index).GetData strData, vbString
    Select Case strData
        Case Is = "PL"
            MMControl.Command = "Play"
        Case Is = "ST"
            MMControl.Command = "Stop"
        Case Is = "PA"
            MMControl.Command = "Pause"
        Case Is = "NE"
            MMControl.Command = "Next"
        Case Is = "PR"
            MMControl.Command = "Prev"
    End Select
End Sub

Private Sub mp3ip_server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Error = MsgBox(Description, vbOKOnly, "Winsock Error")
End Sub

Private Sub addto_network(client_name As String, client_type As Integer, connection As Long)
    Dim ok2add As Boolean
    ok2add = True
    For existing = 1 To network.Nodes.Count
        If network.Nodes.Item(existing) = client_name Then ok2add = False
    Next existing
    If ok2add = True Then network.Nodes.Add , , client_name & ":" & connection, client_name & ":" & connection, client_type
End Sub

Private Sub delfrom_network(client_name As String, connection As Long)
    For existing = 1 To network.Nodes.Count
        If network.Nodes(existing) = client_name & ":" & connection Then network.Nodes.Remove (client_name & ":" & connection)
        Exit For
    Next existing
End Sub

Private Sub commond_ctrl()
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    MMControl.FileName = CommonDialog1.FileName
    MMControl.Command = "Open"
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3-IP-Client"
   ClientHeight    =   1104
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   3564
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1104
   ScaleWidth      =   3564
   Begin VB.CommandButton cmd_mmcontrol 
      Height          =   372
      Index           =   4
      Left            =   2760
      Picture         =   "Client.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   492
   End
   Begin VB.CommandButton cmd_mmcontrol 
      Height          =   372
      Index           =   3
      Left            =   2100
      Picture         =   "Client.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   492
   End
   Begin VB.CommandButton cmd_mmcontrol 
      Height          =   372
      Index           =   2
      Left            =   1500
      Picture         =   "Client.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   492
   End
   Begin VB.CommandButton cmd_mmcontrol 
      Height          =   372
      Index           =   1
      Left            =   900
      Picture         =   "Client.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   492
   End
   Begin MSWinsockLib.Winsock mp3ip_client 
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Multi-Media Controls"
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3312
      Begin VB.CommandButton cmd_mmcontrol 
         Height          =   372
         Index           =   0
         Left            =   180
         Picture         =   "Client.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   492
      End
   End
End
Attribute VB_Name = "client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_mmcontrol_Click(Index As Integer)
    Select Case Index
        Case Is = 0
            mp3ip_client.SendData "PR"
        Case Is = 1
            mp3ip_client.SendData "ST"
        Case Is = 2
            mp3ip_client.SendData "PL"
        Case Is = 3
            mp3ip_client.SendData "PA"
        Case Is = 4
            mp3ip_client.SendData "NE"
    End Select
End Sub

Private Sub Form_Load()
    Dim inistring As String
    Dim line As Integer
    line = 1
    Open "startup.ini" For Input As #1
    Do While Not EOF(1)
        Input #1, inistring
        Select Case line
            Case Is = 1
                mp3ip_client.RemoteHost = inistring
            Case Is = 2
                mp3ip_client.RemotePort = inistring
        End Select
        line = line + 1
    Loop
    Close #1
    mp3ip_client.Connect
End Sub

Private Sub mp3ip_client_Close()
    Unload client
End Sub

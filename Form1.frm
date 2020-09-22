VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Make Ping V1.0, by Patricio 'DiNeSat4' Tapia  (patricio.tapia@gmail.com)"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   330
      Left            =   8775
      TabIndex        =   7
      Top             =   3870
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7335
      Top             =   2070
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP PING!"
      Height          =   330
      Left            =   8640
      TabIndex        =   4
      Top             =   3510
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PING!"
      Height          =   330
      Left            =   7470
      TabIndex        =   2
      Top             =   3510
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   3510
      Width           =   7350
   End
   Begin VB.TextBox Text1 
      Height          =   3480
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   10005
   End
   Begin VB.Label Label3 
      Caption         =   "patricio.tapia@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8235
      TabIndex        =   6
      Top             =   4365
      Width           =   1770
   End
   Begin VB.Label Label2 
      Caption         =   "CODE MAKED BY PATRICIO 'DINESAT4' TAPIA "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   4815
      TabIndex        =   5
      Top             =   4365
      Width           =   3390
   End
   Begin VB.Label Label1 
      Caption         =   "ADD YOUR IP OR HOST, PRESS ""PING!"" OR ENTER AND ""STOP PING"" OR ESC FOR STOP THE PINGING"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   3870
      Width           =   8160
   End
   Begin VB.Menu mabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dir As String    'Declare a String Public, this variable will keep the address ip
Private Sub Command1_Click()
    sendping 'Call to procedure called "sendping"
End Sub

Private Sub Command2_Click()
    stopping 'Call to "stopping" procedure
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub mabout_Click()
    Form2.Show vbModal
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'If the key pressed is ENTER
        sendping    'Call to procedure called "sendping"
    End If
    If KeyAscii = 27 Then
        stopping 'Call to "stopping" procedure
    End If
End Sub

Private Sub Timer1_Timer()
     If Not makeping(dir) Then 'Call to method "makeping" (in the Module), If the ping has failed then not send more pings
        Timer1.Enabled = False
     End If
     Text1.SelLength = Len(Text1)
End Sub

Private Sub sendping()
    If Trim(Text2) = "" Then
        MsgBox "Please add a ip or host", vbCritical + vbOKOnly
        Exit Sub
    End If
    Text1 = ""
    Text1 = "Pinging " & Text2 & "..."
    dir = Text2 'Save the address into the "dir" variable
    Timer1.Enabled = True   'Activated the TIMER control
End Sub
Private Sub stopping()
    If Not Timer1.Enabled Then
        Exit Sub
    End If
    Timer1.Enabled = False  'Deactivated the TIMER control
    Text1 = Text1 & vbCrLf & "Ping to " & dir & " has stopped"
End Sub

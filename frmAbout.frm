VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   2676
   ClientLeft      =   2304
   ClientTop       =   1668
   ClientWidth     =   5736
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1842.359
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      ClipControls    =   0   'False
      Height          =   432
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   263.118
      ScaleMode       =   0  'User
      ScaleWidth      =   263.118
      TabIndex        =   0
      Top             =   240
      Width           =   432
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CREATED BY RALPHY BOY. "
      ForeColor       =   &H00FFFFFF&
      Height          =   336
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   4116
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "DOWNLOAD SOURCE FROM PLANET-SOURCE-CODE.COM"
      ForeColor       =   &H00FFFFFF&
      Height          =   576
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   4236
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Super Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HelloNum As Integer

Unload Me


End Sub

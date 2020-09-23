VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSource 
   Caption         =   "HTML Source"
   ClientHeight    =   5784
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   7680
   Icon            =   "frmSource.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5784
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   4680
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbSource 
      Height          =   5772
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7692
      _ExtentX        =   13568
      _ExtentY        =   10181
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSource.frx":0442
   End
   Begin VB.Menu SmnuCopy 
      Caption         =   "&Copy"
   End
   Begin VB.Menu SmnuQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
rtbSource.Move 0, 0, ScaleWidth, ScaleHeight
End Sub



Private Sub SmnuCopy_Click()

If Text1.SelLength = 0 Then
    Beep
    Exit Sub
End If
Clipboard.SetText Text1.SelText
End Sub
End Sub

Private Sub SmnuQuit_Click()
    Unload Me
End Sub

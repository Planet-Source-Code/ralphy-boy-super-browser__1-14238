VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Super Web Browser"
   ClientHeight    =   6348
   ClientLeft      =   132
   ClientTop       =   360
   ClientWidth     =   7200
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6348
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1560
      Top             =   5160
      _ExtentX        =   804
      _ExtentY        =   804
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   360
      Top             =   5040
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   6096
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "CONNECTED"
            TextSave        =   "CONNECTED"
            Key             =   "connect"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   372
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   852
   End
   Begin VB.TextBox txtWebsite 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   6252
   End
   Begin MSComctlLib.Toolbar tlbmain 
      Align           =   1  'Align Top
      Height          =   576
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1016
      ButtonWidth     =   1185
      ButtonHeight    =   889
      ImageList       =   "imltoolbarpictures"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Description     =   "Back Button"
            Object.ToolTipText     =   "Go Back a Webpage"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Description     =   "Forward Button"
            Object.ToolTipText     =   "Go Forward a Webpage"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Description     =   "Stop the Transfer"
            Object.ToolTipText     =   "Stop the download"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Description     =   "Refresh Button"
            Object.ToolTipText     =   "Refresh this page"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Description     =   "Go Home Button"
            Object.ToolTipText     =   "Go back home"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imltoolbarpictures 
      Left            =   720
      Top             =   5160
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0706
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4932
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   7212
      ExtentX         =   12721
      ExtentY         =   8700
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Page"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Page"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print Page"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy "
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditEdit 
         Caption         =   "&Edit Source"
      End
   End
   Begin VB.Menu mnuOptView 
      Caption         =   "&View"
      Begin VB.Menu mnuOptViewTool 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptViewAdd 
         Caption         =   "&Address"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SitesVisit As Integer
Dim urlurl As String
Dim urlstring As String

Private Sub cmdGo_Click()
'You go to the website of whatever is in the textbox
    WebBrowser1.Navigate txtWebsite.Text
    
End Sub

Private Sub Form_Resize()
'This just resizes the WebBrowser if the form is resized
WebBrowser1.Move 0, 1050, ScaleWidth, ScaleHeight - 1050
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuEditCopy_Click()
    '// tell the web browser to save the page
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub mnuEditEdit_Click()
    On Error GoTo 1
    
    frmSource!rtbSource.Text = Inet1.OpenURL(urlurl)
    
    frmSource.Show
    
1 End Sub

Private Sub mnuFileExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFileOpen_Click()

On Error Resume Next
    Com.Filter = "All Internet Files (*.hmt,*.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    Com.ShowOpen
If Com.FileName = "" Then
    Exit Sub
Else
    WebBrowser1.Navigate (Com.FileName)
End If

End Sub

Private Sub mnuFilePrint_Click()

    '// tell the web browser control to print the current page
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
    
End Sub

Private Sub mnuFileSave_Click()
    
    Com.Filter = "htm (*.htm) | *.htm"
    Com.ShowSave
If Com.FileName = "" Then
    Exit Sub
Else
    Open Com.FileName For Output As #1
     Print #1, WebBrowser1.Document
   Close #1
End If

End Sub


Private Sub mnuOptViewAdd_Click()

    If txtWebsite.Visible = True Then   'If its visible then get rid of it and uncheck the menu
        txtWebsite.Visible = False
        cmdGo.Visible = False
        mnuOptViewAdd.Checked = False
    Else                                'If it is not visible then bring it back and check the menu
        txtWebsite.Visible = True
        cmdGo.Visible = True
        mnuOptViewAdd.Checked = True
    End If
    
            
End Sub

Private Sub mnuOptViewTool_Click()

    If tlbmain.Visible = True Then         'If the toolbars is visible then
        tlbmain.Visible = False            'Get rid of it and
        mnuOptViewTool.Checked = False     'Uncheck the menu thing
        
    Else                                   'If it isn't visible then
        tlbmain.Visible = True             'Bring it back and
        mnuOptViewTool.Checked = True      'Check that thing
    End If
    
    
End Sub

Private Sub tlbmain_ButtonClick(ByVal Button As MSComctlLib.Button)
'Find out which button has been pressed and
'do stuff

Select Case Button.Key
    Case Is = "Back"
        WebBrowser1.GoBack
    Case Is = "Forward"
        WebBrowser1.GoForward
    Case Is = "Stop"
        WebBrowser1.Stop
    Case Is = "Refresh"
        WebBrowser1.Refresh
    Case Is = "Home"
        WebBrowser1.GoHome
    End Select
    
End Sub



Private Sub WebBrowser1_TitleChange(ByVal Text As String)
    

    On Error Resume Next
    Let urlname = WebBrowser1.LocationName
    
    frmMain.Caption = urlname
    
    Let urlurl = WebBrowser1.LocationURL
    
    txtWebsite.Text = urlurl
    
End Sub

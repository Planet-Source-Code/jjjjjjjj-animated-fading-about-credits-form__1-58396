VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":1CFA
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Interval        =   10
      Left            =   480
      Top             =   5040
   End
   Begin VB.PictureBox PicShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000A80FF&
      Height          =   3825
      Left            =   960
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   0
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label lbpause 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   360
      Left            =   1320
      TabIndex        =   6
      Top             =   5160
      Width           =   810
   End
   Begin VB.Label lbCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F3F86&
      Height          =   360
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   3600
   End
   Begin VB.Label lbDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F3F86&
      Height          =   240
      Left            =   3000
      TabIndex        =   4
      Top             =   4680
      Width           =   3600
   End
   Begin VB.Label lbCompany 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F3F86&
      Height          =   225
      Left            =   4830
      TabIndex        =   3
      Top             =   285
      Width           =   1200
   End
   Begin VB.Label lbVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F3F86&
      Height          =   240
      Left            =   3120
      TabIndex        =   2
      Top             =   4920
      Width           =   2085
   End
   Begin VB.Label lbClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   360
      Left            =   7440
      TabIndex        =   1
      Top             =   5160
      Width           =   735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************
'(Adout / Credits) Form - Animated
'*********************************
'By     Jim Jose
'email  jimjosev33@yahoo.com
'*********************************
'The RGB values used in this code are
'according to the background image.
'You can use more attractive combinations
'You just edit the 'credits.txt' file
'in the app folder
'*********************************
Option Explicit

Private Sub Form_Load()
    vTop = PicShow.Height
    lbCredits = App.ProductName & " - Credits"
    lbCompany = App.CompanyName
    lbDescription = App.FileDescription
    lbVersion = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
    
'Loading the data from txtfile
Dim TxtSource   As String
    TxtSource = LoadText(App.Path & "\Credits.txt")
    CrdLines = Split(TxtSource, vbCrLf) 'Getting into lines
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbClose.ForeColor = RGB(119, 60, 0)
End Sub

Private Sub lbClose_Click()
    Unload Me
End Sub

Private Sub lbClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbClose.ForeColor = RGB(255, 128, 10)
End Sub

'Pausing the scrolling
Private Sub lbpause_Click()
    tmrUpdate.Enabled = Not tmrUpdate.Enabled
    If tmrUpdate.Enabled = False Then
        lbpause.ForeColor = RGB(255, 128, 10)
    Else
        lbpause.ForeColor = RGB(119, 60, 0)
    End If
End Sub

Private Sub PicShow_Click()
    lbpause_Click
End Sub

'Updating the animation
Private Sub tmrUpdate_Timer()
Dim X As Integer
Dim nTop As Long
    PicShow.Cls
    nTop = vTop
    For X = 0 To UBound(CrdLines)
        'if the 'top' is inside the picturebox then draw
        If nTop > -50 And nTop < PicShow.Height Then SendCredits PicShow, CrdLines(X), 33, nTop, vbBlack, RGB(205, 128, 0), vbBlack, 1 / 6
        nTop = nTop + PicShow.TextHeight(CrdLines(X))
    Next X
    'Reloading at the end of the file
    If vTop + 20 < -PicShow.TextHeight("A") * UBound(CrdLines) Then vTop = PicShow.Height
    vTop = vTop - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("(  PLEASE 'RATE' THIS CODE  ).Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( PLEASE GIVE FEEDBACK ) to improve this code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If
    If X = vbOK Then Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=58396&lngWId=1")
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Password Generator"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H000080FF&
      Caption         =   "&Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   8
      Left            =   3000
      Picture         =   "frmMain.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   7
      Left            =   2640
      Picture         =   "frmMain.frx":1194
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   6
      Left            =   2280
      Picture         =   "frmMain.frx":1A5E
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   5
      Left            =   1920
      Picture         =   "frmMain.frx":2328
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   4
      Left            =   1560
      Picture         =   "frmMain.frx":2BF2
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   3
      Left            =   1200
      Picture         =   "frmMain.frx":34BC
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":3D86
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   2
      Left            =   840
      Picture         =   "frmMain.frx":4650
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPenguin 
      Height          =   360
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":4F1A
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generated Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopy_Click()
    CopyTextToClipboard lblPassword
End Sub

Private Sub cmdGenerate_Click()
    getRandomPassword
End Sub

Sub getRandomPassword()

    Dim randMax, randMin

    Dim days
    days = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")

    Dim colors
    colors = Array("Red", "Orange", "Yellow", "Green", "Blue", "Purple", "Pink", "Brown", "Black", "White")

    ' Get current weekday:    (This could be used in the final concatenation, but for the sake of readability, I'll
    ' do it here.
    Dim currentDay
    currentDay = days(Weekday(Date) - 1)

    ' Get random color:
    randMin = 0
    randMax = UBound(colors)
    Randomize
    Dim randColor
    randColor = colors(Int((randMax - randMin + 1) * Rnd + randMin))
    
    'lblPassword.BackColor = "vb" & randColor
    
    ' Get random number between 0 and 9:
    randMax = 9
    Randomize
    Dim randNumber
    randNumber = Int((randMax - randMin + 1) * Rnd + randMin)

    ' Create the full random password, but in lower case
    Dim randomPassword
    randomPassword = LCase(currentDay & randColor & randNumber)

    lblPassword.Caption = randomPassword
    ChangeColor randColor
    ShowPenguins randNumber

End Sub

Sub ChangeColor(color)
    
    'Reset back to normal
    lblPassword.BackColor = &H8000000F
    lblPassword.ForeColor = vbBlack
    
    '"Red", "Orange", "Yellow", "Green", "Blue", "Purple", "Pink", "Brown", "Black", "White"
    Select Case color
        Case "Red"
            lblPassword.BackColor = vbRed
        Case "Orange"
            lblPassword.BackColor = &H80FF&    'VBOrange
        Case "Yellow"
            lblPassword.BackColor = vbYellow
            lblPassword.ForeColor = vbBlack
        Case "Green"
            lblPassword.BackColor = vbGreen
        Case "Blue"
            lblPassword.BackColor = vbBlue
            lblPassword.ForeColor = vbWhite
        Case "Purple"
            lblPassword.BackColor = &H800080   'VBPurple
            lblPassword.ForeColor = vbWhite
        Case "Pink"
            lblPassword.BackColor = &HFF80FF   'VBPink
        Case "Brown"
            lblPassword.BackColor = &H4040&    'VBBrown
            lblPassword.ForeColor = vbWhite
        Case "Black"
            lblPassword.BackColor = vbBlack
            lblPassword.ForeColor = vbWhite
        Case "White"
            lblPassword.BackColor = vbWhite
            lblPassword.ForeColor = vbBlack
        Case Else
        
    End Select

End Sub

Sub ShowPenguins(numOfP)
    
    'RESET - Set Visible to False
    For i = 0 To 8
        imgPenguin(i).Visible = False
    Next
    
    For i = 0 To numOfP - 1
        imgPenguin(i).Visible = True
    Next
    
End Sub

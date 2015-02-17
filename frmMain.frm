VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Checker"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   ForeColor       =   &H0080C0FF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIsCorrect 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   21
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLock 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lock &Text Boxes"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Lock Text Boxes"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdHide 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hide Text Boxes"
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Hide Text Boxes"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdFont 
      BackColor       =   &H00FFFF80&
      Caption         =   "Change &Font"
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Change Font"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1350
      TabIndex        =   16
      ToolTipText     =   "Please enter Igor here"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtIfIgor 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtIfBaird 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtIfWhiteRock 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtIfBC 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtProvince 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1350
      TabIndex        =   8
      ToolTipText     =   "Please enter BC here"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txtCity 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1350
      TabIndex        =   7
      ToolTipText     =   "Please enter White Rock here"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtStreet 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1350
      TabIndex        =   6
      ToolTipText     =   "Please enter Baird RD here"
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdBC 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check Province - BC"
      Height          =   735
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Check to see if the province is BC"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdBaird 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check Street - Baird Rd"
      Height          =   735
      Left            =   1525
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Check to see if the street is Baird Rd"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdWhiteRock 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check City - White Rock"
      Height          =   735
      Left            =   2735
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Check to see if the city is White Rock"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Height          =   615
      Left            =   360
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit Program"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.CommandButton cmdIgor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check Name - Igor"
      Height          =   735
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Check to see if the name is Igor"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   510
      TabIndex        =   17
      Top             =   600
      Width           =   465
   End
   Begin VB.Label lblStreet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Street:"
      Height          =   195
      Left            =   510
      TabIndex        =   11
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label lblCity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   195
      Left            =   510
      TabIndex        =   10
      Top             =   1530
      Width           =   300
   End
   Begin VB.Label lblProvince 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Province:"
      Height          =   195
      Left            =   510
      TabIndex        =   9
      Top             =   2010
      Width           =   675
   End
   Begin VB.Label lblIntructions 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your info below:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Correct()
txtIsCorrect.Text = "Correct!"
txtIsCorrect.ForeColor = &HFF00&
End Sub

Private Sub Incorrect()
MsgBox "Sorry, that is incorrect.", None, "Wrong!"
txtIsCorrect.Text = "Incorrect!"
txtIsCorrect.ForeColor = &HFF&
End Sub

Private Sub cmdIgor_Click()
If txtName.Text = "Igor" Then
    txtIfIgor.BackColor = &H80FF80
    Call Correct
Else
    txtIfIgor.BackColor = &HFF&
    Call Incorrect
End If
End Sub
Private Sub cmdBaird_Click()
If txtStreet.Text = "Baird Rd" Then
    txtIfBaird.BackColor = &H80FF80
    Call Correct
Else
    txtIfBaird.BackColor = &HFF&
    Call Incorrect
End If
End Sub
Private Sub cmdWhiteRock_Click()
If txtCity.Text = "White Rock" Then
    txtIfWhiteRock.BackColor = &H80FF80
    Call Correct
Else
    txtIfWhiteRock.BackColor = &HFF&
    Call Incorrect
End If
End Sub
Private Sub cmdBC_Click()
If txtProvince.Text = "BC" Then
    txtIfBC.BackColor = &H80FF80
    Call Correct
Else
    txtIfBC.BackColor = &HFF&
    Call Incorrect
End If
End Sub

Private Sub cmdLock_Click()
If txtName.Locked = False Then
    txtName.Locked = True
    txtStreet.Locked = True
    txtCity.Locked = True
    txtProvince.Locked = True
    cmdLock.Caption = "Unlock Text Boxes"
    cmdLock.ToolTipText = "Unlock Text Boxes"
    
ElseIf txtName.Locked = True Then
    txtName.Locked = False
    txtStreet.Locked = False
    txtCity.Locked = False
    txtProvince.Locked = False
    cmdLock.Caption = "Lock Text Boxes"
    cmdLock.ToolTipText = "Lock Text Boxes"
End If

End Sub

Private Sub cmdHide_Click()
    txtName.Visible = Not txtName.Visible
    txtStreet.Visible = Not txtStreet.Visible
    txtCity.Visible = Not txtCity.Visible
    txtProvince.Visible = Not txtProvince.Visible

If txtName.Visible = True Then
    cmdHide.Caption = "Hide Text Boxes"
    cmdHide.ToolTipText = "Hide Text Boxes"


ElseIf txtName.Visible = False Then
    cmdHide.Caption = "Show Text Boxes"
    cmdHide.ToolTipText = "Show Text Boxes"
End If

End Sub

Private Sub cmdFont_Click()
If txtName.Font = "Calibri" Then
    txtName.Font = "Comic Sans MS"
    txtStreet.Font = "Comic Sans MS"
    txtCity.Font = "Comic Sans MS"
    txtProvince.Font = "Comic Sans MS"
    
ElseIf txtName.Font = "Comic Sans MS" Then
    txtName.Font = "Calibri"
    txtStreet.Font = "Calibri"
    txtCity.Font = "Calibri"
    txtProvince.Font = "Calibri"
End If

End Sub

Private Sub cmdExit_Click()
MsgBox "Thank you for using my program!", None, "Goodbye!"
End
End Sub

Private Sub lblCorrect_Click()

End Sub


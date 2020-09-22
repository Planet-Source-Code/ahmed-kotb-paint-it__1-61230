VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter your Name"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4230
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You have make a heigh Score  Please enter your name."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
msg = "Are u sure that u dont want 2 write ur name ?"
msg = MsgBox(msg, vbCritical + vbYesNo)
If msg = vbYes Then Form2.Show
End Sub

Private Sub Command2_Click()
If Trim(Text1.Text) = "" Then
MsgBox "Please write ur name", vbCritical
Exit Sub
End If
percentage = Format(percentage, "0.00")
'Getscore
Select Case BruchWidth
Case 50
per.easy = Text1.Text
per.easyscr = percentage
Case 35
per.med = Text1.Text
per.medscr = percentage
Case 20
per.hard = Text1.Text
per.hardscr = percentage
End Select
Savescore
Scrstat = 1
Form3.Show
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = Form5.Icon
Label2.Caption = "Percentage = " & Format(percentage, "0.00") & " %"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub

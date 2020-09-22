VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint it"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   5880
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   4920
         Top             =   0
      End
      Begin VB.CommandButton Command5 
         Caption         =   "See Top percentages"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3840
         Top             =   0
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   4560
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4200
         Top             =   0
      End
      Begin VB.CommandButton Command3 
         Caption         =   "End Game"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start A new Game"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "###.###"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Percentage is :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   10
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   5775
      Left            =   0
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":27A2
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   600
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
   Begin VB.Shape Shp 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const a As Long = 14933984
Dim tot As Long
Dim c As Long


Dim tmp As Single

Dim cont As Integer

Private Sub Command1_Click()
c = 0
tot = pic.ScaleHeight * pic.ScaleHeight
'MsgBox tot
For i = 0 To pic.ScaleWidth
     For i2 = 0 To pic.ScaleHeight
         'If pic.Point(i, i2) <> -1 And pic.Point(i, i2) <> a Then c = c + 1
         If pic.Point(i, i2) = 255 And pic.Point(i, i2) <> -1 Then c = c + 1
     Next i2
Next i
'MsgBox c
percentage = (c / tot) * 100
'MsgBox percentage
     
End Sub


Private Sub Command2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command3_Click()
MsgBox "Thank u For Using My Game" & vbCrLf & "Kotb Corp. 2006", vbInformation
End
End Sub

Private Sub Command4_Click()
Command2_Click
End Sub

Private Sub Command5_Click()
Scrstat = 0
Form3.Show vbModal
End Sub

Private Sub Form_Load()
Me.Icon = Form5.Icon
Status = True
cont = 0
Frame1.ZOrder 1
pic.DrawWidth = BruchWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
Command4_Click
End Sub

Private Sub Pic_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Status = True Then pic.PSet (X, Y)
'pic.PSet (1, 1)
'MsgBox pic.Point(1, 1)

End Sub


'Private Sub Pic_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pic.PSet (1, 1)
'MsgBox pic.Point(1, 1)
'pic.PSet (X, Y)
'End Sub
Private Sub Timer1_Timer()
cont = cont + 1
If cont = 11 Then
Status = False
Frame1.Top = -Frame1.Height
Frame1.ZOrder 0
Timer2.Enabled = True
Command1_Click
Timer1.Enabled = False
Else
Shp(cont - 1).BackColor = vbYellow
End If

End Sub

Private Sub Timer2_Timer()
Frame1.Top = Frame1.Top + 50
If Frame1.Top >= 80 Then
Timer2.Enabled = False
Frame1.Top = 80
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
tmp = tmp + 5.5
Label3.Caption = Format(tmp, "0.00") & " %"
If tmp >= percentage Then
Label3.Caption = Format(percentage, "0.00") & " %"
Command2.Enabled = True
Command3.Enabled = True
Command5.Enabled = True
Label3.FontSize = 18
If BruchWidth = 50 Then tmp = per.easyscr
If BruchWidth = 35 Then tmp = per.medscr
If BruchWidth = 20 Then tmp = per.hardscr


If percentage > tmp Then
Timer4.Enabled = True
End If

Timer3.Enabled = False
End If

End Sub

Private Sub Timer4_Timer()
Form4.Show
Unload Me
End Sub

Private Sub Timer5_Timer()

End Sub

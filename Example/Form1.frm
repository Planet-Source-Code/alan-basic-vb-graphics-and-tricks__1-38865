VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Grafix Example"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Me"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      MousePointer    =   2  'Cross
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1440
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line6 
      X1              =   3600
      X2              =   2760
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      X1              =   3480
      X2              =   2640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   2880
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Grafix Example"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      MousePointer    =   15  'Size All
      TabIndex        =   14
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin VB.Shape Shape10 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB Example"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1440
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "This is just a little example on how you can spice up your programs.    By: Alan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   1440
      Top             =   360
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Form5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Form4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Form3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Form2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Form1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Shape S1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape S5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape S4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape S3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape S2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   3120
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   0
      Width           =   225
   End
   Begin VB.Shape Shape7 
      Height          =   255
      Left            =   3120
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape6 
      Height          =   255
      Left            =   2880
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Top             =   480
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MousePointer    =   15  'Size All
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label14.Caption = Winsock1.LocalIP
Form2.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape11.Visible = False
Label8.BackColor = &H808080
Label9.BackColor = &H808080

Status.Caption = "Status"
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False
Label8.BackStyle = 0
Label9.BackStyle = 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Status"
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape11.Visible = False
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape11.Visible = False
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = &H808080
Label9.BackColor = &H808080
End Sub

Private Sub Label13_Click()
If Label13.BorderStyle = 0 Then
Timer3.Enabled = True
Label13.BorderStyle = 1
Else
Timer3.Enabled = False
Label13.BorderStyle = 0

End If
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape11.Visible = True
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Form1

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = &H808080
Label9.BackColor = &H808080
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Form1
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Status"
End Sub

Private Sub Label3_Click()
MsgBox "Hello " & Winsock1.LocalIP & "(" & Winsock1.LocalHostName & "), The Time Is: " & Time, vbInformation, "VB Example"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Little Info About Form1"
S1.Visible = True
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False

End Sub

Private Sub Label4_Click()
MsgBox "Hello " & Winsock1.LocalIP & "(" & Winsock1.LocalHostName & "), The Time Is: " & Time, vbInformation, "VB Example"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Little Info About Form5"
S5.Visible = True
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = False

End Sub

Private Sub Label5_Click()
MsgBox "Hello " & Winsock1.LocalIP & "(" & Winsock1.LocalHostName & "), The Time Is: " & Time, vbInformation, "VB Example"
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Little Info About Form4"
S4.Visible = True
S1.Visible = False
S2.Visible = False
S3.Visible = False
S5.Visible = False

End Sub

Private Sub Label6_Click()
MsgBox "Hello " & Winsock1.LocalIP & "(" & Winsock1.LocalHostName & "), The Time Is: " & Time, vbInformation, "VB Example"
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Little Info About Form3"
S3.Visible = True
S1.Visible = False
S2.Visible = False
S4.Visible = False
S5.Visible = False

End Sub

Private Sub Label7_Click()
MsgBox "Hello " & Winsock1.LocalIP & "(" & Winsock1.LocalHostName & "), The Time Is: " & Time, vbInformation, "VB Example"
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.Caption = "Little Info About Form2"
S2.Visible = True
S1.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False

End Sub

Private Sub Label8_Click()
Form1.WindowState = 1
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = &HE0E0E0
Label8.BackStyle = 1
Label9.BackColor = &H808080
Label9.BackStyle = 0
End Sub

Private Sub Label9_Click()
Form_ExitDown Form1
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = &HE0E0E0
Label8.BackColor = &H808080
Label9.BackStyle = 1
Label8.BackStyle = 0
End Sub

Private Sub Status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape11.Visible = False
Status.Caption = "Status"

S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False
End Sub

Private Sub Timer1_Timer()
Label10.Caption = Time
End Sub

Private Sub Timer2_Timer()
Form1.BackColor = vbWhite
Sleep 1
Form1.BackColor = &HE0E0E0
Sleep 1
End Sub

Private Sub Timer3_Timer()
S1.Visible = True
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = False
Sleep 0.2
S1.Visible = False
S2.Visible = True
S3.Visible = False
S4.Visible = False
S5.Visible = False
Sleep 0.2
S1.Visible = False
S2.Visible = False
S3.Visible = True
S4.Visible = False
S5.Visible = False
Sleep 0.2
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = True
S5.Visible = False
Sleep 0.2
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = False
S5.Visible = True
Sleep 0.2
S1.Visible = False
S2.Visible = False
S3.Visible = False
S4.Visible = True
S5.Visible = False
Sleep 0.2
S1.Visible = False
S2.Visible = False
S3.Visible = True
S4.Visible = False
S5.Visible = False
Sleep 0.2
S1.Visible = False
S2.Visible = True
S3.Visible = False
S4.Visible = False
S5.Visible = False
Sleep 0.2
End Sub

VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   28
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      TabIndex        =   24
      Top             =   5880
      Width           =   5415
      Begin VB.OptionButton Option15 
         Height          =   495
         Left            =   480
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option14 
         Height          =   495
         Left            =   2280
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option13 
         Height          =   495
         Left            =   4080
         TabIndex        =   25
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      TabIndex        =   20
      Top             =   4920
      Width           =   5415
      Begin VB.OptionButton Option12 
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option11 
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option10 
         Height          =   495
         Left            =   4080
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      TabIndex        =   16
      Top             =   3960
      Width           =   5415
      Begin VB.OptionButton Option7 
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option9 
         Height          =   495
         Left            =   4080
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      TabIndex        =   12
      Top             =   3000
      Width           =   5415
      Begin VB.OptionButton Option6 
         Height          =   495
         Left            =   4080
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Height          =   495
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      TabIndex        =   8
      Top             =   2040
      Width           =   5415
      Begin VB.OptionButton Option1 
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Q1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Q5"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   6150
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Q4"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   5205
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Q3"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4245
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Q2"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "POOR"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   8640
      TabIndex        =   2
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "AVERAGE"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   6600
      TabIndex        =   1
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   5040
      TabIndex        =   0
      Top             =   600
      Width           =   930
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset
Dim A, B, C, d, E As String

Private Sub Form_Load()
Set ADO = New ADODB.Connection
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=d:\New folder (3)\Database1.mdb;"
ADO.Open

End Sub
Private Sub Command1_Click()


If Option1.Value = True Then
A = "GOOD"
ElseIf Option2.Value = True Then
A = "AVG"
ElseIf Option3.Value = True Then
A = "BAD"
End If

If Option4.Value = True Then
B = "GOOD"
ElseIf Option5.Value = True Then
B = "AVG"
ElseIf Option6.Value = True Then
B = "BAD"
End If

If Option7.Value = True Then
C = "GOOD"
ElseIf Option8.Value = True Then
C = "AVG"
ElseIf Option9.Value = True Then
C = "BAD"
End If

If Option12.Value = True Then
d = "GOOD"
ElseIf Option11.Value = True Then
d = "AVG"
ElseIf Option10.Value = True Then
d = "BAD"
End If

If Option15.Value = True Then
E = "GOOD"
ElseIf Option14.Value = True Then
E = "AVG"
ElseIf Option13.Value = True Then
E = "BAD"
End If
Set RS = New ADODB.Recordset

Dim str10 As String
str10 = "UPDATE TABLE5 SET F1='" + A + "',F2='" + B + "',F3='" + C + "',F4='" + d + "',F5='" + E + "' where TRAINING_NAME='" + Form3.Combo2.Text + "' and TRAINING_TYPE='" + Form3.Combo5.Text + "'"
Set RS = New ADODB.Recordset
RS.Open str10, ADO, adOpenStatic, adLockOptimistic
MsgBox "Data Inserted Successfully"
Form3.Check1.Value = 0
Form9.Hide
Form4.Show

End Sub


VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8970
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14805
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   525
      Left            =   8520
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   525
      Left            =   8520
      TabIndex        =   13
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DETAILS OF THE PROGRAMME"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   11
      Top             =   8160
      Width           =   5775
   End
   Begin VB.TextBox Text5 
      Height          =   1335
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   8520
      TabIndex        =   8
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "FACULTY NAME"
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
      Left            =   480
      TabIndex        =   16
      Top             =   240
      Width           =   2520
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "MINIMUM NUMBER OF PARTICIPANTS"
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
      Left            =   480
      TabIndex        =   15
      Top             =   960
      Width           =   6060
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "CONDUCTED/NONCONDUCTED"
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
      Left            =   480
      TabIndex        =   10
      Top             =   7440
      Width           =   4920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "TRAINING DESCRIPTION"
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
      Left            =   480
      TabIndex        =   4
      Top             =   5760
      Width           =   3870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "NUMBER OF DAYS"
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
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   2925
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "TRAINING TYPE"
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
      Left            =   480
      TabIndex        =   2
      Top             =   3720
      Width           =   2520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "PROGRAMME NAME"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   3300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "TRAINING_ID"
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
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   2040
   End
   Begin VB.Menu N 
      Caption         =   "SCHEDULES"
      Begin VB.Menu RES 
         Caption         =   "RESIDENTIAL"
         Shortcut        =   ^R
      End
      Begin VB.Menu NON 
         Caption         =   "NON-RESIDENTIAL"
         Shortcut        =   ^N
      End
      Begin VB.Menu ADD 
         Caption         =   "ADDITONAL"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu A 
      Caption         =   "ABOUT"
   End
   Begin VB.Menu E 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset

Private Sub A_Click()
Form7.Show
End Sub

Private Sub ADD_Click()
Form6.Show
End Sub

Private Sub Command1_Click()
Form8.Show
End Sub

Private Sub E_Click()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
RS.Close
ADO.Close
End Sub

Private Sub NON_Click()
Form5.Show
End Sub

Private Sub RES_Click()
Form4.Show
End Sub

Private Sub Text1_Change()
Set ADO = New ADODB.Connection
Set RS = New ADODB.Recordset
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Pranay\Desktop\New folder (3)\Database1.mdb;"
ADO.Open

Dim str1 As String
str1 = "select * from table2 where id='" + Text1.Text + "'"
RS.Open str1, ADO, adOpenStatic, adLockOptimistic
Set Text2.DataSource = RS
Text2.DataField = "pname"
Set Text3.DataSource = RS
Text3.DataField = "ttype"
Set Text4.DataSource = RS
Text4.DataField = "ndays"
Set Text5.DataSource = RS
Text5.DataField = "desc"
End Sub

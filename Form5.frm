VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form5 
   Caption         =   "SCHEDULE NON-RESIDENT"
   ClientHeight    =   9375
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13365
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   9375
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   12135
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   6720
         TabIndex        =   31
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   525
         Left            =   6720
         TabIndex        =   21
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   525
         Left            =   6720
         TabIndex        =   20
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   525
         Left            =   6720
         TabIndex        =   19
         Top             =   4920
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   525
         Left            =   6720
         TabIndex        =   18
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   6720
         TabIndex        =   17
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00400000&
         Caption         =   "FEEDBACK"
         Enabled         =   0   'False
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
         Left            =   9360
         TabIndex        =   16
         Top             =   6000
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   6720
         TabIndex        =   15
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "CITY"
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
         Left            =   360
         TabIndex        =   30
         Top             =   3120
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING NAME"
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
         Left            =   360
         TabIndex        =   27
         Top             =   1200
         Width           =   2670
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING DATE"
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
         Left            =   360
         TabIndex        =   26
         Top             =   2280
         Width           =   2595
      End
      Begin VB.Label Label7 
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
         Left            =   360
         TabIndex        =   25
         Top             =   3960
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
         Left            =   360
         TabIndex        =   24
         Top             =   5040
         Width           =   6060
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "CONDUCTED/NOT CONDUCTED"
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
         Left            =   360
         TabIndex        =   23
         Top             =   6120
         Width           =   4995
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING ID"
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
         TabIndex        =   22
         Top             =   360
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   12135
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4080
         TabIndex        =   29
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   525
         Left            =   4080
         TabIndex        =   7
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   6000
         Picture         =   "Form5.frx":A4382
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   525
         Left            =   4080
         TabIndex        =   5
         Top             =   5400
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   525
         Left            =   4080
         TabIndex        =   4
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   4080
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SUBMIT"
         Height          =   435
         Left            =   4320
         TabIndex        =   2
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   4080
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin MSComCtl2.MonthView MonthView1 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMMM yyyy dddd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2370
         Left            =   5880
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   93454337
         CurrentDate     =   40408
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "CITY"
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
         Left            =   360
         TabIndex        =   28
         Top             =   3360
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "MINIMUM NUMBER OF PARTICIPANTS"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   5400
         Width           =   2775
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   12
         Top             =   4320
         Width           =   2520
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING DATE"
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
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING NAME"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   2670
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING ID"
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
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   2040
      End
   End
   Begin VB.Menu P 
      Caption         =   "PROGRAMS"
      Begin VB.Menu AP 
         Caption         =   "ADD PROGRAMS"
         Shortcut        =   ^A
      End
      Begin VB.Menu EP 
         Caption         =   "EXISTING PROGRAMS"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu APP 
      Caption         =   "ADD PARTICIPANTS"
   End
   Begin VB.Menu E 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim flag1 As Integer
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset
Dim ado1 As ADODB.Connection
Dim rs1 As ADODB.Recordset


Private Sub AP_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub
Private Sub APP_Click()
Form4.Hide
Form10.Show
End Sub
Private Sub Check1_Click()
If flag1 = 0 Then
If Check1.Value = 1 Then
Command2.Enabled = True
flag1 = 1
End If
ElseIf Check1.Value = 0 Then
Command2.Enabled = False
flag1 = 0
End If
End Sub
Private Sub Command1_Click()
If flag = 0 Then
MonthView1.Visible = True
flag = 1
Else
MonthView1.Visible = False
flag = 0
End If
Text3.SetFocus
End Sub

Private Sub Command2_Click()
Form9.Show
Form4.Hide

End Sub

Private Sub Command3_Click()
Dim str1 As String
str1 = "INSERT INTO TABLE7 VALUES ('" + Text10.Text + "','" + Text1.Text + "','" + Text2.Text + "','" + Combo1.Text + "','" + Text3.Text + "'," + Text4.Text + ")"
RS.Open str1, ADO, adOpenStatic, adLockOptimistic
MsgBox "Data Inserted Successfully"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Frame1.Visible = False
End Sub
Private Sub EP_Click()
Frame2.Visible = True
Frame1.Visible = False


End Sub
Private Sub EX_Click()
End
End Sub
Private Sub Form_Load()
flag = 0
flag1 = 0
Set ADO = New ADODB.Connection
Set RS = New ADODB.Recordset

ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Pranay\Desktop\New folder (3)\Database1.mdb;"
ADO.Open
End Sub

Private Sub Form_Unload(Cancel As Integer)
RS.Close
ADO.Close
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text2 = MonthView1.Value
End Sub
Private Sub text9_change()


Dim STR2 As String
STR2 = "select * from table7 where TRAINING_ID= '" + Text9.Text + "'"
Set rs1 = New ADODB.Recordset
rs1.Open STR2, ADO, adOpenStatic, adLockOptimistic
Set Text5.DataSource = rs1
Text5.DataField = "TRAINING_NAME"
Set Text8.DataSource = rs1
Text8.DataField = "TRAINING_DATE"
Set Text11.DataSource = rs1
Text11.DataField = "CITY"
Set Text6.DataSource = rs1
Text6.DataField = "FACULTY_NAME"
Set Text7.DataSource = rs1
Text7.DataField = MINIMUM_NUMBER_OF_PARTICIPANTS
End Sub




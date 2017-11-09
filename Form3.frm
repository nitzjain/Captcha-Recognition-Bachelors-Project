VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9075
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11760
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   16320
      TabIndex        =   32
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   16320
      TabIndex        =   30
      Top             =   3840
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   960
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   495
         Left            =   4200
         TabIndex        =   24
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   495
         Left            =   4200
         TabIndex        =   23
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   495
         Left            =   4200
         TabIndex        =   22
         Top             =   3840
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form3.frx":A4382
         Left            =   4200
         List            =   "Form3.frx":A43EF
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form3.frx":A45BE
         Left            =   4200
         List            =   "Form3.frx":A462E
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "PARTICIPANT NAME"
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
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   3300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "COMPANY"
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
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "STATE"
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
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "CONTACT NUMBER"
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
         Left            =   120
         TabIndex        =   25
         Top             =   3960
         Width           =   3060
      End
   End
   Begin VB.CommandButton Command5 
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
      Left            =   6600
      TabIndex        =   18
      Top             =   10440
      Width           =   1935
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   16320
      TabIndex        =   16
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   16320
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form3.frx":A4789
      Left            =   15840
      List            =   "Form3.frx":A478B
      TabIndex        =   11
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
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
      Left            =   11040
      TabIndex        =   9
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MENU"
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
      Left            =   11040
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
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
      Left            =   9000
      TabIndex        =   7
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
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
      Left            =   8880
      TabIndex        =   6
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label14 
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
      Left            =   10560
      TabIndex        =   33
      Top             =   3000
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Label Label13 
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
      Left            =   10560
      TabIndex        =   31
      Top             =   3840
      Width           =   4995
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "NUMBER_OF DAYS"
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
      Left            =   10560
      TabIndex        =   14
      Top             =   2160
      Width           =   2925
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "TRAINING_CITY"
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
      Left            =   10560
      TabIndex        =   13
      Top             =   1440
      Width           =   2385
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "TRAINING_NAME"
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
      Left            =   10560
      TabIndex        =   10
      Top             =   480
      Width           =   2670
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "DESCRIPTION"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2595
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2520
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2040
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs6 As ADODB.Recordset
Dim rs5 As ADODB.Recordset
Dim rs7 As ADODB.Recordset



Private Sub Check1_Click()
If Check1.Value = 0 Then
Command2.Enabled = False
Frame1.Visible = FASLE
Command5.Enabled = False
Command1.Enabled = False
Label14.Visible = False
Text2.Visible = False

Else
Command2.Enabled = True
Frame1.Visible = True
Command5.Enabled = True
Command1.Enabled = True
Label14.Visible = True
Text2.Visible = True

End If
End Sub

Private Sub Combo2_Click()

Dim str6 As String
str6 = "select distinct TRAINING_TYPE from table4 where TRAINING_NAME= '" + Combo2.Text + "'"
Set rs7 = New ADODB.Recordset
rs7.Open str6, ADO, adOpenStatic, adLockOptimistic
Do Until rs7.EOF
Combo5.AddItem rs7!TRAINING_TYPE
rs7.MoveNext
Loop
End Sub


Private Sub Combo3_Click()
If Combo5.Text <> "" Then
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text9.Enabled = True
Combo1.Enabled = True
Command1.Enabled = True
Else
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False

Text9.Enabled = False
Combo1.Enabled = False
Command1.Enabled = False

End If

End Sub



Private Sub Combo5_Click()
Combo3.Clear
Dim STR2 As String
STR2 = "select * from table4 where TRAINING_NAME= '" + Combo2.Text + "'AND TRAINING_TYPE='" + Combo5.Text + "'"
Set rs5 = New ADODB.Recordset
rs5.Open STR2, ADO, adOpenStatic, adLockOptimistic
Do Until rs5.EOF
Combo3.AddItem rs5!TRAINING_DATE
rs5.MoveNext
Loop
Set rs1 = New ADODB.Recordset
rs1.Open STR2, ADO, adOpenStatic, adLockOptimistic
Set Text1.DataSource = rs1
Text1.DataField = "TRAINING_ID"
Set Text3.DataSource = rs1
Text3.DataField = "NUMBER_OF_DAYS"
Set Text5.DataSource = rs1
Text5.DataField = "DESCRIPTION"
Set Text5.DataSource = rs1
Text5.DataField = ""
If Combo5.Text = "Residential" Then
Label11.Visible = False
Text7.Visible = False
Else
Label11.Visible = True
Text7.Visible = True
Set Text7.DataSource = rs1
Text7.DataField = "CITY"

End If

End Sub

Private Sub Command5_Click()
Form9.Show
Form3.Hide
End Sub

Private Sub Form_Load()
Combo2.Clear
Combo5.Clear
Set ADO = New ADODB.Connection
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=d:\New folder (3)\Database1.mdb;"
ADO.Open
Dim str3 As String
str3 = "select DISTINCT TRAINING_NAME from table4 "
Set rs6 = New ADODB.Recordset
rs6.Open str3, ADO, adOpenStatic, adLockOptimistic
Do Until rs6.EOF
Combo2.AddItem rs6!TRAINING_NAME
rs6.MoveNext
Loop

End Sub
Private Sub Command1_Click()
Set rs1 = New ADODB.Recordset
Dim str1 As String
If Combo5.Text = "RESIDENTIAL" Then
str1 = "INSERT INTO TABLE5(TRAINING_ID,TRAINING_TYPE,TRAINING_NAME,TRAINING_DATE,TRAINING_CITY,NUMBER_OF_DAYS,PARTICIPANT_NAME,COMPANY,CITY,STATE,CONTACT_NUMBER) VALUES ('" + Text1.Text + "','" + Combo5.Text + "','" + Combo2.Text + "','" + Combo3.Text + "','" + Null + "'," + Text3.Text + ",'" + Text4.Text + "','" + Text6.Text + "','" + Combo4.Text + "','" + Combo1.Text + "'," + Text9.Text + ")"
Else
str1 = "INSERT INTO TABLE5(TRAINING_ID,TRAINING_TYPE,TRAINING_NAME,TRAINING_DATE,TRAINING_CITY,NUMBER_OF_DAYS,PARTICIPANT_NAME,COMPANY,CITY,STATE,CONTACT_NUMBER) VALUES ('" + Text1.Text + "','" + Combo5.Text + "','" + Combo2.Text + "','" + Combo3.Text + "','" + Text7.Text + "'," + Text3.Text + ",'" + Text4.Text + "','" + Text6.Text + "','" + Combo4.Text + "','" + Combo1.Text + "'," + Text9.Text + ")"
End If
rs1.Open str1, ADO, adOpenStatic, adLockOptimistic
MsgBox "Data Inserted Successfully"
Text4 = ""
Text6 = ""
Text9 = ""
Combo1.Text = ""
Combo4.Text = ""


End Sub

Private Sub Command2_Click()
Text1 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text9 = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo5.Text = ""
Combo4.Text = ""

End Sub

Private Sub Command3_Click()
Form4.Show
Form3.Hide
End Sub

Private Sub Command4_Click()
End
End Sub



VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form4 
   Caption         =   "PROGRAMS"
   ClientHeight    =   11070
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11760
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   11070
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   8535
      Left            =   120
      TabIndex        =   47
      Top             =   480
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox Text25 
         Height          =   495
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text24 
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text23 
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text22 
         Height          =   495
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         Height          =   495
         Left            =   0
         TabIndex        =   52
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         ItemData        =   "Form4.frx":A4382
         Left            =   1440
         List            =   "Form4.frx":A4392
         TabIndex        =   50
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         ItemData        =   "Form4.frx":A43C5
         Left            =   8160
         List            =   "Form4.frx":A43C7
         TabIndex        =   49
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         ItemData        =   "Form4.frx":A43C9
         Left            =   4920
         List            =   "Form4.frx":A43CB
         TabIndex        =   48
         Top             =   1200
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2055
         Left            =   1440
         TabIndex        =   51
         Top             =   2400
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3625
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   17295
      Begin VB.TextBox Text20 
         Height          =   495
         Left            =   12000
         TabIndex        =   46
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   495
         Left            =   10800
         TabIndex        =   45
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   495
         Left            =   9600
         TabIndex        =   44
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   495
         Left            =   8400
         TabIndex        =   43
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Height          =   495
         Left            =   7200
         TabIndex        =   42
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   6000
         TabIndex        =   41
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   4800
         TabIndex        =   40
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   3600
         TabIndex        =   39
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   2400
         TabIndex        =   38
         Top             =   7320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1200
         TabIndex        =   37
         Top             =   7320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2415
         Left            =   0
         TabIndex        =   14
         Top             =   4440
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedCols       =   0
         BackColorFixed  =   16777215
         BackColorSel    =   -2147483643
         BackColorBkg    =   4194304
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   360
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         ItemData        =   "Form4.frx":A43CD
         Left            =   3960
         List            =   "Form4.frx":A43CF
         TabIndex        =   34
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form4.frx":A43D1
         Left            =   5880
         List            =   "Form4.frx":A43D3
         TabIndex        =   33
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text12 
         Height          =   855
         Left            =   13800
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   13800
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   13800
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   13800
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form4.frx":A43D5
         Left            =   3960
         List            =   "Form4.frx":A43D7
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   3960
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label19 
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
         Left            =   8400
         TabIndex        =   13
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label17 
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
         Left            =   8400
         TabIndex        =   9
         Top             =   2040
         Width           =   2925
      End
      Begin VB.Label Label15 
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
         Left            =   8400
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "TRAINING  TYPE"
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
         Left            =   8400
         TabIndex        =   6
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "PROGRAM NAME"
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
         TabIndex        =   4
         Top             =   1200
         Width           =   2805
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
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   2040
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
         TabIndex        =   1
         Top             =   2280
         Width           =   2595
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   16815
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   4080
         TabIndex        =   24
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   5880
         Picture         =   "Form4.frx":A43D9
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Command3 
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
         Height          =   435
         Left            =   4320
         TabIndex        =   22
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   120
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3960
         TabIndex        =   20
         Top             =   1320
         Width           =   3255
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form4.frx":A45D9
         Left            =   13200
         List            =   "Form4.frx":A45E6
         TabIndex        =   19
         Top             =   120
         Width           =   2055
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "Form4.frx":A4614
         Left            =   13200
         List            =   "Form4.frx":A4681
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   13200
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Height          =   855
         Left            =   13200
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   3240
         Width           =   2895
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
         Left            =   6000
         TabIndex        =   25
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "PROGRAM TYPE"
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
         Left            =   9240
         TabIndex        =   29
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label14 
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
         Left            =   9240
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label16 
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
         Left            =   9240
         TabIndex        =   27
         Top             =   2400
         Width           =   2925
      End
      Begin VB.Label Label18 
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
         Left            =   9240
         TabIndex        =   26
         Top             =   3480
         Width           =   2175
      End
   End
   Begin VB.Menu P 
      Caption         =   "PROGRAMS"
      Begin VB.Menu AP 
         Caption         =   "ADD PROGRAM"
         Shortcut        =   ^A
      End
      Begin VB.Menu EP 
         Caption         =   "EXISTING PROGRAM"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu APP 
      Caption         =   "ADD PARTICIPANTS"
   End
   Begin VB.Menu CTE 
      Caption         =   "CONVERT TO EXCEL"
   End
   Begin VB.Menu A 
      Caption         =   "ABOUT"
   End
   Begin VB.Menu EX 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i1, i2, c1, c2 As Integer
Dim l As Double
Dim j, k, j1, k1 As String
Dim flag As Integer
Dim flag1 As Integer
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim RS9 As ADODB.Recordset
Dim RS11 As ADODB.Recordset
Dim RS10 As ADODB.Recordset
Dim d As Date
Dim d1, D2 As Date
Dim s1, S3 As Integer
Dim s2 As Double
Dim s As String
Dim A1, a2 As Integer
Dim nj As Date
Private Sub Combo10_Click()

Dim rs6 As ADODB.Recordset

Combo4.Clear


Set rs6 = New ADODB.Recordset
rs6.Open "select * from table4 WHERE TRAINING_NAME='" + Combo1.Text + "' AND TRAINING_TYPE='" + Combo5.Text + "'", ADO, adOpenStatic, adLockOptimistic
Do Until rs6.EOF
Set Text4.DataSource = rs6
Text4.DataField = "TRAINING_DATE"
d1 = Text4.Text
s2 = Year(d1)
l = Val(Combo10.Text)
'MsgBox (s2)
'MsgBox (l)
If s2 = l Then
Combo4.AddItem rs6!TRAINING_DATE
rs6.MoveNext
Else
rs6.MoveNext
End If
Loop



End Sub

Private Sub A_Click()
Form7.Show
End Sub
Private Sub AP_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Combo2.Clear
Dim STR9 As String
STR9 = "select distinct TRAINING_NAME from table4 "
Set RS9 = New ADODB.Recordset
RS9.Open STR9, ADO, adOpenStatic, adLockOptimistic
Do Until RS9.EOF
Combo2.AddItem RS9!TRAINING_NAME
RS9.MoveNext
Loop

End Sub
Private Sub APP_Click()
Form4.Hide
Form3.Show
End Sub


Private Sub CAL_Click()
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False

End Sub

Private Sub Combo1_Click()
Combo4.Clear
Text5 = ""
Text8 = ""
Text9 = ""
Text12 = ""
Combo5.Clear
Combo10.Clear

Dim str5 As String
str5 = "select distinct TRAINING_TYPE from table4 where TRAINING_NAME= '" + Combo1.Text + "'"
Set rs5 = New ADODB.Recordset
rs5.Open str5, ADO, adOpenStatic, adLockOptimistic
Do Until rs5.EOF
Combo5.AddItem rs5!TRAINING_TYPE
rs5.MoveNext
Loop

End Sub




Private Sub Combo4_CLICK()
'MSHFlexGrid1.Clear
MSHFlexGrid1.Enabled = True
Dim rs5 As ADODB.Recordset
Set rs5 = New ADODB.Recordset
'On Error GoTo L2
'nj = Format(Combo4.Text, "mm/dd/yyyy")
rs5.Open "select * from table5 WHERE TRAINING_NAME='" + Combo1.Text + "' AND TRAINING_TYPE='" + Combo5.Text + "' ", ADO, adOpenStatic, adLockOptimistic
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 0
MSHFlexGrid1.Text = "PARTICIPANT_NAME"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 1
MSHFlexGrid1.Text = "COMPANY"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 2
MSHFlexGrid1.Text = "CITY"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 3
MSHFlexGrid1.Text = "STATE"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 4
MSHFlexGrid1.Text = "CONTACT_NUMBER"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 5
MSHFlexGrid1.Text = "F1"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 6
MSHFlexGrid1.Text = "F2"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 7
MSHFlexGrid1.Text = "F3"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 8
MSHFlexGrid1.Text = "F4"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 9
MSHFlexGrid1.Text = "F5"

MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1500

I = 1
Do Until rs5.EOF
Set Text6.DataSource = rs5
Text6.DataField = "PARTICIPANT_NAME"
Set Text7.DataSource = rs5
Text7.DataField = "COMPANY"
Set Text13.DataSource = rs5
Text13.DataField = "CITY"
Set Text14.DataSource = rs5
Text14.DataField = "STATE"
Set Text15.DataSource = rs5
Text15.DataField = "CONTACT_NUMBER"
Set Text16.DataSource = rs5
Text16.DataField = "F1"
Set Text17.DataSource = rs5
Text17.DataField = "F2"
Set Text18.DataSource = rs5
Text18.DataField = "F3"
Set Text19.DataSource = rs5
Text19.DataField = "F4"
Set Text20.DataSource = rs5
Text20.DataField = "F5"
MSHFlexGrid1.Row = I
MSHFlexGrid1.Col = 0
MSHFlexGrid1.Text = Text6
MSHFlexGrid1.Col = 1
MSHFlexGrid1.Text = Text7
MSHFlexGrid1.Col = 2
MSHFlexGrid1.Text = Text13

MSHFlexGrid1.Col = 3
MSHFlexGrid1.Text = Text14

MSHFlexGrid1.Col = 4
MSHFlexGrid1.Text = Text15


MSHFlexGrid1.Col = 5
MSHFlexGrid1.Text = Text16
MSHFlexGrid1.Col = 6
MSHFlexGrid1.Text = Text17
MSHFlexGrid1.Col = 7
MSHFlexGrid1.Text = Text18
MSHFlexGrid1.Col = 8
MSHFlexGrid1.Text = Text19

MSHFlexGrid1.Col = 9
MSHFlexGrid1.Text = Text20

MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
I = I + 1
'Set MSHFlexGrid1.DataSource = rs5
rs5.MoveNext
Loop
'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1

End Sub

Private Sub Combo5_Click()
Combo4.Clear
Combo10.Clear
If Combo5.Text = "Residential" Then
Label15.Visible = False
Text8.Visible = False
Else
Label15.Visible = True
Text8.Visible = True
Set Text8.DataSource = rs1
Text8.DataField = "CITY"
End If
Dim STR2 As String
STR2 = "select * from table4 where TRAINING_TYPE= '" + Combo5.Text + "' AND TRAINING_NAME='" + Combo1.Text + "'"

Set rs1 = New ADODB.Recordset
rs1.Open STR2, ADO, adOpenStatic, adLockOptimistic
Set Text9.DataSource = rs1
Text9.DataField = "TRAINING_ID"
Set Text5.DataSource = rs1
Text5.DataField = "NUMBER_OF_DAYS"
Set Text8.DataSource = rs1
Text8.DataField = "CITY"
Set Text12.DataSource = rs1
Text12.DataField = "DESCRIPTION"
Set rs4 = New ADODB.Recordset
rs4.Open STR2, ADO, adOpenStatic, adLockOptimistic

Do Until rs4.EOF
Set Text3.DataSource = rs4
Text3.DataField = "TRAINING_DATE"
d = Text3
s1 = Year(d)
Combo10.AddItem (s1)
rs4.MoveNext
Loop

c1 = Combo10.ListCount
For i1 = 0 To (c1 - 1)
j = Combo10.List(i1)
'c1 = Combo10.ListCount
For A1 = i1 + 1 To (c1 - 1)
'c1 = Combo10.ListCount

k = Combo10.List(A1)
If (k = j) Then
Combo10.RemoveItem (A1)
c1 = Combo10.ListCount

End If

Next A1
'c1 = Combo10.ListCount

Next i1

Combo10.Text = s1
End Sub
Private Sub Combo3_Click()
If Combo3.Text = "Residential" Then
Label14.Visible = False
Combo6.Visible = False
Text1 = 2
Else
Label14.Visible = True
Combo6.Visible = True
Text1 = 1
End If
Set RS10 = New ADODB.Recordset
RS10.Open "select * from table4 WHERE TRAINING_NAME='" + Combo2.Text + "' AND TRAINING_TYPE='" + Combo3.Text + "'", ADO, adOpenStatic, adLockOptimistic
Set Text10.DataSource = RS10
Text10.DataField = "TRAINING_ID"
End Sub





Private Sub Combo7_Click()
Dim A As String

Dim RS12 As ADODB.Recordset
Combo9.Clear

Set RS12 = New ADODB.Recordset

If Combo7.Text = "All" Then
A = "SELECT * FROM TABLE4"
Else
A = "SELECT * FROM TABLE4 WHERE TRAINING_TYPE='" + Combo7.Text + "'"
End If
RS12.Open A, ADO, adOpenStatic, adLockOptimistic
Do Until RS12.EOF
Set Text25.DataSource = RS12
Text25.DataField = "TRAINING_DATE"
D2 = Text25
S3 = Year(D2)
Combo9.AddItem (S3)
RS12.MoveNext
Loop

c2 = Combo9.ListCount
For i2 = 0 To (c2 - 1)
j1 = Combo9.List(i2)

For a2 = i2 + 1 To (c2 - 1)
k1 = Combo9.List(a2)
If (k1 = j1) Then
Combo9.RemoveItem (a2)

End If
Next a2
Next i2
End Sub

Private Sub Combo8_Click()
MSHFlexGrid2.ColWidth(0) = 2500
MSHFlexGrid2.ColWidth(1) = 2500
MSHFlexGrid2.ColWidth(2) = 2500
MSHFlexGrid2.ColWidth(3) = 2500

RS11.Open "select * from table4 WHERE TRAINING_TYPE='" + Combo7.Text + "' AND ", ADO, adOpenStatic, adLockOptimistic
MSHFlexGrid2.Row = 0
MSHFlexGrid2.Col = 0
MSHFlexGrid2.Text = "TRAINING_TYPE"
MSHFlexGrid2.Col = 0
MSHFlexGrid2.Text = "TRAINING_NAME"
MSHFlexGrid2.Col = 0
MSHFlexGrid2.Text = "TRAINING_DATE"
MSHFlexGrid2.Col = 0
MSHFlexGrid2.Text = "NUMBER_OF_DAYS"
I = 1
Do Until RS11.EOF
Set Text21.DataSource = RS11
Text21.DataField = "TRAINING_TYPE"
Set Text22.DataSource = RS11
Text22.DataField = "TRAINING_NAME"
Set Text23.DataSource = RS11
Text23.DataField = "TRAINING_DATE"
Set Text24.DataSource = RS11
Text24.DataField = "NUMBER_OF_DAYS"

MSHFlexGrid2.Row = I
MSHFlexGrid2.Col = 0
MSHFlexGrid2.Text = Text21
MSHFlexGrid2.Col = 1
MSHFlexGrid2.Text = Text22
MSHFlexGrid2.Col = 2
MSHFlexGrid2.Text = Text23

MSHFlexGrid2.Col = 3
MSHFlexGrid2.Text = Text24

MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
I = I + 1
'Set MSHFlexGrid1.DataSource = rs5
RS11.MoveNext
Loop

End Sub

Private Sub Command1_Click()
If flag = 0 Then
MonthView1.Visible = True
flag = 1
Else
MonthView1.Visible = False
flag = 0
End If
Text1.SetFocus
End Sub
Private Sub Command3_Click()
Dim str1 As String
If Combo3.Text = "RESIDENTIAL" Then
str1 = "INSERT INTO TABLE4(TRAINING_ID,TRAINING_TYPE,TRAINING_NAME,TRAINING_DATE,NUMBER_OF_DAYS,CITY,DESCRIPTION) VALUES ('" + Text10.Text + "','" + Combo3.Text + "','" + Combo2.Text + "','" + Text2.Text + "'," + Text1.Text + ",'" + Combo6.Text + "','" + Text11.Text + "')"
Else
str1 = "INSERT INTO TABLE4(TRAINING_ID,TRAINING_TYPE,TRAINING_NAME,TRAINING_DATE,NUMBER_OF_DAYS,CITY,DESCRIPTION) VALUES ('" + Text10.Text + "','" + Combo3.Text + "','" + Combo2.Text + "','" + Text2.Text + "'," + Text1.Text + ",'" + Combo6.Text + "','" + Text11.Text + "')"
End If
RS.Open str1, ADO, adOpenStatic, adLockOptimistic
MsgBox "Data Inserted Successfully"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text10 = ""
Combo2 = ""
Combo6 = ""
Combo3 = ""
Text11 = ""
Frame1.Visible = False
End Sub

Private Sub CTE_Click()
Form12.Show
Form4.Hide
End Sub

Private Sub EP_Click()
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
Combo1.Clear
Combo5.Clear
Dim str3 As String
str3 = "select DISTINCT TRAINING_NAME from table4 "
Set rs2 = New ADODB.Recordset
rs2.Open str3, ADO, adOpenStatic, adLockOptimistic
Do Until rs2.EOF
Combo1.AddItem rs2!TRAINING_NAME
rs2.MoveNext
Loop

End Sub
Private Sub EX_Click()
End
End Sub
Private Sub Form_Load()
flag = 0
flag1 = 0
Set ADO = New ADODB.Connection
Set RS = New ADODB.Recordset
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=d:\New folder (3)\Database1.mdb;"
ADO.Open
End Sub
Private Sub Form_Unload(Cancel As Integer)
s = Combo1.Text
'RS.Close
'ADO.Close
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text2 = MonthView1.Value
MonthView1.Visible = False
Text1.SetFocus

End Sub


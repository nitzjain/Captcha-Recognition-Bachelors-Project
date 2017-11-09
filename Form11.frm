VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17250
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   10350
   ScaleWidth      =   17250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Text            =   "Text12"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   8040
      TabIndex        =   18
      Text            =   "Text11"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Text            =   "Text10"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00400000&
      Caption         =   "Proceed"
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
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Text            =   "Text9"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Text            =   "Text8"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00400000&
      Caption         =   "Check3"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   15000
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00400000&
      Caption         =   "Check2"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   12600
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00400000&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   10440
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   12720
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   10800
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   16777215
      BackColorBkg    =   4194304
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "CUSTOMER"
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
      Left            =   10320
      TabIndex        =   3
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "DISTRIBUTOR"
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
      Left            =   12480
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "SALES STAFF"
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
      Left            =   14880
      TabIndex        =   1
      Top             =   840
      Width           =   2070
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Date
Dim s As Date
Dim I As Integer
Dim STR As String
Dim STR2 As String
Dim STR9 As Date
Dim RS8 As ADODB.Recordset
Dim ADO As ADODB.Connection
Dim rs5 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text6 = "Yes"
Else
Text6 = "No"
End If
Set rs2 = New ADODB.Recordset
rs2.Open "update table4 set CUSTOMER1='" + Text6 + "' WHERE TRAINING_NAME='" + STR + "'And TRAINING_TYPE = '" + Text9.Text + "'", ADO, adOpenStatic, adLockOptimistic
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text7 = "Yes"
Else
Text7 = "No"
End If
Set rs3 = New ADODB.Recordset
rs3.Open "update table4 set DISTRIBUTOR1='" + Text6 + "' WHERE TRAINING_NAME='" + STR + "'And TRAINING_TYPE = '" + Text9.Text + "' ", ADO, adOpenStatic, adLockOptimistic

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Text8 = "Yes"
Else
Text8 = "No"
End If
Set rs4 = New ADODB.Recordset
rs4.Open "update table4 set SALESSTAFF1='" + Text6 + "' WHERE TRAINING_NAME='" + STR + "'And TRAINING_TYPE = '" + Text9.Text + "'", ADO, adOpenStatic, adLockOptimistic

End Sub

Private Sub Command1_Click()
Form4.Show
Unload Me
End Sub

Private Sub Form_Load()

Set ADO = New ADODB.Connection
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=d:\New folder (3)\Database1.mdb;"
ADO.Open
Set rs5 = New ADODB.Recordset
rs5.Open "select * from table4 ", ADO, adOpenStatic, adLockOptimistic
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 0
MSHFlexGrid1.Text = "TRAINING_ID"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 1
MSHFlexGrid1.Text = "TRAINING_NAME"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 2
MSHFlexGrid1.Text = "TRAINING_TYPE"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 3
MSHFlexGrid1.Text = "TRAINING_DATE"
MSHFlexGrid1.Row = 0
MSHFlexGrid1.Col = 4
MSHFlexGrid1.Text = "CITY"
MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1500
I = 1
Do Until rs5.EOF
Set Text1.DataSource = rs5
Text1.DataField = "TRAINING_ID"
Set Text2.DataSource = rs5
Text2.DataField = "TRAINING_NAME"
Set Text3.DataSource = rs5
Text3.DataField = "TRAINING_TYPE"
Set Text4.DataSource = rs5
Text4.DataField = "TRAINING_DATE"
Set Text5.DataSource = rs5
Text5.DataField = "CITY"
Set Text10.DataSource = rs5
Text10.DataField = "CUSTOMER1"
Set Text11.DataSource = rs5
Text11.DataField = "DISTRIBUTOR1"
Set Text12.DataSource = rs5
Text12.DataField = "SALESSTAFF1"
If Text10 = "Yes" And Text11 = "Yes" And Text12 = "Yes" Then
GoTo L1
Else

d = Text4
Text4.Text = Format(Text4.Text, "dd-mm-yyyy")
s = Date + 30

If d < s And d > Date Then

MSHFlexGrid1.Row = I
MSHFlexGrid1.Col = 0
MSHFlexGrid1.Text = Text1
MSHFlexGrid1.Col = 1
MSHFlexGrid1.Text = Text2
MSHFlexGrid1.Col = 2
MSHFlexGrid1.Text = Text3
MSHFlexGrid1.Col = 3
MSHFlexGrid1.Text = Text4
MSHFlexGrid1.Col = 4
MSHFlexGrid1.Text = Text5

MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
I = I + 1
'Set MSHFlexGrid1.DataSource = rs5
End If
L1:
rs5.MoveNext
End If
Loop

MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
'rs5.Close


'MsgBox ("PLEASE SELECT THE TRAINING NAME FIRST")
End Sub

Private Sub MSHFlexGrid1_Click()
'Check1.Value = 0
'Check2.Value = 0
'Check3.Value = 0

STR = MSHFlexGrid1.Text
MSHFlexGrid1.Col = MSHFlexGrid1.Col + 1
Text9 = MSHFlexGrid1.Text
'MsgBox (Text9.Text)

'Text9 = Format(Text9.Text, "mm/dd/yyyy")

Set RS8 = New ADODB.Recordset
Dim STR2 As String
STR2 = "select *  from table4 where TRAINING_NAME='" + STR + "' And TRAINING_TYPE = '" + Text9.Text + "'"
RS8.Open STR2, ADO, adOpenStatic, adLockOptimistic
Set Text6.DataSource = RS8
Text6.DataField = "CUSTOMER1"
Set Text7.DataSource = RS8
Text7.DataField = "DISTRIBUTOR1"
Set Text8.DataSource = RS8
Text8.DataField = "SALESSTAFF1"

If Text6 = "Yes" Then
Check1.Value = 1
Else
Check1.Value = 0
End If

If Text7 = "Yes" Then
Check2.Value = 1
Else
Check2.Value = 0
End If

If Text8 = "Yes" Then
Check3.Value = 1
Else
Check3.Value = 0
End If
End Sub

VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14520
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   6720
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00400000&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00400000&
      Caption         =   "OK"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C0C0&
      Height          =   1845
      Left            =   6600
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C0C0&
      Height          =   2115
      Left            =   6600
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECT THE MDB FILE >>>>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   480
      TabIndex        =   9
      Top             =   5520
      Width           =   5085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECT THE DIRECTORY >>>>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   5475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECT THE DRIVE >>>>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "CONVERSION OF ACCESS DATA TO EXCEL "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10185
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fld As Field
Dim sname As String
Dim db As Database
Dim TB As TableDef
Dim e1 As Excel.Application
Dim k As String
Dim wb As Workbook
Dim ws As Worksheet

Dim BG As String
Dim EN As String
Dim STR As String
Dim str1 As String
Dim COMD As New ADODB.Command
Dim CONECT As New ADODB.Connection
Dim RECORD As New ADODB.Recordset
Private Sub Command1_Click()
Dim I As Integer
Dim str1 As String
Dim STR As String
Dim ltr As String
Dim STR2 As String
Dim P As Integer
Dim PTR As String
Dim PTR1 As String
str1 = File1.Path
STR = File1.Path + "\" + File1.FileName
STR = " provider=Microsoft.Jet.OLEDB.4.0;Data Source="
str1 = Text1.Text
STR2 = STR + str1
CONECT.Open (STR2)
COMD.ActiveConnection = CONECT

P = 1
I = 1
Set e1 = CreateObject("excel.application")
Set wb = e1.Workbooks.ADD
Set db = OpenDatabase(str1)
For Each TB In db.TableDefs
 If Left(TB.Name, 4) <> "MSys" And Left(TB.Name, 4) <> "USys" Then
  If P <= 3 Then
     Set ws = wb.Sheets(P)
  Else
     Sheets.ADD
     Sheets.Move after:=Sheets(Sheets.Count)
     Set ws = wb.ActiveSheet
  End If
  str1 = TB.Name
  STR = "select * from "
  STR2 = STR & str1
  COMD.CommandText = STR2
  Set RECORD = COMD.Execute
  If RECORD.EOF = True And RECORD.BOF = True Then
  MsgBox "SORRY DEAR "
  Else
  If IsNull(RECORD.Fields()) Then
     MsgBox "sorry"
       Else
         I = 1
         While Not RECORD.EOF
           For j = 0 To TB.Fields.Count - 1
              ytr = Null
              If IsNull(RECORD.Fields(j)) Then
              ltr = "null value"
              Else
              ltr = RECORD.Fields(j)
              End If
              ws.Cells(I, j + 1).Value = ltr
              Next j
              RECORD.MoveNext
              I = I + 1
         Wend
     End If
   End If
 P = P + 1
 ws.Name = TB.Name
 End If
 Next
 Unload Me
 e1.Visible = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
Dim STR As String
Dir1.Refresh
File1.Refresh
File1.Pattern = "*.MDB"
File1.Path = Dir1.Path
File1.Enabled = True
End Sub
Private Sub Drive1_Change()
Drive1.Refresh
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
Dim STR As String
Command1.Enabled = True
Dim str1 As String
str1 = File1.Path
STR = File1.Path + "\" + File1.FileName
Text1.Text = STR
End Sub
Private Sub Form_Load()

Text1.Enabled = False
File1.Enabled = True
Command1.Enabled = False
File1.Pattern = "*.MDB"
File1.Refresh

End Sub



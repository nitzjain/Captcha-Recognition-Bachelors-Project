VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "LOGIN"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   Picture         =   "Form1.frx":0492
   ScaleHeight     =   8565
   ScaleWidth      =   14460
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   2040
      TabIndex        =   35
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "CHANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   5040
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox Text13 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6000
         TabIndex        =   31
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "SAVE"
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
         Left            =   5280
         TabIndex        =   30
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   6000
         TabIndex        =   29
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   6000
         TabIndex        =   28
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00400000&
         Caption         =   "COMPANY NAME"
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
         Height          =   495
         Left            =   600
         TabIndex        =   34
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label13 
         BackColor       =   &H00400000&
         Caption         =   "NEW ANSWER"
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
         Height          =   495
         Left            =   600
         TabIndex        =   33
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00400000&
         Caption         =   "CONFIRM ANSWER"
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
         Height          =   495
         Left            =   600
         TabIndex        =   32
         Top             =   2880
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "CHANGE"
      Height          =   5655
      Left            =   5160
      TabIndex        =   20
      Top             =   3480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   6960
         TabIndex        =   10
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   6960
         TabIndex        =   9
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
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
         TabIndex        =   11
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00400000&
         Caption         =   "CONFIRM PASSWORD"
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
         Height          =   495
         Left            =   720
         TabIndex        =   23
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00400000&
         Caption         =   "NEW PASSWORD"
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
         Height          =   495
         Left            =   720
         TabIndex        =   22
         Top             =   1740
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00400000&
         Caption         =   "OLD PASSWORD"
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
         Height          =   495
         Left            =   720
         TabIndex        =   21
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
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
      Left            =   15840
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "FORGOT PASSWORD"
      Height          =   5655
      Left            =   5160
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   6120
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000002&
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
         Left            =   6480
         TabIndex        =   13
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00400000&
         Caption         =   "CHANGE ANSWER?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9240
         MouseIcon       =   "Form1.frx":16E320
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00400000&
         Caption         =   "ENTER YOUR MIDDLE NAME?"
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
         Height          =   1095
         Left            =   600
         TabIndex        =   25
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00400000&
         Caption         =   "ENTER YOUR USER_ID"
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
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   1320
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
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
      Left            =   13920
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   11400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4900
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Form1.frx":16E472
      Top             =   4320
      Width           =   11415
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   725
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "ABOUT SKF INDIA"
      Top             =   3600
      Width           =   11415
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "www.skfindia.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8880
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   10680
      Width           =   2520
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "CHANGE PASSWORD?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11400
      MouseIcon       =   "Form1.frx":16E9E1
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "FORGOT PASSWORD?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "PASSWORD"
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
      Height          =   495
      Left            =   9120
      TabIndex        =   16
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "USER_ID"
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
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pass As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal ipoperation As String, ByVal ipfile As String, ByVal ipparameters As String, ByVal ipdiectory As String, ByVal nshowcmd As Long) As Long
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Command1_Click()
If Text1 = "admin" And Text2 = pass Then
Form11.Show
Unload Me
Else
MsgBox ("INVALID PASSWORD/ID")
Text1 = ""
Text2 = ""
Text1.SetFocus
End If
End Sub
Private Sub Command2_Click()
If Text3 = "admin" And Text10 = Text14 Then
MsgBox ("Your password is " & pass)
Text3 = ""
Text10 = ""
Frame1.Visible = False
Else
MsgBox ("You Entered Wrong ID/answer")
Text3 = ""
Text10 = ""
End If
End Sub
Private Sub Command3_Click()
If Text4 = pass Then
If Text6 = Text5 Then
Dim str1 As String
str1 = "update table1 set pass ='" + Text5.Text + "'"
Set RS = ADO.Execute(str1)
MsgBox ("Password Changed")
Frame2.Visible = False
Text4 = ""
Text5 = ""
Text6 = ""
Else
MsgBox ("Enter same password")
Text4 = ""
Text5 = ""
Text6 = ""
End If
Else
MsgBox ("wrong old password")
Text4 = ""
Text5 = ""
Text6 = ""
End If
Unload Me
Form1.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
If Text13 = "skf" Then
If Text11 = Text12 Then
Dim str1 As String
str1 = "update table1 set ANS ='" + Text11.Text + "'"
Set RS = ADO.Execute(str1)
MsgBox ("Answer Changed")
Frame3.Visible = False
Text11 = ""
Text12 = ""
Text13 = ""
Else
MsgBox ("Enter same answer")
Text11 = ""
Text12 = ""
Text13 = ""
End If
Else
MsgBox ("wrong company name")
Text11 = ""
Text12 = ""
Text13 = ""
End If
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
Set ADO = New ADODB.Connection
Set RS = New ADODB.Recordset
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=d:\New folder (3)\Database1.mdb;"
ADO.Open
RS.Open "select * from table1", ADO, adOpenStatic, adLockOptimistic
Set Text7.DataSource = RS
Text7.DataField = "pass"
pass = Text7.Text
Set Text14.DataSource = RS
Text14.DataField = "ANS"

End Sub

Private Sub Form_Terminate()
RS.Close
ADO.Close
End Sub

Private Sub Label10_Click()
ShellExecute Form1.hwnd, "open", "http:\\www.skfindia.com", "", "", sw_show
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.MouseIcon = LoadPicture("d:\New folder (3)\Hand.ico")

End Sub


Private Sub Label11_Click()
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False

End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.MouseIcon = LoadPicture("d:\New folder (3)\Hand.ico")

End Sub

Private Sub Label3_Click()
'Text3.SetFocus
Frame1.Visible = True
Frame2.Visible = False
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.MouseIcon = LoadPicture("d:\New folder (3)\Hand.ico")
End Sub
Private Sub Label4_Click()
'Text4.SetFocus

Frame2.Visible = True
Frame1.Visible = False
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.MouseIcon = LoadPicture("d:\New folder (3)\Hand.ico")
End Sub



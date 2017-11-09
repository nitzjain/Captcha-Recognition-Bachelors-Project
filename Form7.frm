VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17235
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   17235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   12000
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   21000
      Begin VB.CommandButton Command1 
         Height          =   1455
         Left            =   6720
         Picture         =   "Form7.frx":A4382
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7440
         Width           =   1815
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
         ForeColor       =   &H80000005&
         Height          =   725
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "ABOUT SKF INDIA"
         Top             =   120
         Width           =   11415
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
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "Form7.frx":A7E18
         Top             =   840
         Width           =   11415
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub Command1_Click()
Form4.Show
Form7.Hide
End Sub

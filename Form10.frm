VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16485
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4440
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   5040
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   8880
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
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
      Left            =   8760
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
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
      TabIndex        =   2
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
      Left            =   11160
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
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
      Left            =   11160
      TabIndex        =   0
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label11 
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
      Left            =   360
      TabIndex        =   23
      Top             =   7920
      Width           =   1020
   End
   Begin VB.Label Label10 
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
      Left            =   240
      TabIndex        =   22
      Top             =   3240
      Width           =   690
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
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   2040
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
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   2520
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
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   2595
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
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   3300
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
      Left            =   240
      TabIndex        =   17
      Top             =   5160
      Width           =   2175
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
      Left            =   240
      TabIndex        =   16
      Top             =   6000
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
      Height          =   585
      Left            =   240
      TabIndex        =   15
      Top             =   6960
      Width           =   930
   End
   Begin VB.Label Label8 
      Caption         =   "STATE"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   6960
      Width           =   855
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
      Left            =   240
      TabIndex        =   13
      Top             =   9000
      Width           =   3060
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

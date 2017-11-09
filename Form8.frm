VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17460
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   10020
   ScaleWidth      =   17460
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   4194304
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   4194304
      BackColorBkg    =   4194304
      WordWrap        =   -1  'True
      GridLineWidth   =   2
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLineWidthBand=   2
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim ADO As ADODB.Connection
Dim RS As ADODB.Recordset
Set ADO = New ADODB.Connection
Set RS = New ADODB.Recordset
ADO.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Pranay\Desktop\New folder (3)\Database1.mdb;"
ADO.Open
RS.Open "select * from table3", ADO, adOpenStatic, adLockOptimistic
Set MSHFlexGrid1.DataSource = RS
MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 4500
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 1500
End Sub


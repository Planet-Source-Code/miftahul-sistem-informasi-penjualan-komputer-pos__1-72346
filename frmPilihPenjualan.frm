VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPilihPenjualan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pillih Penjualan yang akan di Angsur"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8655
   Begin Project1.DMSXpButton cmdBatal 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Batal"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdPilihPenjualan 
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pilih"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPilihPenjualan.frx":0000
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frmPilihPenjualan.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmPilihPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdPilihPenjualan_Click()
    frmAngsuran.txtAngsuran(0).Text = Data1.Recordset(0)
    Unload Me
End Sub

Private Sub Form_Load()
    Set rc = DB.OpenRecordset("tblPenjualan", dbOpenDynaset)
    Set Data1.Recordset = rc
    Set rc = Nothing
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    Me.Width = 8745
    Me.Height = 3885
End Sub


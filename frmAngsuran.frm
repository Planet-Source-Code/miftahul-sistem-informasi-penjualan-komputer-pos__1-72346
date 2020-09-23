VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmAngsuran 
   Caption         =   "Angsuran"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   Icon            =   "frmAngsuran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10755
   Begin VB.Data dtaPenjualan 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Width           =   1455
   End
   Begin Project1.DMSXpButton cmdHapus 
      Height          =   375
      Left            =   4320
      TabIndex        =   37
      Top             =   6240
      Width           =   615
      _ExtentX        =   1085
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
      Caption         =   "Hapus"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   720
   End
   Begin Project1.DMSXpButton cmdProses 
      Height          =   255
      Left            =   3360
      TabIndex        =   35
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Proses"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmAngsuran.frx":030A
      Height          =   3615
      Left            =   240
      OleObjectBlob   =   "frmAngsuran.frx":031E
      TabIndex        =   26
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtAngsuran 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
   End
   Begin Project1.DMSXpButton cmdPilihNoPenjualan 
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "..."
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.TextBox txtAngsuran 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Detail Penjualan ]"
      Enabled         =   0   'False
      Height          =   5655
      Left            =   5040
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   10
         Left            =   2160
         TabIndex        =   29
         Top             =   4920
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   2160
         TabIndex        =   27
         Top             =   4560
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   22
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   20
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   16
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   14
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   12
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   10
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Barang g dibeli :"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Dengan rincian pembayaran :"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label lblDetail 
         Caption         =   "Sisa yg belum terbayar"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   30
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label lblDetail 
         Caption         =   "Terbayar"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   28
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Dengan Harga"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   21
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Optical Drive"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   19
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "VGA"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   17
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Memory"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Storage"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Processor"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Nama Barang"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "No KTP / SIM"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblDetail 
         Caption         =   "Nama Konsumen"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtAngsuran 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Untuk menghapus pembayaran terakhir tekan tombol ini"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   6360
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label4 
      Caption         =   "tekan tombol ini untuk memilih penjualan Lain yg ingin dibayar"
      Height          =   255
      Left            =   4320
      TabIndex        =   36
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Total Angsuran sampai saat ini :"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblAngsuran 
      Caption         =   "Angsuran Rp."
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblAngsuran 
      Caption         =   "Tanggal"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblAngsuran 
      Caption         =   "No. Penjualan"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmAngsuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdHapus_Click()
 With Data1
    If .Recordset.EOF And .Recordset.BOF Then Exit Sub
    .Recordset.MoveLast
    If MsgBox("Yakin akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
        UpdateJual False, .Recordset(2)
        .Recordset.Delete
    End If
    .Refresh
    Hitung
    Cek
 End With
End Sub

Private Sub cmdPilihNoPenjualan_Click()
    Unload Me
    frmPilihPenjualan.Show
End Sub

Private Sub cmdProses_Click()
 Dim StatOk As Boolean
  With Data1
    StatOk = True
    For x = 0 To txtAngsuran.Count - 1
        If txtAngsuran(x).Text = "" Then
            MsgBox "Maaf " & lblAngsuran(x).Caption & " tidak boleh kosong", vbCritical, "Peringatan"
            StatOk = False
            txtAngsuran(x).SetFocus
            Exit For
        End If
    Next
    If StatOk = False Then Exit Sub
    Txt2Data
    UpdateJual True, txtAngsuran(2).Text
    Hitung
    Cek
  End With
End Sub

Private Sub Data1_Reposition()
    AdjustMasterGrid
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    Me.Width = 10875
    Me.Height = 7590
End Sub

Private Sub Form_Activate()
  Dim x As Byte
  PaintForm3D Me
      
    txtAngsuran(1).Text = Format(Now, "dd-mm-yyyy")
    Timer1.Enabled = True
    ForceDetail
    Hitung
    Cek
End Sub

Private Sub Form_Load()
    Set rc = DB.OpenRecordset("tblAngsuran", dbOpenDynaset)
    Set Data1.Recordset = rc
    Set rc = Nothing
    
    Set rc = DB.OpenRecordset("tblPenjualan", dbOpenDynaset)
    Set dtaPenjualan.Recordset = rc
    Set rc = Nothing

End Sub

Private Sub Data2DetailJual()
 With Data2
    txtDetail(0).Text = .Recordset(0)
    txtDetail(1).Text = .Recordset(1)
    txtDetail(2).Text = .Recordset(2)
    txtDetail(3).Text = .Recordset(3)
    txtDetail(4).Text = .Recordset(4)
    txtDetail(5).Text = .Recordset(5)
    txtDetail(6).Text = .Recordset(6)
    txtDetail(7).Text = .Recordset(7)
    txtDetail(8).Text = ViewFormat(.Recordset(8))
    txtDetail(9).Text = ViewFormat(.Recordset(9))
    txtDetail(10).Text = ViewFormat(.Recordset(10))
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub


Private Sub Timer1_Timer()
    txtAngsuran(2).Text = Format(txtAngsuran(2).Text, "###,###,###")
    txtAngsuran(2).SelStart = Len(txtAngsuran(2).Text)
End Sub

Private Sub txtAngsuran_GotFocus(Index As Integer)
    txtAngsuran(Index).BackColor = vbGreen
End Sub

Private Sub txtAngsuran_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 38 Then
    If Index = 0 Then Exit Sub
    txtAngsuran(Index - 1).SetFocus
 ElseIf KeyCode = 40 Or KeyCode = 13 Then
    If Index = txtAngsuran.Count - 1 Then Exit Sub
    txtAngsuran(Index + 1).SetFocus
 End If
End Sub

Private Sub txtAngsuran_KeyPress(Index As Integer, KeyAscii As Integer)
 If Index = 2 Then
    If KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 8 Then
        KeyAscii = 0
        Beep
    End If
 End If
End Sub

Private Sub txtAngsuran_LostFocus(Index As Integer)
    txtAngsuran(Index).BackColor = vbWhite
End Sub

Private Sub Txt2Data()
    With Data1
        .Recordset.AddNew
        .Recordset(0) = txtAngsuran(0).Text
        .Recordset(1) = txtAngsuran(1).Text
        .Recordset(2) = DefaFormat(txtAngsuran(2).Text)
        .Recordset.Update
    End With
End Sub

Private Sub Hitung()
Dim TotAng As String
 lblTotal.Caption = "0"
 ForceDetail
 With Data1
    If .Recordset.EOF And .Recordset.BOF Then Exit Sub
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
        TotAng = Val(TotAng) + Val(.Recordset(2))
        .Recordset.MoveNext
    Loop
    lblTotal.Caption = ViewFormat(TotAng)
 End With
End Sub

Private Sub AdjustMasterGrid()
  With DBGrid1
    .Columns(0).Width = 800
    .Columns(1).Width = 1500
    .Columns(2).Width = 1700
  End With
End Sub

Private Sub ForceDetail()
Dim SQL As String
    SQL = "SELECT tblKonsumen.Nama,tblKonsumen.ID,"
    SQL = SQL & "tblBarang.NamaBarang,tblBarang.Processor,tblBarang.Storage,tblBarang.Memory,tblBarang.VGA,tblBarang.OpticalDrive,"
    SQL = SQL & "tblPenjualan.Harga,tblPenjualan.Terbayar,tblPenjualan.Sisa "
    SQL = SQL & "FROM tblKonsumen,tblBarang,tblPenjualan "
    SQL = SQL & "WHERE tblPenjualan.KodeKonsumen=tblKonsumen.KodeKonsumen "
    SQL = SQL & "AND tblPenjualan.KodeBarang=tblBarang.KodeBarang "
    SQL = SQL & "AND tblPenjualan.NoPenjualan='" & txtAngsuran(0).Text & "'"
    
    Set rc = DB.OpenRecordset("tblAngsuran", dbOpenDynaset)
    Set Data2.Recordset = rc
    Data2.RecordSource = SQL
    Set rc = Nothing
    DoEvents
    Data2.Refresh
    Data2DetailJual
    
    Data1.RecordSource = "SELECT * FROM tblAngsuran WHERE NoPenjualan='" & txtAngsuran(0).Text & "'"
    Data1.Refresh
        
End Sub

Private Sub UpdateJual(ByVal DiBayar As Boolean, Nilai As String)
 Dim Harga As Double, Terbayar As Double, Sisa As Double
 With dtaPenjualan
    .RecordSource = "SELECT * FROM tblPenjualan WHERE NoPenjualan='" & txtAngsuran(0).Text & "'"
    .Refresh
    If DiBayar = True Then
       Harga = .Recordset(4)
       Terbayar = .Recordset(5)
       Terbayar = Terbayar + Val(DefaFormat(Nilai))
       Sisa = Harga - Terbayar
     Else
       Harga = .Recordset(4)
       Terbayar = .Recordset(5)
       Terbayar = Terbayar - Val(DefaFormat(Nilai))
       Sisa = Harga - Terbayar
    End If
    .Recordset.Edit
    .Recordset(5) = Terbayar
    .Recordset(6) = Sisa
    .Recordset.Update
    .Refresh
 End With
End Sub

Private Sub Cek()
    If DefaFormat(lblTotal.Caption) = DefaFormat(txtDetail(8).Text) And lblTotal.Caption <> "0" Then
        cmdProses.Enabled = False
      Else
        cmdProses.Enabled = True
    End If
End Sub

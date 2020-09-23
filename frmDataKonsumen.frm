VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDataKonsumen 
   Caption         =   "Data Konsumen"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   Icon            =   "frmDataKonsumen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11685
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   9855
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   16
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   15
         Top             =   1680
         Width           =   7695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   14
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   13
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "No. HP"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No. Telepon Rumah"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tempat Tanggal Lahir"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "No. KTP/SIM"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Konsumen"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Width           =   2580
   End
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   0
      Top             =   6840
      Width           =   2415
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   2
      Left            =   10200
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
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
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   1
      Left            =   10200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Koreksi"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   0
      Left            =   10200
      TabIndex        =   3
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Baru"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDataKonsumen.frx":08CA
      Height          =   3615
      Left            =   240
      OleObjectBlob   =   "frmDataKonsumen.frx":08DE
      TabIndex        =   4
      Top             =   3120
      Width           =   11175
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   5
      Left            =   10200
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tutup"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   3
      Left            =   10200
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Simpan"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   4
      Left            =   10200
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
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
   Begin VB.Label lblHasilCari 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Cari Data"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   6840
      Width           =   855
   End
End
Attribute VB_Name = "frmDataKonsumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NoDisplay As Boolean

Private Sub cmdData_Click(Index As Integer)
Dim x As Byte, StatOK As Boolean
On Error GoTo LocErr
 With Data1
    If Index = 0 Then
        ModeEdit True
        .Recordset.AddNew
        BlankTxt
        Text1(0).SetFocus
        Text1(5).Text = "None"
        Text1(6).Text = "None"
    ElseIf Index = 1 Then
        If .Recordset.EOF And .Recordset.BOF Then
            MsgBox "Tidak ada data yang akan dikoreksi, silahkan pilih data yang bersangkutan" & vbNewLine & _
                   "setelah itu ulangi lagi proses yang ingin anda lakukan.", vbCritical, "Peringatan"
            Exit Sub
        End If
        ModeEdit True
        .Recordset.Edit
        Text1(1).SetFocus
    ElseIf Index = 2 Then
        If MsgBox("Yakin akan dihapus?", vbYesNo + vbQuestion, "Hapus Data") = vbYes Then
            .Recordset.Delete
            .Recordset.MovePrevious
            If .Recordset.BOF Then
                .Recordset.MoveNext
                If .Recordset.EOF Then
                    BlankTxt
                    .Refresh
                End If
            End If
        End If
    ElseIf Index = 3 Then
        StatOK = True
        For x = 0 To Text1.Count - 1
            If Text1(x).Text = "" Then
                MsgBox "Maaf " & Label1(x).Caption & " tidak boleh kosong", vbCritical, "Peringatan"
                StatOK = False
                Text1(x).SetFocus
                Exit For
            End If
        Next
        If Len(Text1(0).Text) < 5 Then
            MsgBox "Maaf, Kode Konsumen harus 5 huruf", vbCritical, "Peringatan"
            Text1(0).SetFocus
            StatOK = False
        End If
        If StatOK = False Then Exit Sub
        Txt2Data
        .Recordset.Update
        ModeEdit False
    ElseIf Index = 4 Then
        BlankTxt
        .Recordset.CancelUpdate
        .Refresh
        ModeEdit False
    ElseIf Index = 5 Then
        Me.Hide
    End If
    AdjustMasterGrid
 End With
 Exit Sub
LocErr:
  If Err.Number = 3022 Then
    MsgBox "Data Konsumen dengan Kode [ " & Text1(0).Text & " ] sudah ada." & vbNewLine & _
           "Harap diubah dengan kode lain lalu Klik Tombol Simpan!"
  ElseIf Err.Number = 3200 Then
    MsgBox "Data ini tidak dapat dihapus karena masih digunakan pada Data Penjualan" & vbNewLine & _
           "Agar dapat dihapus, Hapus atau Koreksi terlebih dahulu Data Penjualan yang masih berhubungan" & vbNewLine & _
           "dengan Konsumen ini ", vbCritical, "Peringatan"
  ElseIf Err.Number = 3021 Then
    MsgBox "Tidak ada data yang akan diproses, silahkan pilih data yang bersangkutan" & vbNewLine & _
           "setelah itu ulangi lagi proses yang ingin anda lakukan.", vbCritical, "Peringatan"
  Else
    MsgBox Err.Description, vbCritical, Err.Number
  End If
End Sub

Private Sub Data1_Reposition()
   AdjustMasterGrid
   Data1.Caption = "Posisi " & Data1.Recordset.AbsolutePosition + 1 & " Dari " & Data1.Recordset.RecordCount
   If NoDisplay = False Then Data2Txt
End Sub

Private Sub Form_Activate()
  Dim x As Byte
  PaintForm3D Me
  
  ModeEdit False
End Sub

Private Sub Form_Load()
 Set rc = DB.OpenRecordset("tblKonsumen", dbOpenDynaset)
 Set Data1.Recordset = rc

 Set rc = Nothing
 Data1.Refresh
 AdjustMasterGrid
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    Me.Width = 11715
    Me.Height = 7905
End Sub

Private Sub ModeEdit(ByVal mVal As Boolean)
 Dim a As Byte
    NoDisplay = mVal
    Frame1.Enabled = mVal
    
    DBGrid1.Enabled = Not mVal
    Data1.Enabled = Not mVal
    
    cmdData(0).Enabled = Not mVal
    cmdData(1).Enabled = Not mVal
    cmdData(2).Enabled = Not mVal
    cmdData(3).Enabled = mVal
    cmdData(4).Enabled = mVal
    cmdData(5).Enabled = Not mVal
    txtCari.Visible = Not mVal
    Label2.Visible = Not mVal
    DisableClose Me, mVal
    AdjustMasterGrid
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    SelText Index
    Text1(Index).BackColor = vbGreen
End Sub

Private Sub SelText(ByVal Index As Byte)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 38 Then
    If Index = 0 Then Exit Sub
    Text1(Index - 1).SetFocus
 ElseIf KeyCode = 40 Or KeyCode = 13 Then
    If Index = Text1.Count - 1 Then Exit Sub
    Text1(Index + 1).SetFocus
 End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
End Sub

Private Sub txtCari_Change()
Dim SQL As String
SQL = "KodeKonsumen like '*" & txtCari.Text & "*' OR "
SQL = SQL & "ID like '*" & txtCari.Text & "*' OR "
SQL = SQL & "Nama like '*" & txtCari.Text & "*' OR "
SQL = SQL & "TTL like '*" & txtCari.Text & "*' OR "
SQL = SQL & "Alamat like '*" & txtCari.Text & "*' OR "
SQL = SQL & "NoTelepon like '*" & txtCari.Text & "*' OR "
SQL = SQL & "NoHP like '*" & txtCari.Text & "*'"
 With Data1
       .Recordset.FindFirst SQL
       If .Recordset.NoMatch Then
           .Recordset.FindNext SQL
           lblHasilCari.ForeColor = vbRed
           lblHasilCari.Caption = "Data tidak ditemukan"
         Else
           lblHasilCari.ForeColor = vbBlue
           lblHasilCari.Caption = "Data ditemukan"
       End If
 End With
End Sub

Private Sub AdjustMasterGrid()
  With DBGrid1
    .Columns(0).Width = 800
    .Columns(1).Width = 1700
    .Columns(2).Width = 3000
    .Columns(3).Width = 1600
    .Columns(4).Width = 1600
  End With
End Sub

Private Sub Data2Txt()
Dim x As Byte
 With Data1
  If .Recordset.BOF And .Recordset.BOF Then Exit Sub
  For x = 0 To Text1.Count - 1
    If x <= 6 Then
        Text1(x).Text = .Recordset(x)
      ElseIf x > 6 Then
        Text1(x).Text = ViewFormat(.Recordset(x))
    End If
  Next
 End With
End Sub

Private Sub Txt2Data()
Dim x As Byte
  For x = 0 To Text1.Count - 1
    If x <= 6 Then
        Data1.Recordset(x) = Text1(x).Text
      ElseIf x > 6 Then
        Data1.Recordset(x) = DefaFormat(Text1(x).Text)
    End If
  Next
End Sub

Private Sub BlankTxt()
Dim x As Byte
  For x = 0 To Text1.Count - 1
    Text1(x).Text = ""
  Next
End Sub

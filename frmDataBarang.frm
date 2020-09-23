VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDataBarang 
   Caption         =   "Data Barang"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   Icon            =   "frmDataBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11595
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9360
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   8415
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   18
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   17
         Top             =   960
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1320
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2040
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2400
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Barang"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Barang"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Type Processor"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Storage"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Memory"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Display Adapter"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Optical Drive"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Harga Jual"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Harga Belli"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   6960
      Width           =   2415
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
      Top             =   6960
      Width           =   2700
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   2
      Left            =   10200
      TabIndex        =   2
      Top             =   1320
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
      TabIndex        =   3
      Top             =   840
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
      TabIndex        =   4
      Top             =   360
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
      Bindings        =   "frmDataBarang.frx":08CA
      Height          =   3015
      Left            =   240
      OleObjectBlob   =   "frmDataBarang.frx":08DE
      TabIndex        =   5
      Top             =   3840
      Width           =   11175
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   5
      Left            =   10200
      TabIndex        =   6
      Top             =   6960
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
      TabIndex        =   7
      Top             =   1800
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
      TabIndex        =   8
      Top             =   2280
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
   Begin VB.Label Label2 
      Caption         =   "Cari Data"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   6960
      Width           =   1335
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
      Left            =   6960
      TabIndex        =   9
      Top             =   6960
      Width           =   3015
   End
End
Attribute VB_Name = "frmDataBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NoDisplay As Boolean

Private Sub cmdData_Click(Index As Integer)
Dim x As Byte, StatOk As Boolean, No As String
On Error GoTo LocErr
 With Data1
    If Index = 0 Then
        No = AutoNumber
        ModeEdit True
        .Recordset.AddNew
        BlankTxt
        Text1(0).Text = No
        Text1(1).SetFocus
    ElseIf Index = 1 Then
        If .Recordset.EOF And .Recordset.BOF Then
            MsgBox "Tidak ada data yang akan dikoreksi, silahkan pilih data yang bersangkutan" & vbNewLine & _
                   "setelah itu ulangi lagi proses yang ingin anda lakukan.", vbCritical, "Peringatan"
            Exit Sub
        End If
        ModeEdit True
        .Recordset.Edit
        Text1(0).SetFocus
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
        StatOk = True
        For x = 0 To Text1.Count - 1
            If Text1(x).Text = "" Then
                MsgBox "Maaf " & Label1(x).Caption & " tidak boleh kosong", vbCritical, "Peringatan"
                StatOk = False
                Text1(x).SetFocus
                Exit For
            End If
        Next
        If StatOk = False Then Exit Sub
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
    MsgBox "Data Barang dengan Kode [ " & Text1(0).Text & " ] sudah ada"
    ModeEdit False
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
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
 Set rc = DB.OpenRecordset("tblBarang", dbOpenDynaset)
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

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
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

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
 If Index = 7 Or Index = 8 Then
    If KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 8 Then
        KeyAscii = 0
        Beep
    End If
 End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
    Text1(7).Text = Format(Text1(7).Text, "###,###,###")
    Text1(7).SelStart = Len(Text1(7).Text)
    Text1(8).Text = Format(Text1(8).Text, "###,###,###")
    Text1(8).SelStart = Len(Text1(8).Text)
End Sub

Private Sub txtCari_Change()
Dim SQL As String
SQL = "KodeBarang like '*" & txtCari.Text & "*' OR "
SQL = SQL & "NamaBarang like '*" & txtCari.Text & "*' OR "
SQL = SQL & "Processor like '*" & txtCari.Text & "*' OR "
SQL = SQL & "Storage like '*" & txtCari.Text & "*' OR "
SQL = SQL & "VGA like '*" & txtCari.Text & "*' OR "
SQL = SQL & "OpticalDrive like '*" & txtCari.Text & "*'"
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

Private Function AutoNumber() As String
  Dim No As String
  With Data1
    .RecordSource = "SELECT * FROM tblBarang ORDER BY RIGHT(KodeBarang,4) ASC"
    .Refresh
    If .Recordset.EOF And .Recordset.BOF Then
        AutoNumber = "B0001"
        Exit Function
    End If
    .Recordset.MoveLast
    No = Val(Right(.Recordset(0), 4)) + 1
    AutoNumber = "B" & String(4 - Len(No), "0") & No
  End With
End Function

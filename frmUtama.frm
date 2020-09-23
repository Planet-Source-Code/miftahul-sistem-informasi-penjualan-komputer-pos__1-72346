VERSION 5.00
Begin VB.MDIForm frmUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Sistem Informasi Penjualan Komputer Versi 1.0.0 (Beta)"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7965
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmUtama.frx":1CFA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileDataBarang 
         Caption         =   "Data &Barang"
      End
      Begin VB.Menu mnuFileDataKonsumen 
         Caption         =   "Data &Konsumen"
      End
      Begin VB.Menu mnuFileGaris1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileKeluar 
         Caption         =   "Kel&uar"
      End
   End
   Begin VB.Menu mnuTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnuTransaksiPenjualan 
         Caption         =   "Penjualan"
      End
      Begin VB.Menu mnuTransaksiPembayaran 
         Caption         =   "Pembayaran Angsuran"
      End
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "Laporan"
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFileDataBarang_Click()
    frmDataBarang.Show
End Sub

Private Sub mnuFileDataKonsumen_Click()
    frmDataKonsumen.Show
End Sub

Private Sub mnuFileKeluar_Click()
    End
End Sub

Private Sub mnuTransaksiPembayaran_Click()
    frmPilihPenjualan.Show
End Sub

Private Sub mnuTransaksiPenjualan_Click()
    frmPenjualan.Show
End Sub

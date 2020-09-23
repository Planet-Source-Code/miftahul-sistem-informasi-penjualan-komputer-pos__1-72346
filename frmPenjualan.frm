VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPenjualan 
   Caption         =   "Penjualan"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmPenjualan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11475
   Begin VB.PictureBox picBarang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   3480
      ScaleHeight     =   4665
      ScaleWidth      =   6585
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtCariBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   69
         Top             =   4200
         Width           =   1935
      End
      Begin Project1.DMSXpButton cmdPilihBarang 
         Height          =   375
         Left            =   5400
         TabIndex        =   68
         Top             =   4200
         Width           =   975
         _ExtentX        =   1720
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
      Begin Project1.DMSXpButton cmdNavBarang 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   4200
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   "<"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.TextBox txtBarang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   55
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   54
         Top             =   3000
         Width           =   6135
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   53
         Top             =   2640
         Width           =   6135
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   52
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         Top             =   1920
         Width           =   6135
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   675
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   675
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtBarang 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   48
         Top             =   120
         Width           =   1215
      End
      Begin Project1.DMSXpButton cmdNavBarang 
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   65
         Top             =   4200
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   "<|"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.DMSXpButton cmdNavBarang 
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   66
         Top             =   4200
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   "|>"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.DMSXpButton cmdNavBarang 
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   67
         Top             =   4200
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   ">"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.Label lblPosisiBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   63
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Optical Drive"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   62
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Display Adapter"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   61
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Memory"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   60
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Storage"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   59
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Type Processor"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblBarang 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox picKonsumen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   3480
      ScaleHeight     =   4065
      ScaleWidth      =   5625
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   45
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   44
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   43
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   42
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   795
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   41
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   40
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtKonsumen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   39
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtCariKonsumen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         TabIndex        =   38
         Top             =   3600
         Width           =   2055
      End
      Begin Project1.DMSXpButton cmdNavKonsumen 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   "<"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.DMSXpButton cmdPilihKonsumen 
         Height          =   375
         Left            =   4680
         TabIndex        =   33
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
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
      Begin Project1.DMSXpButton cmdNavKonsumen 
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   35
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "<|"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.DMSXpButton cmdNavKonsumen 
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   36
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "|>"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin Project1.DMSXpButton cmdNavKonsumen 
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   37
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   ">"
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.Label lblPosisiKonsumen 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Konsumen"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "No. KTP/SIM"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Tanggal Lahir"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telepon Rumah"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblKonsumen 
         BackStyle       =   0  'Transparent
         Caption         =   "No. HP"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   8775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin Project1.DMSXpButton cmdRelasi 
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin Project1.DMSXpButton cmdRelasi 
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin VB.Label Label1 
         Caption         =   "No. Penjualan"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Konsumen"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Barang"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Harga"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Terbayar"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sisa"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblRelasi 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   17
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label lblRelasi 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   16
         Top             =   1320
         Width           =   4815
      End
   End
   Begin VB.Data dtaKonsumen 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaBarang 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPenjualan.frx":08CA
      Height          =   3375
      Left            =   240
      OleObjectBlob   =   "frmPenjualan.frx":08DE
      TabIndex        =   5
      Top             =   3120
      Width           =   11055
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   2
      Left            =   9960
      TabIndex        =   0
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
      Caption         =   "Simpan"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   1
      Left            =   9960
      TabIndex        =   1
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
      Caption         =   "Hapus"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   0
      Left            =   9960
      TabIndex        =   2
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
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   3
      Left            =   9960
      TabIndex        =   3
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
      Caption         =   "Batal "
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Project1.DMSXpButton cmdData 
      Height          =   405
      Index           =   4
      Left            =   9960
      TabIndex        =   4
      Top             =   6720
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
End
Attribute VB_Name = "frmPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rcBarang As Recordset
Dim rcKonsumen As Recordset
Dim StatKonsumen As Boolean
Dim StatBarang As Boolean
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
        Text1(1).Text = Format(Now, "dd-mm-yyyy")
        Text1(1).SetFocus
    ElseIf Index = 1 Then
        If MsgBox("Yakin akan dihapus?", vbYesNo + vbQuestion, "Hapus Data") = vbYes Then
            .Recordset.Delete
            .Recordset.MovePrevious
            If .Recordset.BOF Then
                .Recordset.MoveNext
                If .Recordset.EOF Then .Refresh
            End If
        End If
    ElseIf Index = 2 Then
        StatOk = True
        For x = 0 To Text1.Count - 1
            If Text1(x).Text = "" Then
                MsgBox "Maaf " & Label1(x).Caption & " tidak boleh kosong", vbCritical, "Peringatan"
                StatOk = False
                Text1(x).SetFocus
                Exit For
            End If
        Next
        If CekTgl(Text1(1).Text) = False Then
            Text1(1).SetFocus
            StatOk = False
        End If
        If StatOk = False Then Exit Sub
        Txt2Data
        .Recordset.Update
        ModeEdit False
    ElseIf Index = 3 Then
        .Recordset.CancelUpdate
        .Refresh
        ModeEdit False
    ElseIf Index = 4 Then
        Me.Hide
    End If
    AdjustMasterGrid
 End With
 Exit Sub
LocErr:
  If Err.Number = 3022 Then
    MsgBox "Data dengan Kode Konsumen tersebut sudah ada"
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

Private Sub cmdNavKonsumen_Click(Index As Integer)
 With dtaKonsumen
    Select Case Index
        Case 0
            .Recordset.MovePrevious
            If .Recordset.BOF Then .Recordset.MoveFirst
        Case 1
            .Recordset.MoveFirst
        Case 2
            .Recordset.MoveLast
        Case 3
            .Recordset.MoveNext
            If .Recordset.EOF Then .Recordset.MoveLast
    End Select
 End With
End Sub

Private Sub cmdPilihBarang_Click()
    Text1(3).Text = txtBarang(0).Text
    Text1(4).Text = ViewFormat(txtBarang(7).Text)
    Text1(5).Text = "0"
    Text1(6).Text = ViewFormat(txtBarang(7).Text)
    
    lblRelasi(1).Caption = txtBarang(1).Text
    picBarang.Visible = False
End Sub

Private Sub cmdPilihKonsumen_Click()
    Text1(2).Text = txtKonsumen(0).Text
    lblRelasi(0).Caption = txtKonsumen(2).Text
    picKonsumen.Visible = False
End Sub

Private Sub cmdRelasi_Click(Index As Integer)
    If Index = 0 Then
        If StatKonsumen = False Then
            MsgBox "Data Konsumen masih kosong mohon diisi dahulu", vbInformation, "Pemberitahuan"
            Exit Sub
        End If
        dtaKonsumen.Refresh
        picKonsumen.Visible = True
    Else
        If StatBarang = False Then
            MsgBox "Data Barang masih kosong, mohon diisi dahulu", vbInformation, "Pemberitahuan"
            Exit Sub
        End If
        dtaBarang.Refresh
        picBarang.Visible = True
    End If
End Sub

Private Sub Data1_Reposition()
If Data1.Recordset.BOF Or Data1.Recordset.EOF Then Exit Sub
   If NoDisplay = False Then Data2Txt
   AdjustMasterGrid
   Data1.Caption = "Posisi " & Data1.Recordset.AbsolutePosition + 1 & " Dari " & Data1.Recordset.RecordCount
End Sub

Private Sub dtaBarang_Validate(Action As Integer, Save As Integer)
 If dtaBarang.Recordset.BOF Or dtaBarang.Recordset.EOF Then Exit Sub
    DataBarang2Txt
    lblPosisiBarang.Caption = "Posisi " & dtaBarang.Recordset.AbsolutePosition + 1 & " Dari " & dtaBarang.Recordset.RecordCount
End Sub

Private Sub dtaKonsumen_Reposition()
 If dtaKonsumen.Recordset.BOF Or dtaKonsumen.Recordset.EOF Then Exit Sub
    DataKonsumen2Txt
    lblPosisiKonsumen.Caption = "Posisi " & dtaKonsumen.Recordset.AbsolutePosition + 1 & " Dari " & dtaKonsumen.Recordset.RecordCount
End Sub

Private Sub Form_Activate()
  Dim x As Byte
  PaintForm3D Me
  
  StatBarang = True
  ModeEdit False
  With dtaBarang
    If .Recordset.EOF And .Recordset.BOF Then
        StatBarang = False
    End If
  End With
  
  StatKonsumen = True
  With dtaKonsumen
    If .Recordset.EOF And .Recordset.BOF Then
        StatKonsumen = False
    End If
  End With
  
End Sub

Private Sub Form_Load()
 Set rc = DB.OpenRecordset("tblPenjualan", dbOpenDynaset)
 Set rcBarang = DB.OpenRecordset("tblBarang", dbOpenDynaset)
 Set rcKonsumen = DB.OpenRecordset("tblKonsumen", dbOpenDynaset)
 Set Data1.Recordset = rc
 Set dtaBarang.Recordset = rcBarang
 Set dtaKonsumen.Recordset = rcKonsumen
 
 Set rcBarang = Nothing
 Set rcKonsumen = Nothing
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
    
    cmdRelasi(0).Visible = mVal
    cmdRelasi(1).Visible = mVal
    
    cmdData(0).Enabled = Not mVal
    cmdData(1).Enabled = Not mVal
    cmdData(2).Enabled = mVal
    cmdData(3).Enabled = mVal
    cmdData(4).Enabled = Not mVal
    DisableClose Me, mVal
    AdjustMasterGrid
    lblRelasi(0).Caption = ""
    lblRelasi(1).Caption = ""
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Index <> 1 Then
        If CekTgl(Text1(1).Text) = False Then Text1(1).SetFocus
    End If
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
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
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
    If x <= 3 Then
        Text1(x).Text = .Recordset(x)
      ElseIf x > 3 Then
        Text1(x).Text = ViewFormat(.Recordset(x))
    End If
  Next
 End With
End Sub

Private Sub Txt2Data()
  With Data1
    .Recordset(0) = Text1(0).Text
    .Recordset(1) = Text1(1).Text
    .Recordset(2) = Text1(2).Text
    .Recordset(3) = Text1(3).Text
    .Recordset(4) = Text1(4).Text
    .Recordset(5) = Text1(5).Text
    .Recordset(6) = Text1(6).Text
  End With
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
    .RecordSource = "SELECT * FROM tblPenjualan ORDER BY RIGHT(NoPenjualan,4) ASC"
    .Refresh
    If .Recordset.EOF And .Recordset.BOF Then
        AutoNumber = "P0001"
        Exit Function
    End If
    .Recordset.MoveLast
    No = Val(Right(.Recordset(0), 4)) + 1
    AutoNumber = "P" & String(4 - Len(No), "0") & No
  End With
End Function

Private Sub DataBarang2Txt()
Dim x As Byte
 With dtaBarang
    txtBarang(0).Text = .Recordset(0)
    txtBarang(1).Text = .Recordset(1)
    txtBarang(2).Text = .Recordset(2)
    txtBarang(3).Text = .Recordset(3)
    txtBarang(4).Text = .Recordset(4)
    txtBarang(5).Text = .Recordset(5)
    txtBarang(6).Text = .Recordset(6)
    txtBarang(7).Text = .Recordset(8)
 End With
End Sub

Private Sub DataKonsumen2Txt()
Dim x As Byte
 With dtaKonsumen
  For x = 0 To txtKonsumen.Count - 1
     txtKonsumen(x).Text = .Recordset(x)
  Next
 End With
End Sub

Private Sub txtCariBarang_Change()
Dim SQL As String
SQL = "KodeBarang like '*" & txtCariBarang.Text & "*' OR "
SQL = SQL & "NamaBarang like '*" & txtCariBarang.Text & "*' OR "
SQL = SQL & "Processor like '*" & txtCariBarang.Text & "*' OR "
SQL = SQL & "Storage like '*" & txtCariBarang.Text & "*' OR "
SQL = SQL & "VGA like '*" & txtCariBarang.Text & "*' OR "
SQL = SQL & "OpticalDrive like '*" & txtCariBarang.Text & "*'"
 With dtaBarang
       .Recordset.FindFirst SQL
       If .Recordset.NoMatch Then
           .Recordset.FindNext SQL
       End If
 End With
End Sub

Private Sub txtCariKonsumen_Change()
Dim SQL As String
SQL = "KodeKonsumen like '*" & txtCariKonsumen.Text & "*' OR "
SQL = SQL & "ID like '*" & txtCariKonsumen.Text & "*' OR "
SQL = SQL & "Nama like '*" & txtCariKonsumen.Text & "*' OR "
SQL = SQL & "TTL like '*" & txtCariKonsumen.Text & "*' OR "
SQL = SQL & "Alamat like '*" & txtCariKonsumen.Text & "*' OR "
SQL = SQL & "NoTelepon like '*" & txtCariKonsumen.Text & "*' OR "
SQL = SQL & "NoHP like '*" & txtCariKonsumen.Text & "*'"
 With dtaKonsumen
       .Recordset.FindFirst SQL
       If .Recordset.NoMatch Then
           .Recordset.FindNext SQL
       End If
 End With
End Sub

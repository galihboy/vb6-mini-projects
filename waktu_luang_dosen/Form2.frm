VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplikasi Pengelolaan Waktu Luang Dosen - Skripsi - Developed by Galih Hermawan @ IF. UNIKOM. v6. November. 2016"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16635
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   16635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRuang 
      Caption         =   "Ruang"
      Height          =   255
      Left            =   7440
      TabIndex        =   83
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H008080FF&
      Caption         =   "Ruang"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   82
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H008080FF&
      Caption         =   "Dosen"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   81
      Top             =   240
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H00FF8080&
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   4800
      TabIndex        =   80
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdKelolaPesan 
      Caption         =   "Kelola Pesan"
      Height          =   375
      Left            =   14400
      TabIndex        =   79
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtPesan 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   14400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   78
      Text            =   "Form2.frx":0000
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdUbahDosen 
      Caption         =   "Ubah Dosen"
      Height          =   375
      Left            =   1560
      TabIndex        =   77
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBersih 
      Caption         =   "Bersihkan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   76
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSemuaDosen 
      Caption         =   "Semua Dosen"
      Height          =   255
      Left            =   7440
      TabIndex        =   75
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdHapusDosen 
      Caption         =   "Hapus Dosen"
      Height          =   375
      Left            =   3000
      TabIndex        =   74
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdTambahDosen 
      Caption         =   "Tambah Dosen"
      Height          =   375
      Left            =   120
      TabIndex        =   73
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Berdasarkan Hari"
      Height          =   2895
      Left            =   120
      TabIndex        =   55
      Top             =   4800
      Width           =   14175
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   7
         Left            =   11880
         TabIndex        =   72
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   6
         Left            =   10200
         TabIndex        =   71
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   5
         Left            =   8520
         TabIndex        =   70
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   4
         Left            =   6840
         TabIndex        =   69
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   3
         Left            =   5160
         TabIndex        =   68
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   2
         Left            =   3480
         TabIndex        =   67
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   1
         Left            =   1800
         TabIndex        =   66
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ListBox lstDosen 
         Height          =   1425
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cboHari 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "16.45 - 18.00"
         Height          =   255
         Index           =   7
         Left            =   12000
         TabIndex        =   64
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "15.30 - 16.45"
         Height          =   255
         Index           =   6
         Left            =   10320
         TabIndex        =   63
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "14.15 - 15.30"
         Height          =   255
         Index           =   5
         Left            =   8640
         TabIndex        =   62
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "13.00 - 14.15"
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   61
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "11.45 - 13.00"
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   60
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "10.30 - 11.45"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   59
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "09.15 - 10.30"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   58
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "08.00 - 09.15"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12720
      TabIndex        =   54
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   11040
      TabIndex        =   53
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Area Pengelolaan Waktu Luang"
      Height          =   3975
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   11775
      Begin VB.Frame frmHari 
         Caption         =   "Sabtu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   5
         Left            =   9840
         TabIndex        =   47
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox sabtu0800 
            Caption         =   "08.00 - 09.15"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox sabtu0915 
            Caption         =   "09.15 - 10.30"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox sabtu1030 
            Caption         =   "10.30 - 11.45"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox sabtu1145 
            Caption         =   "11.45 - 13.00"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox sabtu1300 
            Caption         =   "13.00 - 14.15"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   1440
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   1440
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   1440
            Y1              =   2280
            Y2              =   2280
         End
      End
      Begin VB.Frame frmHari 
         Caption         =   "Jum'at "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   4
         Left            =   7920
         TabIndex        =   39
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox jumat1645 
            Caption         =   "16.45 - 18.00"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox jumat1530 
            Caption         =   "15.30 - 16.45"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox jumat1415 
            Caption         =   "14.15 - 15.30"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox jumat1300 
            Caption         =   "13.00 - 14.15"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox jumat1030 
            Caption         =   "10.30 - 11.45"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox jumat0915 
            Caption         =   "09.15 - 10.30"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox jumat0800 
            Caption         =   "08.00 - 09.15"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   1440
            Y1              =   1560
            Y2              =   1560
         End
      End
      Begin VB.Frame frmHari 
         Caption         =   "Kamis "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   3
         Left            =   6000
         TabIndex        =   30
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox kamis0800 
            Caption         =   "08.00 - 09.15"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox kamis0915 
            Caption         =   "09.15 - 10.30"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox kamis1030 
            Caption         =   "10.30 - 11.45"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox kamis1145 
            Caption         =   "11.45 - 13.00"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox kamis1300 
            Caption         =   "13.00 - 14.15"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox kamis1415 
            Caption         =   "14.15 - 15.30"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox kamis1530 
            Caption         =   "15.30 - 16.45"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox kamis1645 
            Caption         =   "16.45 - 18.00"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2880
            Width           =   1335
         End
      End
      Begin VB.Frame frmHari 
         Caption         =   "Rabu "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   2
         Left            =   4080
         TabIndex        =   21
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox rabu1645 
            Caption         =   "16.45 - 18.00"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox rabu1530 
            Caption         =   "15.30 - 16.45"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox rabu1415 
            Caption         =   "14.15 - 15.30"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox rabu1300 
            Caption         =   "13.00 - 14.15"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox rabu1145 
            Caption         =   "11.45 - 13.00"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox rabu1030 
            Caption         =   "10.30 - 11.45"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox rabu0915 
            Caption         =   "09.15 - 10.30"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox rabu0800 
            Caption         =   "08.00 - 09.15"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmHari 
         Caption         =   "Selasa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox selasa0800 
            Caption         =   "08.00 - 09.15"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox selasa0915 
            Caption         =   "09.15 - 10.30"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox selasa1030 
            Caption         =   "10.30 - 11.45"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox selasa1145 
            Caption         =   "11.45 - 13.00"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox selasa1300 
            Caption         =   "13.00 - 14.15"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox selasa1415 
            Caption         =   "14.15 - 15.30"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox selasa1530 
            Caption         =   "15.30 - 16.45"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox selasa1645 
            Caption         =   "16.45 - 18.00"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2880
            Width           =   1335
         End
      End
      Begin VB.Frame frmHari 
         Caption         =   "Senin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         Begin VB.CheckBox senin1645 
            Caption         =   "16.45 - 18.00"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox senin1530 
            Caption         =   "15.30 - 16.45"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CheckBox senin1415 
            Caption         =   "14.15 - 15.30"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox senin1300 
            Caption         =   "13.00 - 14.15"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox senin1145 
            Caption         =   "11.45 - 13.00"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox senin1030 
            Caption         =   "10.30 - 11.45"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox senin0915 
            Caption         =   "09.15 - 10.30"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox senin0800 
            Caption         =   "08.00 - 09.15"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.ListBox lstDaftarDosen 
      Height          =   3180
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Dosen"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim INIKunci As New iniTool
Dim iSimpan As Integer
Dim iUbah As Integer
Dim iStart As Integer
Dim sArea As String

Dim isiData() As arrData

Private Sub cboHari_Click()
    Dim strHari As String, strNamaDosen As String
    Dim i As Integer
    
    If iStart = 1 Then
        'strNamaDosen = vbNullString
        'MsgBox cboHari.List(cboHari.ListIndex)
        strHari = cboHari.List(cboHari.ListIndex)
        
        For i = 0 To 7
            lstDosen(i).Clear
        Next
        
        For i = 0 To lstDaftarDosen.ListCount - 1
            If BacaData(lstDaftarDosen.List(i), strHari & "0800") = 1 Then
                lstDosen(0).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "0915") = 1 Then
                lstDosen(1).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "1030") = 1 Then
                lstDosen(2).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "1145") = 1 Then
                lstDosen(3).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "1300") = 1 Then
                lstDosen(4).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "1415") = 1 Then
                lstDosen(5).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "1530") = 1 Then
                lstDosen(6).AddItem lstDaftarDosen.List(i)
            End If
            If BacaData(lstDaftarDosen.List(i), strHari & "1645") = 1 Then
                lstDosen(7).AddItem lstDaftarDosen.List(i)
            End If
        Next
    End If
End Sub

Private Sub cmdBersih_Click()
    Dim c As Control
    'If Check1.Value = vbUnchecked Then
      For Each c In Me.Controls
        If TypeOf c Is CheckBox Then
          'check for one of the six here
          c.Value = Unchecked
        End If
      Next
    'End If
End Sub

Private Sub cmdHapusDosen_Click()
    Dim nmDosen As String
    Dim tanya As Integer
    'Dim MyINI As New iniTool
    
    nmDosen = lstDaftarDosen.Text
    tanya = MsgBox("Apakah Anda yakin akan menghapus data " & sArea & " '" & nmDosen & "'", vbYesNo, "Konfirmasi Hapus")
    If tanya = vbYes Then
        INIKunci.DeleteSection nmDosen
    End If
    cmdSemuaDosen_Click
    lstDaftarDosen.ListIndex = 0
    'MsgBox INIKunci.FileName
End Sub

Private Sub cmdKelolaPesan_Click()
    frmPesan.Show , Me
End Sub

Private Sub cmdRuang_Click()
    Dim lstDosen() As String
    'INIKunci.FileName = alamatFileRuang
    lstDosen = GetINISectionNames(INIKunci.FileName, 30)
    'MsgBox UBound(lstDosen)
    Dim jml As Integer, i As Integer
    'jml = INIKunci.GetAllSections.Count
    jml = UBound(lstDosen)
    lstDaftarDosen.Clear
    For i = 0 To jml
        lstDaftarDosen.AddItem lstDosen(i)
    Next
    'lstDosen() = INIKunci.GetAllSections()
    'MsgBox INIKunci.GetAllSections.Count
End Sub

Private Sub cmdSemuaDosen_Click()
    Dim lstDosen() As String
    'INIKunci.FileName = alamatFileINI
    lstDosen = GetINISectionNames(INIKunci.FileName, 30)
    'MsgBox UBound(lstDosen)
    Dim jml As Integer, i As Integer
    'jml = INIKunci.GetAllSections.Count
    jml = UBound(lstDosen)
    lstDaftarDosen.Clear
    For i = 0 To jml
        lstDaftarDosen.AddItem lstDosen(i)
    Next
    'lstDosen() = INIKunci.GetAllSections()
    'MsgBox INIKunci.GetAllSections.Count
End Sub

Private Sub cmdSimpan_Click()
    Call AktifkanCheckbox(False)
    cmdSimpan.Enabled = False
    cmdUbah.Caption = "Ubah"
    iUbah = 0
    
    Call TulisINI(lstDaftarDosen.List(lstDaftarDosen.ListIndex))
    Call TampilKonfigurasi(lstDaftarDosen.List(lstDaftarDosen.ListIndex))
        
End Sub

Private Sub cmdTambahDosen_Click()
    Dim strNamaDosen As String
    Dim FileIsi As New FileTool
    
    strNamaDosen = StrConv(InputBox("Silakan masukkan Nama " & sArea, "Input " & sArea & " Baru"), vbProperCase)
    'MsgBox AdaDosen(strNamaDosen)
    If AdaDosen(strNamaDosen) = -1 Then 'dosen tidak ditemukan, tambahkan data baru
    'MsgBox strNamaDosen
        Call INIKunci.WriteValue(strNamaDosen, "senin0800", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin0915", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin1030", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin1145", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin1300", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin1415", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin1530", "0")
        Call INIKunci.WriteValue(strNamaDosen, "senin1645", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa0800", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa0915", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa1030", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa1145", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa1300", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa1415", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa1530", "0")
        Call INIKunci.WriteValue(strNamaDosen, "selasa1645", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu0800", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu0915", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu1030", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu1145", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu1300", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu1415", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu1530", "0")
        Call INIKunci.WriteValue(strNamaDosen, "rabu1645", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis0800", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis0915", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis1030", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis1145", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis1300", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis1415", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis1530", "0")
        Call INIKunci.WriteValue(strNamaDosen, "kamis1645", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat0800", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat0915", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat1030", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat1300", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat1415", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat1530", "0")
        Call INIKunci.WriteValue(strNamaDosen, "jumat1645", "0")
        Call INIKunci.WriteValue(strNamaDosen, "sabtu0800", "0")
        Call INIKunci.WriteValue(strNamaDosen, "sabtu0915", "0")
        Call INIKunci.WriteValue(strNamaDosen, "sabtu1030", "0")
        Call INIKunci.WriteValue(strNamaDosen, "sabtu1145", "0")
        Call INIKunci.WriteValue(strNamaDosen, "sabtu1300", "0")
        
        'alamatFile = App.path & "\" & nmFile
        'Text1.Text = FileIsi.ReadFile(alamatFile)
        'lstDaftarDosen.AddItem FileIsi.ReadFile(alamatFile)
        
        'Call TambahData(alamatFile, strNamaDosen)
        
        'lstDaftarDosen.Clear
        'AmbilIsiFile FileIsi.ReadFile(alamatFile)
        
        'lstDaftarDosen.ListIndex = 0
        'cmdSemuaDosen_Click
        KlikData
        'tampilkan data dosen bersangkutan
        lstDaftarDosen.ListIndex = GetListBoxIndex(lstDaftarDosen.HWND, strNamaDosen)
    Else
        MsgBox sArea & " dengan nama: " & strNamaDosen & " sudah ada.", vbExclamation, "Peringatan!"
        lstDaftarDosen.ListIndex = AdaDosen(strNamaDosen) 'tampilkan data dosen bersangkutan
        'MsgBox AdaDosen(strNamaDosen)
    End If
End Sub

Private Sub cmdTutup_Click()
    End
End Sub

Private Sub cmdUbah_Click()
    If iUbah = 1 Then
        iUbah = 0
        cmdUbah.Caption = "Ubah"
        cmdSimpan.Enabled = False
        cmdBersih.Enabled = False
        Call AktifkanCheckbox(False)
        Call TampilKonfigurasi(lstDaftarDosen.List(lstDaftarDosen.ListIndex))
    Else
        iUbah = 1
        cmdUbah.Caption = "Batal"
        cmdSimpan.Enabled = True
        cmdBersih.Enabled = True
        Call AktifkanCheckbox(True)
        
    End If
    'Call TampilKonfigurasi(lstDaftarDosen.List(lstDaftarDosen.ListIndex))
End Sub

Private Sub cmdUbahDosen_Click()
    Dim strNamaDosen As String, nmDefault As String
    Dim FileIsi As New FileTool
    Dim Rekaman() As arrData
    ReDim Rekaman(26)
    
    nmDefault = lstDaftarDosen.List(lstDaftarDosen.ListIndex)
    
    strNamaDosen = StrConv(InputBox("Silakan masukkan Nama " & sArea & " terbaru", "Ubah Data " & sArea, nmDefault), vbProperCase)
    
    'MsgBox strNamaDosen
    
    If strNamaDosen = "" Then
        MsgBox "Data " & sArea & " tidak diisi." & vbCr & "Perubahan data dibatalkan.", vbExclamation, "Peringatan!"
    ElseIf (StrPtr(strNamaDosen) = 0&) Then
        'MsgBox "Batal!"
    ElseIf strNamaDosen = nmDefault Then
        MsgBox "Data " & sArea & " tidak mengalami perubahan." & vbCr & "Perubahan data dibatalkan.", vbExclamation, "Peringatan!"
    Else
        If AdaDosen(strNamaDosen) = -1 Then 'dosen tidak ditemukan, tambahkan data baru
            'Backup Key dan Value
            BackupKeyValue nmDefault
            
            'Hapus Section berkaitan
            INIKunci.DeleteSection nmDefault
            
            'Tambahkan hasil update sbg data baru
             TulisINIUpdate strNamaDosen

            'cmdSemuaDosen_Click
            KlikData
            'tampilkan data dosen bersangkutan
            lstDaftarDosen.ListIndex = GetListBoxIndex(lstDaftarDosen.HWND, strNamaDosen)
        Else
            MsgBox sArea & " dengan nama: " & strNamaDosen & " sudah ada." & vbCr & "Perubahan data dibatalkan.", vbExclamation, "Peringatan!"
            'lstDaftarDosen.ListIndex = AdaDosen(strNamaDosen) 'tampilkan data dosen bersangkutan
        End If
    End If
    
    
End Sub

Private Sub Form_Load()
    
    Dim FileAda As New FileTool
    Dim FileIsi As New FileTool
    
    iSimpan = 0
    iStart = 0
    sArea = "Dosen"
    'alamatFile = App.path & "\" & nmFile
    alamatFileINI = App.path & "\" & nmFileINI
    alamatFileRuang = App.path & "\" & nmFileRUANG
        
    INIKunci.FileName = alamatFileINI
    
    'alamatFile = App.path & "\" & nmFile
    'Text1.Text = FileIsi.ReadFile(alamatFile)
    'lstDaftarDosen.AddItem FileIsi.ReadFile(alamatFile)
    'AmbilIsiFile FileIsi.ReadFile(alamatFile)
    'lstDaftarDosen.ListIndex = 0
    Call AktifkanCheckbox(False)
    cmdSemuaDosen_Click
    
    Call IsiHari
    iStart = 1
    
    'Pengelolaan pesan
    TampilkanPesan
End Sub

Private Sub AmbilIsiFile(strIsi As String)
    Dim filterIsi() As String
    Dim i As Integer
    
    filterIsi = Split(strIsi, vbNewLine)
    
    For i = LBound(filterIsi) To UBound(filterIsi)
        'Baris kosong jangan ditampilkan
        If filterIsi(i) <> vbNullString Then lstDaftarDosen.AddItem filterIsi(i)
    Next i
    
    jmlDosen = UBound(filterIsi)
    'MsgBox UBound(filterIsi)
End Sub

Private Sub lstDaftarDosen_Click()
    Dim INIKunci As New iniTool
    Dim sJudul As String
    
    'INIKunci.FileName = alamatFileINI
    'MsgBox INIKunci.GetValue(lstDaftarDosen.List(lstDaftarDosen.ListIndex), "senin0800")
    If opt(0).Value = True Then
        sJudul = "Pengelolaan Data Dosen : "
    Else
        sJudul = "Pengelolaan Data Ruang : "
    End If
    Frame1.Caption = sJudul & lstDaftarDosen.List(lstDaftarDosen.ListIndex) ' & " - " & lstDaftarDosen.ListIndex
    Call TampilKonfigurasi(lstDaftarDosen.List(lstDaftarDosen.ListIndex))
End Sub

Sub TampilKonfigurasi(strNamaDosen As String)
    Dim iCek As Integer
    Dim Ctl As Control
    
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is CheckBox Then
            iCek = Val(INIKunci.GetValue(strNamaDosen, Ctl.Name))
            'MsgBox iCek
            Ctl.Value = iCek
        End If
    Next
    
    Call AktifkanCheckbox(False)
End Sub

Sub BackupKeyValue(strNamaDosen As String)
    Dim iCek As Integer, sKey As String
    Dim Ctl As Control
    'Dim isi() As arrData
    Dim i As Integer
    i = 0
    
    'Dim sKey(50) As String
    
    'sKey(0) = INIKunci.GetAllKeys(strNamaDosen).Item(0)
    
    'MsgBox sKey
    'Exit Function
    
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is CheckBox Then
            i = i + 1
            ReDim Preserve isiData(i) As arrData
            isiData(i).sKunci = Ctl.Name
            isiData(i).sNilai = Val(INIKunci.GetValue(strNamaDosen, Ctl.Name))
            'iCek = Val(INIKunci.GetValue(strNamaDosen, Ctl.Name))
            'sKey = INIKunci.get
            'MsgBox iCek
            'Ctl.Value = iCek
        End If
    Next
    
    'MsgBox isi(i).sKunci & " - " & isi(i).sNilai
    
    'BackupKeyValue = isi
    'Call AktifkanCheckbox(False)
End Sub

Sub AktifkanCheckbox(status As Boolean)
    Dim Ctl As Control
    
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is CheckBox Then
            Ctl.Enabled = status
        End If
    Next
End Sub

Sub TulisINI(strNamaDosen As String)
    Dim iCek As Integer
    Dim Ctl As Control
    
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is CheckBox Then
            Call INIKunci.WriteValue(strNamaDosen, Ctl.Name, Ctl.Value)
            'iCek = Val(INIKunci.GetValue(strNamaDosen, ctl.Name))
            'MsgBox iCek
            'ctl.Value = iCek
        End If
    Next
    
    
    'Call INIKunci.WriteValue(strNamaDosen, "senin0800", senin0800.Value)
End Sub

Sub TulisINIUpdate(strNamaDosen As String)
    Dim iCek As Integer, i As Integer
    Dim Ctl As Control
    'i = 0
    
    'For Each Ctl In Me.Controls
        'If TypeOf Ctl Is CheckBox Then
    For i = 1 To UBound(isiData)
            'i = i + 1
            Call INIKunci.WriteValue(strNamaDosen, isiData(i).sKunci, isiData(i).sNilai)
            'iCek = Val(INIKunci.GetValue(strNamaDosen, ctl.Name))
            'MsgBox iCek
            'ctl.Value = iCek
       ' End If
    Next
    
    
    'Call INIKunci.WriteValue(strNamaDosen, "senin0800", senin0800.Value)
End Sub

Sub IsiHari()
    cboHari.AddItem "Senin"
    cboHari.AddItem "Selasa"
    cboHari.AddItem "Rabu"
    cboHari.AddItem "Kamis"
    cboHari.AddItem "Jumat"
    cboHari.AddItem "Sabtu"
    
    cboHari.ListIndex = 0
End Sub

Function BacaData(strNamaDosen As String, strWaktu As String) As Integer
    Dim iCek As Integer

    iCek = Val(INIKunci.GetValue(strNamaDosen, strWaktu))
    BacaData = iCek

End Function


Private Sub lstDosen_DblClick(Index As Integer)
    Dim i As Integer
    Dim s As String
    
    Form1.Show
    Form1.Text1 = vbNullString
    s = vbNullString
    For i = 0 To lstDosen(Index).ListCount - 1
        s = s & lstDosen(Index).List(i) & vbNewLine
    Next
    Clipboard.Clear
    Clipboard.SetText s, vbCFText
    If Clipboard.GetFormat(vbCFText) Then
        Form1.Text1.Text = Clipboard.GetText(vbCFText)
    End If
End Sub

Private Function AdaDosen(nm As String) As Integer
    Dim lstDosen() As String
    
    lstDosen = GetINISectionNames(INIKunci.FileName, 30)
    'MsgBox UBound(lstDosen)
    Dim jml As Integer, i As Integer
    'jml = INIKunci.GetAllSections.Count
    jml = UBound(lstDosen)
    'lstDaftarDosen.Clear
    For i = 0 To jml
        'MsgBox i & " " & lstDosen(i)
        If LCase(lstDosen(i)) = LCase(nm) Then
            AdaDosen = i
            Exit Function
        Else
            AdaDosen = -1
        End If
        'MsgBox i & " " & AdaDosen
    Next
End Function

Sub TampilkanPesan()
    Dim isiFile() As String, almFile As String
    Dim ft As New FileTool, X As Integer, i As Integer, iAktif As Integer, iTidakAktif As Integer
    Dim PesanAsli() As String, PesanAktif() As String, PesanTidakAktif() As String
    Dim PesanSementara() As String
    
    Dim status As Integer, nama As String, pesan As String
    
    almFile = App.path & "\" & nmFilePESAN
    
    'MsgBox almFile & " - " & ft.FileExist(almFile)
    isiFile = Split(ft.ReadFile(almFile), vbCrLf)
    
    For X = 0 To UBound(isiFile) - 1
        If Left(isiFile(X), 2) = "->" Then
            
            ReDim Preserve PesanAsli(i + 1)
            PesanAsli(i) = Mid(isiFile(X), 3, Len(isiFile(X)) - 2)
            'MsgBox i
            i = i + 1
            'Pesan aktif
        End If
    Next
    
    If i <= 0 Then
        MsgBox "Pesan Kosong atau isi tidak sesuai format.", vbExclamation, "Peringatan!"
    Else
        txtPesan = vbNullString
        txtPesan = "PESAN AKTIF :" & vbNewLine & vbNewLine
        'pesansementara=split(pesanasi
        For i = 0 To UBound(PesanAsli) - 1
            EkstrakPesan PesanAsli(i), status, nama, pesan
            If status = 1 Then 'pesan aktif
                txtPesan = txtPesan & UCase(nama) & vbNewLine
                txtPesan = txtPesan & pesan & vbNewLine & vbNewLine
            End If
            'lstPesanAktif.AddItem PesanAsli(i)
        Next
        
        txtPesan = txtPesan & "------------------------" & vbNewLine & _
                   "PESAN TIDAK AKTIF :" & vbNewLine & vbNewLine
        For i = 0 To UBound(PesanAsli) - 1
            EkstrakPesan PesanAsli(i), status, nama, pesan
            If status <> 1 Then 'pesan aktif
                txtPesan = txtPesan & UCase(nama) & vbNewLine
                txtPesan = txtPesan & pesan & vbNewLine & vbNewLine
            End If
            'lstPesanAktif.AddItem PesanAsli(i)
        Next
    End If
    
End Sub

Function statusPesan(Str As String) As Boolean
    Dim Pos As Integer, strStatus As String
    
    Pos = InStr(1, Str, "@")
    strStatus = Left(Str, Pos - 1)
    If Val(strStatus) = 1 Then
        statusPesan = True
    Else
        statusPesan = False
    End If
    'MsgBox strStatus
End Function

Sub EkstrakPesan(Str As String, ByRef statusPesan As Integer, ByRef sNama As String, _
                    ByRef sPesan As String)
    Dim strStatus() As String
    
    strStatus = Split(Str, "@")
    
    statusPesan = Val(strStatus(0))
    sNama = strStatus(1)
    sPesan = strStatus(2)
End Sub

Private Sub opt_Click(Index As Integer)
    KelolaData
End Sub

Sub KelolaData()
    If opt(0).Value = True Then 'dosen
        Label1.Caption = "Daftar Dosen"
        Frame1.Caption = "Area Pengelolaan Waktu Luang"
        cmdTambahDosen.Caption = "Tambah Dosen"
        cmdUbahDosen.Caption = "Ubah Dosen"
        cmdHapusDosen.Caption = "Hapus Dosen"
        sArea = "Dosen"
        
        INIKunci.FileName = alamatFileINI
        cmdSemuaDosen_Click
    Else 'ruang
        Label1.Caption = "Daftar Ruang"
        Frame1.Caption = "Area Pengelolaan Ruang Seminar"
        cmdTambahDosen.Caption = "Tambah Ruang"
        cmdUbahDosen.Caption = "Ubah Ruang"
        cmdHapusDosen.Caption = "Hapus Ruang"
        sArea = "Ruang"
        
        INIKunci.FileName = alamatFileRuang
        cmdRuang_Click
    End If
End Sub

Sub KlikData()
    If opt(0).Value = True Then
        cmdSemuaDosen_Click
    Else
        cmdRuang_Click
    End If
End Sub

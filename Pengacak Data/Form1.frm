VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplikasi Pengacak Data - Galih Hermawan"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Solusi Acak Langkah Demi Langkah"
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6615
      Begin VB.CommandButton cmdTutup 
         Caption         =   "TUTUP"
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   3120
         Width           =   2055
      End
      Begin VB.OptionButton optTampilkan 
         Caption         =   "Indeks dan Nilai"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optTampilkan 
         Caption         =   "Indeks"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optTampilkan 
         Caption         =   "Nilai ( Isi Data )"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdAmbilSeluruhData 
         Caption         =   "Seluruh Data"
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdAmbilData 
         Caption         =   "Ambil Data Per Langkah"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtHasil 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         TabIndex        =   9
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox txtDataKini 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtDataPenampung 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   "Tampilkan"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Hasil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Data yang diambil"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Penampung data tersisa"
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.TextBox txtAngkaRandom 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Developed by Galih Hermawan (http://galih.eu)"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   6615
   End
   Begin VB.Line Line2 
      BorderStyle     =   2  'Dash
      X1              =   6700
      X2              =   6700
      Y1              =   1560
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   6700
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "Contoh : 2 5 D 9 10 100 C"
      Height          =   375
      Left            =   3120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   3
      ToolTipText     =   "Klik untuk lihat contoh."
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Ket: data dapat berisi data berupa huruf atau angka yang di antara setiap data harus dipisahkan menggunakan spasi. "
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan data acak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Galih Hermawan @ http://galih.eu #3-8-2013
Dim strPenampung As String
Dim arrPenampung() As String
Dim iBatasBawah As Long
Dim iBatasAtas As Long
Dim iPilihan As Integer
Dim Mulai As Boolean

Private Sub cmdAmbilData_Click()
    Dim idxData As Integer, i As Integer
    Static arrPenampungBaru() As String
    Dim strKeluaran As String, strPenampungTemp As String
    Dim strKeluaranIndeks As String, strKeluaranIndeksNilai As String
    Static Pencacah As Long
    Static arrPenampungTemp() As String
    Static iAtas As Long
    Static iBawah As Long
    
    If Not Mulai Then
        If Trim(txtAngkaRandom) = "" Then Exit Sub
        
        strPenampung = Trim(txtAngkaRandom)
        arrPenampungTemp = Split(strPenampung, " ")
        
        iAtas = UBound(arrPenampungTemp)
        iBawah = LBound(arrPenampungTemp)
        
        Mulai = True
        txtHasil = vbNullString
    Else
        
    End If
    
    Pencacah = iAtas
    
    If Pencacah > 0 Then
        idxData = RandomPositive(0, Pencacah)
        'Menyimpan data random ke textbox data kini
        If iPilihan = 0 Then
            txtDataKini = arrPenampungTemp(idxData)
        ElseIf iPilihan = 1 Then
            txtDataKini = idxData
        Else
            txtDataKini = idxData & " - " & arrPenampungTemp(idxData)
        End If
        'Menyimpan data random ke textbox hasil
        txtHasil = txtHasil & " " & arrPenampungTemp(idxData)
        
        SalinDataBaru idxData, iAtas, arrPenampungTemp, arrPenampungBaru
        SalinUtuh arrPenampungBaru, arrPenampungTemp
        
        iAtas = UBound(arrPenampungTemp)
        
        For i = 0 To iAtas
            strKeluaran = strKeluaran & " " & arrPenampungBaru(i)
            strKeluaranIndeks = strKeluaranIndeks & " " & i
            strKeluaranIndeksNilai = strKeluaranIndeksNilai & " (" & i & "," & arrPenampungBaru(i) & ")"
        Next
        'Isi array uptodate
        If iPilihan = 0 Then
            txtDataPenampung = strKeluaran
        ElseIf iPilihan = 1 Then
            txtDataPenampung = strKeluaranIndeks
        Else
            txtDataPenampung = strKeluaranIndeksNilai
        End If
    Else
        'Menampilkan data yang diambil hasil pengacakan (terakhir)
        If iPilihan = 0 Then
            txtDataKini = arrPenampungTemp(0)
        ElseIf iPilihan = 1 Then
            txtDataKini = 0
        Else
            txtDataKini = 0 & " - " & arrPenampungTemp(0)
        End If
        'Isi array kosong
        txtDataPenampung = vbNullString
        'Menyimpan data random terakhir ke textbox hasil
        txtHasil = txtHasil & " " & arrPenampungTemp(0)
        Mulai = False
    End If
    
End Sub

Private Sub cmdAmbilSeluruhData_Click()
    Dim idxData As Integer, i As Integer
    Dim arrPenampungBaru() As String
    Dim strKeluaran As String, strPenampungTemp As String
    
    txtDataKini = vbNullString
    txtDataPenampung = vbNullString
    Mulai = False
    
    If Trim(txtAngkaRandom) = "" Then Exit Sub
    
    strPenampung = Trim(txtAngkaRandom)
    arrPenampung = Split(strPenampung, " ")
    
    iBatasAtas = UBound(arrPenampung)
    iBatasBawah = LBound(arrPenampung)
    
    Do While iBatasAtas > 0
        i = i + 1
        idxData = RandomPositive(iBatasBawah, iBatasAtas)
        strKeluaran = strKeluaran & " " & arrPenampung(idxData)
        
        SalinDataBaru idxData, iBatasAtas, arrPenampung, arrPenampungBaru
        SalinUtuh arrPenampungBaru, arrPenampung
        iBatasAtas = UBound(arrPenampung)
    Loop
    
    'Menampung data terakhir
    strKeluaran = strKeluaran & " " & arrPenampung(0)
    txtHasil = strKeluaran
    
End Sub

Private Sub cmdTutup_Click()
    End
End Sub

Private Sub Label3_Click()
    txtAngkaRandom.Text = "2 5 D 9 10 100 C"
End Sub

Private Sub SalinDataBaru(idx As Integer, iBatasAtas As Long, arrDataLama() As String, ByRef arrDataBaru() As String)
    Dim i As Integer, j As Integer
    'MsgBox idx
    ReDim arrDataBaru(iBatasAtas - 1)
    For i = 0 To idx - 1
        arrDataBaru(i) = arrDataLama(i)
    Next i
    For i = idx To iBatasAtas - 1
        arrDataBaru(i) = arrDataLama(i + 1)
    Next i
End Sub

Private Sub SalinUtuh(arrLama() As String, arrBaru() As String)
    Dim iAtasLama As Integer, i As Integer
    iAtasLama = UBound(arrLama)
    ReDim arrBaru(iAtasLama)
    For i = 0 To UBound(arrBaru)
        arrBaru(i) = arrLama(i)
    Next i
End Sub

Private Sub optTampilkan_Click(Index As Integer)
    iPilihan = Index
End Sub

VERSION 5.00
Begin VB.Form frmPesan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pengelolaan Pesan"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox cboPesan 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   6375
   End
   Begin VB.OptionButton optStatus 
      Caption         =   "Tidak Aktif"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton optStatus 
      Caption         =   "Aktif"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtPesan 
      Height          =   1455
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmPesan.frx":0000
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtNamaDosen 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "-"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Status Pesan"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Isi Pesan"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Dosen"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim strAlamatFile As String
Dim bTambah As Boolean, bUbah As Boolean
Dim strDiubah As String

Sub TampilkanPesan()
    Dim isiFile() As String, almFile As String
    Dim ft As New FileTool, X As Integer, i As Integer, iAktif As Integer, iTidakAktif As Integer
    Dim PesanAsli() As String, PesanAktif() As String, PesanTidakAktif() As String
    Dim PesanSementara() As String
    
    Dim status As Integer, nama As String, pesan As String, waktu As String
    
    'almFile = App.path & "\" & nmFilePESAN
    
    'MsgBox almFile & " - " & ft.FileExist(almFile)
    isiFile = Split(ft.ReadFile(strAlamatFile), vbCrLf)
    'MsgBox UBound(isiFile)
    For X = 0 To UBound(isiFile)
        'MsgBox isiFile(X)
        If Left(isiFile(X), 2) = "->" Then
            
            ReDim Preserve PesanAsli(i + 1)
            PesanAsli(i) = Mid(isiFile(X), 3, Len(isiFile(X)) - 2)
            'MsgBox i
            i = i + 1
            
            'Pesan aktif
        End If
        
    Next
    
    cboPesan.Clear
    If i <= 0 Then
        MsgBox "Pesan Kosong atau isi tidak sesuai format.", vbExclamation, "Peringatan!"
    Else
        'MsgBox UBound(PesanAsli)
        For i = 0 To UBound(PesanAsli) - 1
            'MsgBox PesanAsli(i)
            EkstrakPesan PesanAsli(i), status, nama, pesan, waktu
            cboPesan.AddItem UCase(nama) & " :>> " & pesan & " :>> " & status & " :>> " & waktu
        Next
        'cboPesan.ListIndex = 0
    End If
    
End Sub

Sub EkstrakPesan(str As String, ByRef statusPesan As Integer, ByRef sNama As String, _
                ByRef sPesan As String, ByRef sWaktu As String, Optional pemisah As String = "@")
    Dim strStatus() As String
    
    strStatus = Split(str, pemisah)
    
    'If UBound(strStatus) <= 0 Then Exit Sub
    'MsgBox str & vbNewLine & UBound(strStatus)
    
    If pemisah = "@" Then
        statusPesan = Val(strStatus(0))
        sNama = strStatus(1)
        sPesan = strStatus(2)
        sWaktu = strStatus(3)
    Else
        statusPesan = Val(strStatus(2))
        sNama = strStatus(0)
        sPesan = strStatus(1)
        sWaktu = strStatus(3)
    End If
End Sub

Private Sub cboPesan_Click()
    Dim status As Integer, nama As String, pesan As String, waktu As String
    Dim PesanAsli As String
    Dim i As Integer
   
    EkstrakPesan cboPesan.Text, status, nama, pesan, waktu, ":>>"
        'cboPesan.AddItem UCase(nama) & " : " & pesan

    txtNamaDosen = Trim(UCase(nama))
    txtPesan = Trim(pesan)
    If status = 1 Then
        optStatus(0).Value = vbChecked
    Else
        optStatus(1).Value = vbChecked
    End If
End Sub

Private Sub cmdHapus_Click()
    Dim strData As String
    Dim tanya As Integer
    Dim status As Integer, nama As String, pesan As String, waktu As String
 
    tanya = MsgBox("Anda yakin akan menghapus data berikut?" & vbCrLf & vbCrLf & _
                   "Nama Dosen: " & UCase(txtNamaDosen) & vbCrLf & _
                   "Pesan: " & txtPesan, _
                   vbYesNo + vbExclamation, "Peringatan!")
    
    
    
    If tanya = vbYes Then
        EkstrakPesan cboPesan.Text, status, nama, pesan, waktu, ":>>"
        strData = "->" & status & "@" & Trim(nama) & "@" & Trim(pesan) & "@" & Trim(waktu)
        'MsgBox "HAPUS" & strData
        HapusDataBaru strAlamatFile, LCase(strData)
        TampilkanPesan
        cboPesan.ListIndex = 0
        'MsgBox strData
    End If
    
End Sub

Private Sub cmdSimpan_Click()
    'MsgBox "ubah=" & bUbah & vbNewLine & "tambah=" & bTambah
    Dim strData As String, strStatus As String, waktu As String
    Dim semuaData() As String, DataOK As String, f As New FileTool
    Dim strCari As String, strBaru As String
    Dim statusLama As Integer, namaLama As String, pesanLama As String, waktuLama As String
    Dim pil As Integer
    
    pil = cboPesan.ListIndex
    
    semuaData = Split(f.ReadFile(strAlamatFile), vbNewLine)
    
    waktu = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
    
    If optStatus(0).Value = True Then
        strStatus = "1"
    Else
        strStatus = "0"
    End If
    
    Dim i As Integer
    For i = 0 To UBound(semuaData)
        If Len(Trim(semuaData(i))) <> 0 Then
            DataOK = DataOK + semuaData(i) + vbCrLf
        End If
    Next
    'Data untuk update
    EkstrakPesan cboPesan.Text, statusLama, namaLama, pesanLama, waktuLama, ":>>"
    strCari = "->" & statusLama & "@" & Trim(namaLama) & "@" & Trim(pesanLama) & "@" & Trim(waktuLama)
    strBaru = "->" & strStatus & "@" & Trim(txtNamaDosen) & "@" & Trim(Replace(txtPesan, vbCrLf, " ")) & "@" & waktu
    'Data untuk tambah
    strData = Trim(DataOK) & _
            "->" & strStatus & "@" & Trim(txtNamaDosen) & "@" & Trim(Replace(txtPesan, vbCrLf, " ")) & "@" & waktu
    
    If bTambah Then
        bTambah = False
        'TambahData strAlamatFile, strData
        TambahDataBaru strAlamatFile, strData
        cmdTambah.Caption = "Tambah"
        TampilkanPesan 'update combobox
        cboPesan.ListIndex = cboPesan.ListCount - 1
    ElseIf bUbah Then
        bUbah = False
        'txtNIM.Locked = False
        UbahDataBaru strAlamatFile, strCari, strBaru
        cmdUbah.Caption = "Ubah"
        TampilkanPesan 'update combobox
        cboPesan.ListIndex = pil
        'cboPesan.ListIndex = GetListBoxIndex(cboPesan.HWND, strCari)
    End If
    
    KunciTextBox True
    TombolNormal True
    
End Sub

Private Sub cmdTambah_Click()
    If cmdTambah.Caption = "Tambah" Then
        bTambah = True
        KunciTextBox False
        txtNamaDosen = vbNullString
        txtPesan = vbNullString
        optStatus(0).Value = True
        TombolNormal False
        cmdTambah.Enabled = True
        cmdTambah.Caption = "Batal"
        'TampilkanPesan
    Else
        bTambah = False
        KunciTextBox True
        'txtNamaDosen = vbNullString
        'txtPesan = vbNullString
        'optStatus(0).Value = True
        TombolNormal True
        cmdTambah.Caption = "Tambah"
        'cboPesan.ListIndex = 0
        cboPesan_Click
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "Ubah" Then
        bUbah = True
        KunciTextBox False
        'txtNamaDosen = vbNullString
        'txtPesan = vbNullString
        'optStatus(0).Value = True
        TombolNormal False
        cmdUbah.Enabled = True
        cmdUbah.Caption = "Batal"
    Else
        bUbah = False
        KunciTextBox True
        'txtNamaDosen = vbNullString
        'txtPesan = vbNullString
        'optStatus(0).Value = True
        TombolNormal True
        cmdUbah.Caption = "Ubah"
        'cboPesan.ListIndex = 0
        cboPesan_Click
    End If
End Sub

Private Sub Form_Load()
    strAlamatFile = App.path & "\" & nmFilePESAN
    TampilkanPesan
    cboPesan.ListIndex = 0
    KunciTextBox True
    TombolNormal True
    
    'Dim waktu As String
    'waktu = Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
    'MsgBox waktu
End Sub

Private Sub TambahData(strAlamatFile As String, strData As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Open strAlamatFile For Append As #iFileNo
        Print #iFileNo, strData
    Close #iFileNo
End Sub

Private Sub TambahDataBaru(strAlamatFile As String, strData As String)
    Dim iFileNo As Integer
    iFileNo = FreeFile
    
    Open strAlamatFile For Output As #iFileNo
        Print #iFileNo, Trim(strData)
    Close #iFileNo
End Sub

Sub KunciTextBox(boolNilai As Boolean)
    txtNamaDosen.Enabled = Not boolNilai
    txtPesan.Enabled = Not boolNilai
    optStatus.Item(0).Enabled = Not boolNilai
    optStatus.Item(1).Enabled = Not boolNilai
End Sub

Sub TombolNormal(boolNilai As Boolean)
    cmdTambah.Enabled = boolNilai
    cmdUbah.Enabled = boolNilai
    cmdSimpan.Enabled = Not boolNilai
    'cmdBatal.Enabled = Not boolNilai
    cmdHapus.Enabled = boolNilai
    cboPesan.Enabled = boolNilai
End Sub

Private Sub HapusDataBaru(strAlamatFile As String, strHapus As String)
    Dim iFileNo As Integer, i As Integer, j As Integer
    Dim strData As String
    Dim boolAda As Boolean
    Dim arrData() As String
    Dim f As New FileTool
    
    arrData = Split(f.ReadFile(strAlamatFile), vbNewLine)
    
    For i = 0 To UBound(arrData) - 1
        'MsgBox LCase(arrData(i)) & vbNewLine & LCase(strHapus) & vbNewLine & (LCase(arrData(i)) = LCase(strHapus))
        If Not (LCase(arrData(i)) = LCase(strHapus)) Then
            If Len(Trim(arrData(i))) <> 0 Then
                If i = UBound(arrData) - 1 Then
                    strData = strData & arrData(i)
                Else
                    strData = strData & arrData(i) & vbNewLine
                End If
            End If
            'MsgBox arrData(i)
        End If
    Next

    'boolAda = False
    iFileNo = FreeFile
    
    Open strAlamatFile For Output As #iFileNo
        Print #iFileNo, Trim(strData)
    Close #iFileNo
End Sub

Private Sub UbahDataBaru(strAlamatFile As String, strCari As String, strBaru As String)
    Dim iFileNo As Integer, i As Integer, j As Integer
    Dim strData As String
    Dim boolAda As Boolean
    Dim arrData() As String
    Dim f As New FileTool
    
    arrData = Split(f.ReadFile(strAlamatFile), vbNewLine)
    
    For i = 0 To UBound(arrData) - 1
        'MsgBox LCase(arrData(i)) & vbNewLine & LCase(strHapus) & vbNewLine & (LCase(arrData(i)) = LCase(strHapus))
        If Not (LCase(arrData(i)) = LCase(strCari)) Then
            If Len(Trim(arrData(i))) <> 0 Then
                If i = UBound(arrData) - 1 Then
                    strData = strData & arrData(i)
                Else
                    strData = strData & arrData(i) & vbNewLine
                End If
            End If
        Else
            If i = UBound(arrData) - 1 Then
                strData = strData & strBaru
            Else
                strData = strData & strBaru & vbNewLine
            End If
        End If
    Next

    'boolAda = False
    iFileNo = FreeFile
    
    Open strAlamatFile For Output As #iFileNo
        Print #iFileNo, Trim(strData)
    Close #iFileNo
End Sub

Private Sub UbahData(strAlamatFile As String, strBaruUbah As String, strDataDiubah As String)
    Dim iFileNo As Integer, i As Integer, j As Integer
    Dim strData As String
    Dim boolAda As Boolean
    Dim arrData() As String
    iFileNo = FreeFile
    i = 0
    
    Open strAlamatFile For Input As #iFileNo
        Do While Not EOF(iFileNo)
            ReDim Preserve arrData(i + 1)
            Input #iFileNo, strData
            If strData = strDataDiubah Then
                'Data yang dicari ditemukan, isi dengan yang baru
                arrData(i) = strBaruUbah
                i = i + 1
            Else
                'Data yang tidak dicari dan panjang karakter > 0 disimpan di array
                If Len(Trim(strData)) > 0 Then
                    'ReDim Preserve arrData(i + 1)
                    arrData(i) = strData
                    i = i + 1
                End If
            End If
        Loop
    Close #iFileNo
    
    'Timpa isi file dengan yang ada di array
    Open strAlamatFile For Output As #iFileNo
        For j = 0 To i - 1
            Print #iFileNo, arrData(j)
        Next
    Close #iFileNo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.TampilkanPesan
End Sub

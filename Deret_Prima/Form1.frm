VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deret Prima"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   1455
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "20"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "2"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Developed by Galih Hermawan - http://Galih.EU"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4200
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label2 
      Caption         =   "Nilai Akhir :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nilai Awal  :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************
' Developed by Galih Hermawan
' http://Galih.EU
' http://if.web.id
'*******************************
Private Sub Command1_Click()
    Dim min As Integer, max As Integer
    min = CInt(Text1.Text)
    max = CInt(Text2.Text)
    Text3.Text = vbNullChar
    Call prima(min, max)
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub prima(min, max As Integer)
    Dim i As Integer, j As Integer
    Dim banyak As Integer, jml As Integer, rata As Double
    Dim prima As Boolean
    
    banyak = 0
    jml = 0
    
    For i = min To max
        prima = False
        If (i = 2) Then
            prima = True
        Else
            For j = 2 To (i - 1)
                If (i Mod j = 0) Then
                    prima = False
                    Exit For
                Else
                    prima = True
                End If
            Next j
        End If
        If prima = True Then
            Text3.Text = Text3.Text & i & " "
            banyak = banyak + 1
            jml = jml + i
        End If
    Next i
    rata = Round(jml / banyak, 2)
    Label4.Caption = "Banyak bilangan = " & banyak & vbCrLf & _
                     "Jumlah total = " & jml & vbCrLf & _
                     "Rata-rata = " & rata
End Sub


Private Sub Form_Load()
    Label4.Caption = "Banyak bilangan = 0" & vbCrLf & _
                     "Jumlah total = 0" & vbCrLf & _
                     "Rata-rata = 0"
End Sub

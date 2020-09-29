VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Menampilkan Tulisan Berjalan di StatusBar"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   2040
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2595
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Rizky Khapidsyah
'Source Code Dimulai dari Sini

Dim Counter As Integer

Private Sub Form_Load()
  Counter = 0
  Timer1.Interval = 50  'Atur kecepatannya di sini
  With StatusBar1
    .Panels(1).Width = 4000
    .Panels(1).Alignment = sbrRight
  End With
End Sub

Private Sub Timer1_Timer()
  Dim Kalimat As String
  Dim pnlX1 As Panel
  Set pnlX1 = StatusBar1.Panels(1)
      Kalimat = "Testing tulisan berjalan"
      Counter = Counter + 1
      DoEvents
      pnlX1.Text = TulisJalan(Counter, Kalimat, 150)
End Sub

Public Function TulisJalan(Hitung As Integer, _
strKalimat As String, Panjang As Integer)

  If Hitung = Len(strKalimat) + Panjang Then
     Hitung = 0
  ElseIf Hitung > Len(strKalimat) Then
     TulisJalan = strKalimat & Space(Hitung - _
                  Len(strKalimat))
  Else
     TulisJalan = Mid(strKalimat, 1, Hitung)
  End If
End Function



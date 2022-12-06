VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "JOGO DA VELHA"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   FillColor       =   &H00000040&
   ForeColor       =   &H00400040&
   Icon            =   "aula.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Height          =   615
         Index           =   0
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Height          =   615
         Index           =   1
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   2
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   4
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   5
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   6
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   7
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   8
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Index           =   0
         X1              =   840
         X2              =   5040
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Index           =   1
         X1              =   840
         X2              =   5040
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Index           =   0
         X1              =   2280
         X2              =   2280
         Y1              =   360
         Y2              =   3120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Index           =   1
         X1              =   3600
         X2              =   3600
         Y1              =   360
         Y2              =   3120
      End
   End
   Begin VB.CommandButton cmdfim 
      Caption         =   "Final"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdnovo 
      Caption         =   "Novo Jogo"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   495
      Left            =   4800
      TabIndex        =   23
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   495
      Left            =   4920
      TabIndex        =   22
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fernanda Lins Bandeira                                        n.º de matricula: 108098 "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   6000
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cristina Maria Xavier da Silva                             n.º de matricula: 109539"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Feito por:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblvelh 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblest2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label lblest1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jogador 2"
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jogador 1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim i As Integer
Dim jog1 As Integer
Dim jog2 As Integer
Dim jogo As Integer
Dim velha As Integer
Dim dif As Integer

Private Sub cmdfim_Click()
Dim RESPOSTA As Integer
If (jog1 > jog2) Then
  RESPOSTA = MsgBox("O Jogador 1 ganhou com " & jog1 & " vitorias", vbInformation, "Vencedor")
ElseIf (jog2 > jog1) Then
RESPOSTA = MsgBox("O Jogador 2 ganhou com " & jog2 & " vitorias", vbInformation, "Vencedor")
Else
RESPOSTA = MsgBox("Deu Empate", vbInformation, "Vencedor")
End If
RESPOSTA = MsgBox("Tem certeza que deseja sair ?", vbYesNo + vbCritical, "Saída")
If RESPOSTA = vbYes Then End

End Sub

Private Sub cmdnovo_Click()
For i = 0 To 8
    Command1(i).Caption = Empty
    Command1(i).BackColor = &H8000000F
    Command1(i).Enabled = True
Next i
dif = 0
dif = jogo - (jog1 + jog2)

If dif > 0 Then
    velha = velha + 1
End If


Text1.Locked = False
Text2.Locked = False
Frame1.Enabled = True
a = 0
jogo = jogo + 1

lblest1.Caption = "Jogador 1 = " & jog1
lblest2.Caption = "Jogador 2 = " & jog2
lblvelh.Caption = "Velha     = " & velha
Label7.Caption = "Jogos     = " & jogo
End Sub

Private Sub Command1_Click(Index As Integer)
Text1.Locked = True
Text2.Locked = True

If a = 0 Then
    Command1(Index).Caption = Text1.Text
    a = 1
    Command1(Index).Enabled = False
Else
    Command1(Index).Caption = Text2.Text
    a = 0
    Command1(Index).Enabled = False
End If

If (Command1(0).Caption = Text1.Text) And (Command1(1).Caption = Text1.Text) And (Command1(2).Caption = Text1.Text) Then
    Command1(0).BackColor = &HFF8080
    Command1(1).BackColor = &HFF8080
    Command1(2).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(3).Caption = Text1.Text) And (Command1(4).Caption = Text1.Text) And (Command1(5).Caption = Text1.Text) Then
    Command1(3).BackColor = &HFF8080
    Command1(4).BackColor = &HFF8080
    Command1(5).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(6).Caption = Text1.Text) And (Command1(7).Caption = Text1.Text) And (Command1(8).Caption = Text1.Text) Then
    Command1(6).BackColor = &HFF8080
    Command1(7).BackColor = &HFF8080
    Command1(8).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(0).Caption = Text1.Text) And (Command1(4).Caption = Text1.Text) And (Command1(8).Caption = Text1.Text) Then
    Command1(0).BackColor = &HFF8080
    Command1(4).BackColor = &HFF8080
    Command1(8).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(2).Caption = Text1.Text) And (Command1(4).Caption = Text1.Text) And (Command1(6).Caption = Text1.Text) Then
    Command1(2).BackColor = &HFF8080
    Command1(4).BackColor = &HFF8080
    Command1(6).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(0).Caption = Text1.Text) And (Command1(3).Caption = Text1.Text) And (Command1(6).Caption = Text1.Text) Then
    Command1(0).BackColor = &HFF8080
    Command1(3).BackColor = &HFF8080
    Command1(6).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(1).Caption = Text1.Text) And (Command1(4).Caption = Text1.Text) And (Command1(7).Caption = Text1.Text) Then
    Command1(1).BackColor = &HFF8080
    Command1(4).BackColor = &HFF8080
    Command1(7).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(2).Caption = Text1.Text) And (Command1(5).Caption = Text1.Text) And (Command1(8).Caption = Text1.Text) Then
    Command1(2).BackColor = &HFF8080
    Command1(5).BackColor = &HFF8080
    Command1(8).BackColor = &HFF8080
    Frame1.Enabled = False
    jog1 = jog1 + 1

ElseIf (Command1(0).Caption = Text2.Text) And (Command1(1).Caption = Text2.Text) And (Command1(2).Caption = Text2.Text) Then
    Command1(0).BackColor = &HFFC0FF
    Command1(1).BackColor = &HFFC0FF
    Command1(2).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1

ElseIf (Command1(3).Caption = Text2.Text) And (Command1(4).Caption = Text2.Text) And (Command1(5).Caption = Text2.Text) Then
    Command1(3).BackColor = &HFFC0FF
    Command1(4).BackColor = &HFFC0FF
    Command1(5).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1

ElseIf (Command1(6).Caption = Text2.Text) And (Command1(7).Caption = Text2.Text) And (Command1(8).Caption = Text2.Text) Then
    Command1(6).BackColor = &HFFC0FF
    Command1(7).BackColor = &HFFC0FF
    Command1(8).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1

ElseIf (Command1(0).Caption = Text2.Text) And (Command1(4).Caption = Text2.Text) And (Command1(8).Caption = Text2.Text) Then
    Command1(0).BackColor = &HFFC0FF
    Command1(4).BackColor = &HFFC0FF
    Command1(8).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1

ElseIf (Command1(2).Caption = Text2.Text) And (Command1(4).Caption = Text2.Text) And (Command1(6).Caption = Text2.Text) Then
    Command1(2).BackColor = &HFFC0FF
    Command1(4).BackColor = &HFFC0FF
    Command1(6).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1

ElseIf (Command1(0).Caption = Text2.Text) And (Command1(3).Caption = Text2.Text) And (Command1(6).Caption = Text2.Text) Then
    Command1(0).BackColor = &HFFC0FF
    Command1(3).BackColor = &HFFC0FF
    Command1(6).BackColor = &HFF8080
    Frame1.Enabled = False
    jog2 = jog2 + 1

ElseIf (Command1(1).Caption = Text2.Text) And (Command1(4).Caption = Text2.Text) And (Command1(7).Caption = Text2.Text) Then
    Command1(1).BackColor = &HFFC0FF
    Command1(4).BackColor = &HFFC0FF
    Command1(7).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1
ElseIf (Command1(2).Caption = Text2.Text) And (Command1(5).Caption = Text2.Text) And (Command1(8).Caption = Text2.Text) Then
    Command1(2).BackColor = &HFFC0FF
    Command1(5).BackColor = &HFFC0FF
    Command1(8).BackColor = &HFFC0FF
    Frame1.Enabled = False
    jog2 = jog2 + 1


End If

End Sub

Private Sub Form_Load()
a = 0

jog1 = 0
jog2 = 0
velha = 0

jogo = 1
dif = 0

lblest1.Caption = "Jogador 1 = " & jog1
lblest2.Caption = "Jogador 2 = " & jog2
lblvelh.Caption = "Velha     = " & velha
End Sub


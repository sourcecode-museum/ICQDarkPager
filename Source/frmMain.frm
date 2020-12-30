VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8865
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   4725
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   5910
      TabIndex        =   5
      Top             =   3465
      Width           =   1680
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   4
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3870
      Width           =   4290
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   150
      TabIndex        =   2
      Top             =   3030
      Width           =   4290
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   1
      Top             =   2160
      Width           =   4290
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   1305
      Width           =   4290
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4860
      Top             =   1410
   End
   Begin MSWinsockLib.Winsock SockPager 
      Left            =   4860
      Top             =   1935
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1950
      Left            =   5910
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1515
      Width           =   1680
   End
   Begin VB.Label lblTempo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99/99/9999"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   5010
      TabIndex        =   18
      Top             =   660
      Width           =   930
   End
   Begin VB.Label lblCaptions 
      BackStyle       =   0  'Transparent
      Caption         =   "ICQ-Ða®k Pager v2.0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   17
      Top             =   45
      Width           =   3120
   End
   Begin VB.Label lblCaptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   6
      Left            =   6945
      TabIndex        =   16
      Top             =   4365
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   6
      Left            =   6630
      Picture         =   "frmMain.frx":1086E
      Top             =   3945
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   8
      Left            =   6960
      Picture         =   "frmMain.frx":10BDB
      Top             =   3945
      Width           =   240
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   7275
      TabIndex        =   15
      Top             =   1215
      Width           =   330
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   5
      Left            =   6105
      TabIndex        =   14
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   615
      TabIndex        =   13
      Top             =   3585
      Width           =   975
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   615
      TabIndex        =   12
      Top             =   2745
      Width           =   750
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   615
      TabIndex        =   11
      Top             =   1875
      Width           =   525
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NickName:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   615
      TabIndex        =   10
      Top             =   1035
      Width           =   855
   End
   Begin VB.Label lblClicks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   1845
      TabIndex        =   9
      Top             =   405
      Width           =   495
   End
   Begin VB.Label lblClicks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   3255
      TabIndex        =   8
      Top             =   405
      Width           =   315
   End
   Begin VB.Label lblClicks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   6105
      TabIndex        =   7
      Top             =   4335
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   7
      Left            =   6855
      Picture         =   "frmMain.frx":10F48
      Top             =   3945
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   5
      Left            =   6525
      Picture         =   "frmMain.frx":112B5
      Top             =   3945
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   10
      Left            =   6615
      Picture         =   "frmMain.frx":11622
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   9
      Left            =   5970
      Picture         =   "frmMain.frx":1198F
      Top             =   4260
      Width           =   960
   End
   Begin VB.Label lblTempo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   3315
      TabIndex        =   6
      Top             =   660
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   5685
      Picture         =   "frmMain.frx":12BD3
      Stretch         =   -1  'True
      Top             =   3870
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   3
      Left            =   150
      Picture         =   "frmMain.frx":12FDB
      Top             =   3540
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   2
      Left            =   150
      Picture         =   "frmMain.frx":1338C
      Top             =   2700
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   150
      Picture         =   "frmMain.frx":1373D
      Top             =   1830
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   150
      Picture         =   "frmMain.frx":13AEE
      Top             =   975
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Option Explicit
Private mnLisSel        As Long
Private mnIdxLBLCLICK   As Integer

Const UINLista    As String = "UINList.dat"

Const LBLEnviar   As Integer = 0
Const LBLSair     As Integer = 1
Const LBLAbout    As Integer = 2

Const CAPTitulo   As Integer = 0
Const CAPStatus   As Integer = 6

Const IMGAdd1     As Integer = 5
Const IMGAdd2     As Integer = 6

Const IMGClear1   As Integer = 7
Const IMGClear2   As Integer = 8

Const IMGEnviar1  As Integer = 9
Const IMGEnviar2  As Integer = 10

Const TXTNick     As Integer = 1
Const TXTEmail    As Integer = 2
Const TXTAssunto  As Integer = 3
Const TXTMsg      As Integer = 4
Const TXTUin      As Integer = 5

Private Sub Form_Initialize()
   Call CarregarLista
   If SockPager.LocalIP = "127.0.0.1" Then
      Image1(IMGEnviar1).Enabled = False
      Image1(IMGEnviar2).Enabled = False
      lblClicks(LBLEnviar).Enabled = False
   Else
      Image1(IMGEnviar1).Enabled = True
      Image1(IMGEnviar2).Enabled = True
      lblClicks(LBLEnviar).Enabled = True
   End If
   SockPager.Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblClicks(mnIdxLBLCLICK).ForeColor = vbWhite
   If Button = vbLeftButton Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
   End If
End Sub

Private Sub Image1_Click(Index As Integer)
   Select Case Index
      Case Is = IMGAdd1, IMGAdd2
         If Len(Trim$(Text(TXTUin))) > 0 Then
            If IsNumeric(Trim(Text(TXTUin))) Then
               List1.AddItem Trim$(Text(TXTUin))
               Labels(5) = List1.ListCount
            End If
         End If
         Text(TXTUin) = ""
         
      Case Is = IMGClear1, IMGClear2
         If List1.ListCount > 0 Then
            If MsgBox("Deseja realmente Limpar sua Lista de Contatos?", _
                      vbQuestion + vbYesNo, _
                      "Limpar Lista de Contatos") = vbYes Then
               List1.Clear
            End If
         End If
      Case Is = IMGEnviar1, IMGEnviar2
         Call Enviar
   End Select
End Sub

Private Sub lblCaptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = CAPTitulo And Button = vbLeftButton Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
   End If
End Sub

Private Sub lblClicks_Click(Index As Integer)
   Select Case Index
      Case Is = LBLAbout
      Case Is = LBLEnviar
         Call Enviar
      Case Is = LBLSair
         Call SalvarLista
         On Error Resume Next
         SockPager.Close
         End
   End Select
End Sub

Private Sub lblClicks_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   mnIdxLBLCLICK = Index
   lblClicks(Index).ForeColor = &H80FFFF
End Sub

Private Sub CarregarLista()
   Dim sItem As String
   
   On Error GoTo TrataErro
   Open UINLista For Input As #1
      While Not EOF(1)
         Line Input #1, sItem
         List1.AddItem sItem
      Wend
      Labels(5) = List1.ListCount
   Close #1
   On Error GoTo 0
Exit Sub
   
TrataErro:
   Open UINLista For Output As #1
   Close #1
End Sub

Private Sub SalvarLista()
   Dim n As Long
   
   On Error Resume Next
   Open UINLista For Output As #1
      For n = 0 To List1.ListCount - 1
         Print #1, List1.List(n)
      Next
   Close #1
   On Error GoTo 0
End Sub

Private Function TratarEspaços(ByVal psValor) As String
   Dim sChar      As String
   Dim sRetorno   As String
   Dim n          As Long
   
   For n = 1 To Len(psValor)
      sChar = Mid$(psValor, n, 1)
      If sChar = " " Then
         sChar = "+"
      End If
      sRetorno = sRetorno + sChar
   Next
   TratarEspaços = sRetorno
End Function

Private Sub List1_ItemCheck(Item As Integer)
   On Error GoTo Sair
   List1.ItemData(Item) = Not List1.ItemData(Item)
   
   If List1.ItemData(Item) = True Then
      mnLisSel = mnLisSel + 1
   Else
      mnLisSel = mnLisSel - 1
   End If
Sair:
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
   If List1.ListIndex = -1 Then Exit Sub
   If KeyCode = vbKeyDelete Then
      List1.RemoveItem List1.ListIndex
      Labels(5) = List1.ListCount
      List1.SetFocus
   End If
End Sub

Private Sub SockPager_Connect()
   lblCaptions(CAPStatus) = "Enviando..."
   On Error Resume Next
   SockPager.SendData SockPager.Tag
   On Error GoTo 0
End Sub

Private Sub SockPager_Error(ByVal Number As Integer, _
                            Description As String, _
                            ByVal Scode As Long, _
                            ByVal Source As String, _
                            ByVal HelpFile As String, _
                            ByVal HelpContext As Long, _
                            CancelDisplay As Boolean)
   lblCaptions(CAPStatus) = "Erro..."
   SockPager.Tag = ""
   MsgBox Description, vbCritical, "Erro: " & Number
End Sub

Private Sub SockPager_SendComplete()
   lblCaptions(CAPStatus) = "Enviado..."
   SockPager.Tag = ""
End Sub

Private Sub Enviar()
   Dim sNick    As String
   Dim sEmail   As String
   Dim sMsg     As String
   Dim sAssunto As String
   Dim sDados   As String
   Dim sEnviar  As String
   Dim n        As Long
   Dim nUIN     As Long
   
   For n = 0 To List1.ListCount - 1
      If List1.ItemData(n) = True Then
         nUIN = List1.List(n)
         If Trim(Text(TXTNick)) = "" Then Text(TXTNick) = App.Title
         If Trim(Text(TXTEmail)) = "" Then Text(TXTEmail) = "www.caiuaficha.com.br"
         If Trim(Text(TXTAssunto)) = "" Then Text(TXTAssunto) = "Informativo"
         If Trim(Text(TXTMsg)) = "" Then Text(TXTMsg) = "Site onde você encontra-rá" & vbCrLf & "as mais variadas informações na NET"
         
         lblCaptions(CAPStatus) = "Iniciando..."
         
         SockPager.Close
         sNick = TratarEspaços(Text(TXTNick))
         sEmail = TratarEspaços(Text(TXTEmail))
         sAssunto = TratarEspaços(Text(TXTAssunto))
         sMsg = TratarEspaços(Text(TXTMsg))
         
         sDados = sDados & "from=" & sNick
         sDados = sDados & " &fromemail=" & sEmail
         sDados = sDados & " &subject=" & sAssunto
         sDados = sDados & " &body=" & sMsg
         sDados = sDados & " &to=" & nUIN
         sDados = sDados & " &Send= Heliomar"
         
         sEnviar = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
         sEnviar = sEnviar & "Referer: http://wwp.mirabilis.com" & vbCrLf
         sEnviar = sEnviar & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
         sEnviar = sEnviar & "Connection: Keep-Alive" & vbCrLf
         sEnviar = sEnviar & "Host: wwp.mirabilis.com:80" & vbCrLf
         sEnviar = sEnviar & "Content-type: application/x-www-form-urlencoded" & vbCrLf
         sEnviar = sEnviar & "Content-length: " & Len(sDados) & vbCrLf
         sEnviar = sEnviar & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
         sEnviar = sEnviar & sDados
         
         SockPager.Tag = sEnviar
         SockPager.Connect "wwp.mirabilis.com", 80
      End If
   Next
End Sub

Private Sub Text_GotFocus(Index As Integer)
   Text(Index).BorderStyle = 1
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Text(Index).BorderStyle = 0
End Sub

Private Sub Timer1_Timer()
   lblTempo(0) = Time
   lblTempo(1) = Date
End Sub

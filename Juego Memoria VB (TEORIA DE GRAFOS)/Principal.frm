VERSION 5.00
Begin VB.Form Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memoria"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   4800
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   19
      Left            =   4560
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   19
      Top             =   3600
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   18
      Left            =   3480
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   18
      Top             =   3600
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   17
      Left            =   2400
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   17
      Top             =   3600
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   16
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   16
      Top             =   3600
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   15
      Left            =   240
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   15
      Top             =   3600
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   14
      Left            =   4560
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   14
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   13
      Left            =   3480
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   13
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   12
      Left            =   2400
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   11
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   10
      Left            =   240
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   10
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   9
      Left            =   4560
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   9
      Top             =   1440
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   8
      Left            =   3480
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   8
      Top             =   1440
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   7
      Left            =   2400
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   7
      Top             =   1440
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   6
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   6
      Top             =   1440
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   5
      Left            =   240
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   1440
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   4
      Left            =   4560
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   360
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   3
      Left            =   3480
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   360
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   2
      Left            =   2400
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   360
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   1
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   360
      Width           =   1000
   End
   Begin VB.PictureBox Imagen 
      AutoRedraw      =   -1  'True
      Height          =   1000
      Index           =   0
      Left            =   240
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
   Begin VB.Label LblFallos 
      Caption         =   "Fallos:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label LblAciertos 
      Caption         =   "Aciertos:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label LblTiempo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   4800
      Width           =   1335
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ContadorSegundos As Integer
Public ContadorMinutos As Integer
Public ContadorHoras As Integer
Private MatrizMemoria(20) As Integer
Private VerticeGrafo As Integer
Private CodigoImagen1 As Integer
Private CodigoImagen2 As Integer
Private NumeroAciertos As Integer
Private NumeroFallos As Integer
Private IndicePrimeraImagen As Integer

Private Sub Form_Load()
    Call EstablecerImagenesIniciales
    Call EstablecerMatrizDeImagenes
    VerticeGrafo = 1
End Sub

Private Sub Imagen_Click(Index As Integer)
    Select Case VerticeGrafo
        Case 1:
            VerticeGrafo = 2
            CodigoImagen1 = MatrizMemoria(Index)
            Imagen(Index).PaintPicture LoadPicture(".\Imagenes\" & CStr(MatrizMemoria(Index)) & ".jpg"), 0, 0, 1000, 1000
            IndicePrimeraImagen = Index
        Case 2:
            VerticeGrafo = 1
            CodigoImagen2 = MatrizMemoria(Index)
            Imagen(Index).PaintPicture LoadPicture(".\Imagenes\" & CStr(MatrizMemoria(Index)) & ".jpg"), 0, 0, 1000, 1000
            If (Comparar(CodigoImagen1, CodigoImagen2)) Then
                NumeroAciertos = NumeroAciertos + 1
                MsgBox "ACERTADO", vbInformation, "Memoria"
                Imagen(IndicePrimeraImagen).Enabled = False
                Imagen(Index).Enabled = False
            Else
                NumeroFallos = NumeroFallos + 1
                MsgBox "INCORRECTO", vbInformation, "Memoria"
                Imagen(IndicePrimeraImagen).PaintPicture LoadPicture(".\Imagenes\Interrogacion.jpg"), 0, 0, 1000, 1000
                Imagen(Index).PaintPicture LoadPicture(".\Imagenes\Interrogacion.jpg"), 0, 0, 1000, 1000
            End If
            LblAciertos.Caption = "Aciertos: " & CStr(NumeroAciertos)
            LblFallos.Caption = "Fallos: " & CStr(NumeroFallos)
    End Select
    If (NumeroAciertos = 10) Then
        MsgBox "Fin del Juego", vbInformation, "GAME OVER"
    End If
End Sub

Private Sub Timer1_Timer()
    ContadorSegundos = ContadorSegundos + 1
    If (ContadorSegundos = 60) Then
        ContadorSegundos = 1
        ContadorMinutos = ContadorMinutos + 1
    End If
    If (ContadorMinutos = 60) Then
        ContadorMinutos = 1
        ContadorHoras = ContadorHoras + 1
    End If
    LblTiempo.Caption = CStr(ContadorHoras) + ":" + CStr(ContadorMinutos) + ":" + CStr(ContadorSegundos)
End Sub

Private Function Comparar(Codigo1 As Integer, Codigo2 As Integer) As Boolean
    If (Codigo1 = Codigo2) Then
        Comparar = True
    Else
        Comparar = False
    End If
End Function

Private Sub EstablecerImagenesIniciales()
    Dim i As Integer
    For i = 0 To 19
        Imagen(i).PaintPicture LoadPicture(".\Imagenes\Interrogacion.jpg"), 0, 0, 1000, 1000
    Next
End Sub

Private Sub EstablecerMatrizDeImagenes()
    Call Randomize
    Dim i As Integer
    Dim CodigoAleatorio As Integer
    For i = 0 To 19
        CodigoAleatorio = CInt(Int((10) * Rnd() + 1))
        While (VerificarExistenciaCodigo(CodigoAleatorio))
            CodigoAleatorio = CInt(Int((10) * Rnd() + 1))
        Wend
        MatrizMemoria(i) = CodigoAleatorio
    Next
End Sub

Private Function VerificarExistenciaCodigo(Codigo As Integer)
    Dim i As Integer
    Dim Resultado As Boolean
    Dim ContadorCodigos As Integer
    
    Resultado = False
    ContadorCodigos = 0

    For i = 0 To 20
        If (MatrizMemoria(i) = Codigo) Then
            ContadorCodigos = ContadorCodigos + 1
        End If
    Next
    If (ContadorCodigos > 1) Then
        Resultado = True
    End If
    
    VerificarExistenciaCodigo = Resultado
End Function

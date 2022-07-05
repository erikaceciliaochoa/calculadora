VERSION 5.00
Begin VB.Form frmCalculadora 
   Caption         =   "Calculadora"
   ClientHeight    =   4740
   ClientLeft      =   7740
   ClientTop       =   3735
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   3855
   Begin VB.CommandButton cmd_mmas 
      BackColor       =   &H00C0FFC0&
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   775
   End
   Begin VB.CommandButton cmd_mr 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   775
   End
   Begin VB.CommandButton cmd_mb 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   775
   End
   Begin VB.CommandButton cmd_9 
      BackColor       =   &H8000000A&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   775
   End
   Begin VB.CommandButton cmd_8 
      BackColor       =   &H8000000A&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   775
   End
   Begin VB.CommandButton cmd_7 
      BackColor       =   &H8000000A&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   775
   End
   Begin VB.CommandButton cmd_6 
      BackColor       =   &H8000000A&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   775
   End
   Begin VB.CommandButton cmd_5 
      BackColor       =   &H8000000A&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   775
   End
   Begin VB.CommandButton cmd_4 
      BackColor       =   &H8000000A&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   775
   End
   Begin VB.CommandButton cmd_igual 
      BackColor       =   &H00FFFFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmd_3 
      BackColor       =   &H8000000A&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton cmd_punto 
      BackColor       =   &H8000000A&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmd_0 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmd_1 
      BackColor       =   &H8000000A&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton cmd_2 
      BackColor       =   &H8000000A&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton cmd_menos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   775
   End
   Begin VB.CommandButton cmd_dividir 
      BackColor       =   &H00FFFFC0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmd_multiplicar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton cmd_mas 
      BackColor       =   &H00FFFFC0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   775
   End
   Begin VB.CommandButton cmd_borrar 
      BackColor       =   &H008080FF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   775
   End
   Begin VB.Label lblDisplay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   3510
      TabIndex        =   20
      Top             =   120
      Width           =   165
   End
   Begin VB.Shape shp_borde 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   705
      Left            =   120
      Top             =   120
      Width           =   3603
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A As Double  'ACUMULADOR
Public C As Double
Public op As String ' opciones
Public Cl As Boolean ' bandera para clear - tienen q poder venir dos numeros

Private Sub cmd_0_Click()
clear 'tiene q borrar
lblDisplay.Caption = lblDisplay.Caption + cmd_0.Caption
End Sub

Private Sub cmd_1_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_1.Caption
End Sub

Private Sub cmd_2_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_2.Caption
End Sub

Private Sub cmd_3_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_3.Caption
End Sub

Private Sub cmd_4_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_4.Caption
End Sub

Private Sub cmd_5_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_5.Caption
End Sub

Private Sub cmd_6_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_6.Caption
End Sub

Private Sub cmd_7_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_7.Caption
End Sub

Private Sub cmd_8_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_8.Caption
End Sub

Private Sub cmd_9_Click()
clear
lblDisplay.Caption = lblDisplay.Caption + cmd_9.Caption
End Sub

Private Sub cmd_borrar_Click()
lblDisplay.Caption = " "
End Sub

Private Sub cmd_dividir_Click()
Calcular
op = "div"
'A = A / Val(lblDisplay.Caption)
End Sub

Private Sub cmd_igual_Click()
    Select Case op
        Case "suma"
            C = A + Val(lblDisplay.Caption)
            lblDisplay.Caption = C
        Case "resta"
            C = A - Val(lblDisplay.Caption)
            lblDisplay.Caption = C
        Case "mult"
            C = A * Val(lblDisplay.Caption)
            lblDisplay.Caption = C
        Case "div"
            If Val(lblDisplay.Caption) <> 0 Then
                C = A / Val(lblDisplay.Caption)
                lblDisplay.Caption = C
            Else
                lblDisplay.Caption = "ERROR!"
            End If
    End Select
    op = ""
    A = 0
    C = 0
    Cl = True
End Sub

Private Sub cmd_mas_Click()
Calcular
op = "suma"
'A = A + Val(lblDisplay.Caption)
End Sub

Private Sub cmd_menos_Click()
Calcular
op = "resta"
'A = A - Val(lblDisplay.Caption)
End Sub

Private Sub cmd_multiplicar_Click()
Calcular
op = "mult"
'A = A * Val(lblDisplay.Caption)
End Sub

Private Sub cmd_punto_Click()
lblDisplay.Caption = lblDisplay.Caption + cmd_punto.Caption
End Sub

Private Sub Form_Load()
    lblDisplay.Caption = ""
    A = 0
    C = 0
    op = ""
    Cl = False
End Sub

Sub Calcular()
    Select Case op
        Case "suma"
            A = A + Val(lblDisplay.Caption)
           ' lblDisplay.Caption = A
        Case "resta"
            A = A - Val(lblDisplay.Caption)
            'lblDisplay.Caption = A
        Case "mult"
            A = A * Val(lblDisplay.Caption)
            'lblDisplay.Caption = A
        Case "div"
            If Val(lblDisplay.Caption) <> 0 Then A = A / Val(lblDisplay.Caption)
           ' lblDisplay.Caption = A
        Case Else
            A = Val(lblDisplay.Caption)
    End Select
    lblDisplay.Caption = A
    Cl = True
End Sub




Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Desea salir?", vbOKCancel + vbQuestion, "Atención") = vbOK Then
End
'Cancel = 0 ' ACEPTAR - podes salir
'Else
'Cancel = 1 ' CANCELAR - seguis en el programa
'End If

End Sub

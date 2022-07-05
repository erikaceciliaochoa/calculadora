VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   6165
   ClientTop       =   4440
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancelar"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ok"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()

End Sub

Private Sub Form_Click()
MsgBox "click"
End Sub

Private Sub Form_Load()
MsgBox "cargando.."
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "mouse down?"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "mouse up?"
End Sub

Private Sub Form_Terminate()
MsgBox "Saliendo.."
End Sub

Private Sub Form_Unload(Cancel As Integer)
' formatear el msgbox
If MsgBox("usted quiere salir?", vbOKCancel + vbQuestion, "Atención") = vbOK Then
Cancel = 0
Else
Cancel = 1 ' salir
End If



End Sub

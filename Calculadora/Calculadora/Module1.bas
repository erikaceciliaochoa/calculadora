Attribute VB_Name = "Module1"
Option Explicit
Public A As Double  'ACUMULADOR
Public C As Double
Public op As String ' opciones
Public Cl As Boolean
Sub clear()
If Cl = True Then
    lblDisplay.Caption = ""
    Cl = False
End If
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "StartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Close_Click()
    Unload Me
End Sub

Private Sub BtnStart_Click()
    
    Started = True
    
    With ActiveDocument.ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayRulers = False
    End With
    Application.ScreenUpdating = False
    Btn_Close.Enabled = False
    Me.Label1 = "Aguarde"
    Me.BtnStart.Locked = True
    
    Select Case Me.caption
        Case "Staves"
            Call mainStaves
        Case "Conjunto Porta Vento"
            Call mainCPV
        Case "Carro Torpedos"
            Call mainCT
        Case "Domos"
            Call mainCowper
        Case "Corpo dos Regeneradores"
            Call mainCorpReg
        Case "Conduto Retilineo"
            Call mainRetilineo
    End Select
    
    Started = False
     With ActiveDocument.ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayRulers = True
    End With
    Application.ScreenUpdating = True
    Btn_Close.Enabled = True
    Me.BtnStart.Locked = False
    
    If successFinal = True Then
        Me.Label1 = "Concluido"
    Else
        Me.Label1 = "Não Concluído"
    End If
    
    MsgBox "Processo finalizado"
    
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If Started = True Then
            Cancel = True
        End If
    End If
End Sub

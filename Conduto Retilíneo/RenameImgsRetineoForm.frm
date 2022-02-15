VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameImgsRetineoForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   OleObjectBlob   =   "RenameImgsRetineoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameImgsRetineoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private labelsHS(0 To 26) As String
Private textBoxsHS(0 To 26) As String

Private Sub CommandButtonOk_Click()
    'On Error GoTo fail
    
    Dim txcontrol As Control
  
    'Verifica se tem algum campo vazio
    For Each txcontrol In Me.Controls
        If TypeOf txcontrol Is TextBox Then
            If txcontrol.Tag = "R" And txcontrol = "" Then
                MsgBox "prencha todos campos Seu Burro!!!"
                Exit Sub
            End If
        End If
    Next
    
    preencheLabelAndTxtboxHS
        
    'Verificar se tem repetidos prenchidos
    Dim i As Integer, j As Integer
    For i = 0 To UBound(textBoxsHS)
        For j = (i + 1) To UBound(textBoxsHS)
            If (Format(textBoxsHS(i), "0000") = Format(textBoxsHS(j), "0000")) Then
                MsgBox "Algum nome repetido!!!" & textBoxsHS(i)
                Exit Sub
            End If
        Next j
    Next i
    
    'Renomeia as imagens
    If WriteLabelAndBoxToFiles(textBoxsHS, labelsHS) Then
        Exit Sub
    End If
    
    Unload Me
    Exit Sub
'fail:
'    MsgBox "Algum error " & Err.Description
End Sub

'renomeia dos arrays labels txtxbox pra os arquivos
Private Function WriteLabelAndBoxToFiles(ByRef arrTxtBox() As String, ByRef arrLabel() As String) As Integer

    Dim fs As Object, ctFolder As Object, ctAllFiles As Object
    Dim file As Object
    
    Static contRenamed As Integer, contRenamed2 As Integer
    Dim dif As Integer ' diferença de imagens
    Dim i As Integer
    Dim nameFormated As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ctFolder = fs.GetFolder(".\IR")
    Set ctAllFiles = ctFolder.Files
    
    For i = 0 To UBound(arrTxtBox)
        For Each file In ctAllFiles
        'file.path->caminho do arquivo, fs.getbasename->pega só o nome do arquivo sem caminho e extensão,
        'Right-> pega 4 caracteres da direita pra esquerda
            nameFormated = Right(fs.getbasename(file.Path), 4)
            If (Format(arrTxtBox(i), "0000") = Format(nameFormated, "0000")) Then
                file.name = arrLabel(i) & ".jpg"
                contRenamed = contRenamed + 1
                Exit For
            End If
        Next file
    Next i
    
    dif = (UBound(arrTxtBox) + 1) - contRenamed
    
    If dif Then
        MsgBox "Não possivel renomear " & dif & " imagens!Não Feche, e confira na pasta IR"
        WriteLabelAndBoxToFiles = contRenamed
        Exit Function
    End If
    
    Set ctFolder = fs.GetFolder(".\Tratadas")
    Set ctAllFiles = ctFolder.Files
    For i = 0 To UBound(arrTxtBox)
        For Each file In ctAllFiles
        'file.path->caminho do arquivo, fs.getbasename->pega só o nome do arquivo sem acmino e extensão,
        'Right-> pega 4 caracteres da direita pra esquerda
            nameFormated = Right(fs.getbasename(file.Path), 4)
            If (Format(arrTxtBox(i), "0000") = Format(nameFormated, "0000")) Then
                file.name = arrLabel(i) & ".jpg"
                contRenamed2 = contRenamed2 + 1
                Exit For
            End If
        Next file
    Next i
    
    dif = (UBound(arrTxtBox) + 1) - contRenamed2
    
    If dif Then
        MsgBox "Não possivel renomear " & dif & " imagens! Não Feche, e confira na pasta Tratada"
        WriteLabelAndBoxToFiles = contRenamed2
        Exit Function
    End If
    
    WriteLabelAndBoxToFiles = 0 ' ok
    
End Function
Private Sub preencheLabelAndTxtboxHS()

    'labelsHS(0) = Me.LabelHS348
    labelsHS(0) = Me.LabelHS808 'era no final
    labelsHS(1) = Me.LabelHS394
    labelsHS(2) = Me.LabelHS395
    labelsHS(3) = Me.LabelHS397
    labelsHS(4) = Me.LabelHS487
    labelsHS(5) = Me.LabelHS489
    labelsHS(6) = Me.LabelHS490
    labelsHS(7) = Me.LabelHS528
    labelsHS(8) = Me.LabelHS529
    labelsHS(9) = Me.LabelHS531
    labelsHS(10) = Me.LabelHS720
    labelsHS(11) = Me.LabelHS722
    labelsHS(12) = Me.LabelHS723
    labelsHS(13) = Me.LabelHS725
    labelsHS(14) = Me.LabelHS727
    labelsHS(15) = Me.LabelHS730
    labelsHS(16) = Me.LabelHS750
    labelsHS(17) = Me.LabelHS775
    'labelsHS(18) = Me.LabelHS779
    labelsHS(18) = Me.LabelHS807 'era no final
    labelsHS(19) = Me.LabelHS780
    labelsHS(20) = Me.LabelHS781
    labelsHS(21) = Me.LabelHS782
    labelsHS(22) = Me.LabelHS797
    labelsHS(23) = Me.LabelHS798
    labelsHS(24) = Me.LabelHS804
    labelsHS(25) = Me.LabelHS805
    labelsHS(26) = Me.LabelHS806
    'labelsHS(27) = Me.LabelHS807
    'labelsHS(28) = Me.LabelHS808
    
    
    'textBoxsHS(0) = Me.TextBoxHS348
    textBoxsHS(0) = Me.TextBoxHS808 'era no final
    textBoxsHS(1) = Me.TextBoxHS394
    textBoxsHS(2) = Me.TextBoxHS395
    textBoxsHS(3) = Me.TextBoxHS397
    textBoxsHS(4) = Me.TextBoxHS487
    textBoxsHS(5) = Me.TextBoxHS489
    textBoxsHS(6) = Me.TextBoxHS490
    textBoxsHS(7) = Me.TextBoxHS528
    textBoxsHS(8) = Me.TextBoxHS529
    textBoxsHS(9) = Me.TextBoxHS531
    textBoxsHS(10) = Me.TextBoxHS720
    textBoxsHS(11) = Me.TextBoxHS722
    textBoxsHS(12) = Me.TextBoxHS723
    textBoxsHS(13) = Me.TextBoxHS725
    textBoxsHS(14) = Me.TextBoxHS727
    textBoxsHS(15) = Me.TextBoxHS730
    textBoxsHS(16) = Me.TextBoxHS750
    textBoxsHS(17) = Me.TextBoxHS775
    'textBoxsHS(18) = Me.TextBoxHS779
    textBoxsHS(18) = Me.TextBoxHS807 'era no final
    textBoxsHS(19) = Me.TextBoxHS780
    textBoxsHS(20) = Me.TextBoxHS781
    textBoxsHS(21) = Me.TextBoxHS782
    textBoxsHS(22) = Me.TextBoxHS797
    textBoxsHS(23) = Me.TextBoxHS798
    textBoxsHS(24) = Me.TextBoxHS804
    textBoxsHS(25) = Me.TextBoxHS805
    textBoxsHS(26) = Me.TextBoxHS806
    'textBoxsHS(27) = Me.TextBoxHS807
    'textBoxsHS(28) = Me.TextBoxHS808
    
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If MsgBox("Deseja mesmo Fechar? O progreeso sera perdido", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal) = vbNo Then
            Cancel = True
        Else
            End
        End If
        
    End If
End Sub

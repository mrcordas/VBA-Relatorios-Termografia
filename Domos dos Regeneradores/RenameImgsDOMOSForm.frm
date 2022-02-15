VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameImgsDOMOSForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   OleObjectBlob   =   "RenameImgsDOMOSForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameImgsDOMOSForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private labelsPOS(0 To 3) As String
Private textBoxsPOS(0 To 3) As String
Private labelsHS(0 To 20) As String
Private textBoxsHS(0 To 20) As String

Private Sub CommandButtonOkHS_Click()

    Dim txcontrol As Control
  
    ' Verifica se tem algum campo vazio
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
                MsgBox "Algum nome repetido!!! " & textBoxsHS(i)
                Exit Sub
            End If
        Next j
    Next i
    
    'Renomeia as imagens
    Call WriteLabelAndBoxToFiles(textBoxsHS, labelsHS)
    
    Unload Me
     

End Sub

Private Sub CommandButtonOkPOS_Click()
    
    Dim txcontrol As Control

  
    ' Verifica se tem algum campo vazio
    For Each txcontrol In Me.Frame1.Controls
        If TypeOf txcontrol Is TextBox Then
            If txcontrol.Tag = "S" And txcontrol = "" Then
                MsgBox "prencha todos campos Seu Burro!!!"
                Exit Sub
            End If
        End If
    Next
    
    preencheLabelAndTxtboxPOS
        
    'Verificar se tem repetidos prenchidos
    Dim i As Integer, j As Integer
    For i = 0 To UBound(textBoxsPOS)
        For j = (i + 1) To UBound(textBoxsPOS)
            If (Format(textBoxsPOS(i), "0000") = Format(textBoxsPOS(j), "0000")) Then
                MsgBox "Algum nome repetido!!! " & textBoxsPOS(i)
                Exit Sub
            End If
        Next j
    Next i
    

    'Renomeia as imagens
    Call WriteLabelAndBoxToFiles(textBoxsPOS, labelsPOS)
    
    Unload Me
         
    
End Sub
'renomeia dos arrays labels txtxbox pra os arquivos
Private Sub WriteLabelAndBoxToFiles(ByRef arrTxtBox() As String, ByRef arrLabel() As String)

    Dim fs As Object, ctFolder As Object, ctAllFiles As Object
    Dim file As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set ctFolder = fs.GetFolder(".\IR\" & Me.LabelEquiGroup)
    Set ctAllFiles = ctFolder.Files
    Dim i As Integer
    Dim nameFormated As String
    For i = 0 To UBound(arrTxtBox)
        For Each file In ctAllFiles
        'file.path->caminho do arquivo, fs.getbasename->pega só o nome do arquivo sem caminho e extensão,
        'Right-> pega 4 caracteres da direita pra esquerda
            nameFormated = Right(fs.getbasename(file.Path), 4)
            If (Format(arrTxtBox(i), "0000") = Format(nameFormated, "0000")) Then
                file.name = arrLabel(i) & ".jpg"
                Exit For
            End If
        Next
    Next i
    
    Set ctFolder = fs.GetFolder(".\Tratadas\" & Me.LabelEquiGroup)
    Set ctAllFiles = ctFolder.Files
    For i = 0 To UBound(arrTxtBox)
        For Each file In ctAllFiles
        'file.path->caminho do arquivo, fs.getbasename->pega só o nome do arquivo sem acmino e extensão,
        'Right-> pega 4 caracteres da direita pra esquerda
            nameFormated = Right(fs.getbasename(file.Path), 4)
            If (Format(arrTxtBox(i), "0000") = Format(nameFormated, "0000")) Then
                file.name = arrLabel(i) & ".jpg"
                Exit For
            End If
        Next
    Next i
    
End Sub
Private Sub preencheLabelAndTxtboxPOS()

    labelsPOS(0) = Me.LabelPOS1
    labelsPOS(1) = Me.LabelPOS2
    labelsPOS(2) = Me.LabelPOS3
    labelsPOS(3) = Me.LabelPIROMETRO
    
    textBoxsPOS(0) = Me.TextBoxPOS1
    textBoxsPOS(1) = Me.TextBoxPOS2
    textBoxsPOS(2) = Me.TextBoxPOS3
    textBoxsPOS(3) = Me.TextBoxPIROMETRO
    
End Sub

Private Sub preencheLabelAndTxtboxHS()

    labelsHS(0) = Me.LabelHS744
    labelsHS(1) = Me.LabelHS745
    labelsHS(2) = Me.LabelHS779
    labelsHS(3) = Me.LabelHS824
    labelsHS(4) = Me.LabelHS825
    labelsHS(5) = Me.LabelHS796
    labelsHS(6) = Me.LabelHS798
    labelsHS(7) = Me.LabelHS801
    labelsHS(8) = Me.LabelHS805
    labelsHS(9) = Me.LabelHS826
    labelsHS(10) = Me.LabelHS748
    labelsHS(11) = Me.LabelHS802
    labelsHS(12) = Me.LabelHS807
    labelsHS(13) = Me.LabelHS820
    labelsHS(14) = Me.LabelHS823
    labelsHS(15) = Me.LabelHS799
    labelsHS(16) = Me.LabelHS808
    labelsHS(17) = Me.LabelHS822
    labelsHS(18) = Me.LabelHS827
    labelsHS(19) = Me.LabelHS828
    labelsHS(20) = Me.LabelHS829
    
    textBoxsHS(0) = Me.TextBoxHS744
    textBoxsHS(1) = Me.TextBoxHS745
    textBoxsHS(2) = Me.TextBoxHS779
    textBoxsHS(3) = Me.TextBoxHS824
    textBoxsHS(4) = Me.TextBoxHS825
    textBoxsHS(5) = Me.TextBoxHS796
    textBoxsHS(6) = Me.TextBoxHS798
    textBoxsHS(7) = Me.TextBoxHS801
    textBoxsHS(8) = Me.TextBoxHS805
    textBoxsHS(9) = Me.TextBoxHS826
    textBoxsHS(10) = Me.TextBoxHS748
    textBoxsHS(11) = Me.TextBoxHS802
    textBoxsHS(12) = Me.TextBoxHS807
    textBoxsHS(13) = Me.TextBoxHS820
    textBoxsHS(14) = Me.TextBoxHS823
    textBoxsHS(15) = Me.TextBoxHS799
    textBoxsHS(16) = Me.TextBoxHS808
    textBoxsHS(17) = Me.TextBoxHS822
    textBoxsHS(18) = Me.TextBoxHS827
    textBoxsHS(19) = Me.TextBoxHS828
    textBoxsHS(20) = Me.TextBoxHS829
    
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

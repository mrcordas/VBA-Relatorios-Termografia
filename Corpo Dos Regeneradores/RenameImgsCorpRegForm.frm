VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameImgsCorpRegForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10875
   OleObjectBlob   =   "RenameImgsCorpRegForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameImgsCorpRegForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private labelsHS(0 To 54) As String
Private textBoxsHS(0 To 54) As String

Private Sub CommandButtonOk_Click()

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
                MsgBox "Algum nome repetido!!!"
                Exit Sub
            End If
        Next j
    Next i
    
    'Renomeia as imagens
    Call WriteLabelAndBoxToFiles(textBoxsHS, labelsHS)
    
    Unload Me

End Sub

'renomeia dos arrays labels txtxbox pra os arquivos
Private Sub WriteLabelAndBoxToFiles(ByRef arrTxtBox() As String, ByRef arrLabel() As String)

    Dim fs As Object, irTratFolder As Object, irTratAllFiles As Object
    Dim file As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    

    Set irTratFolder = fs.GetFolder(".\IR")
    Set irTratAllFiles = irTratFolder.Files
    
    Dim nameFormated As String
    Dim i As Integer, cont As Integer
    
    For i = 0 To UBound(arrTxtBox)
        For Each file In irTratAllFiles
        'file.path->caminho do arquivo, fs.getbasename->pega só o nome do arquivo sem caminho e extensão,
        'Right-> pega 4 caracteres da direita pra esquerda
            nameFormated = Right(fs.getbasename(file.Path), 4)
            If (Format(arrTxtBox(i), "0000") = Format(nameFormated, "0000")) Then
                file.name = arrLabel(i) & ".jpg"
                Exit For
            End If
        Next
    Next i
    
    Set irTratFolder = fs.GetFolder(".\Tratadas")
    Set irTratAllFiles = irTratFolder.Files
    
    For i = 0 To UBound(arrTxtBox)
        For Each file In irTratAllFiles
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
Private Sub preencheLabelAndTxtboxHS()

    labelsHS(0) = Me.LabelHS001
    labelsHS(1) = Me.LabelHS002
    labelsHS(2) = Me.LabelHS003
    labelsHS(3) = Me.LabelHS004
    labelsHS(4) = Me.LabelHS736
    labelsHS(5) = Me.LabelHS737
    labelsHS(6) = Me.LabelHS769
    labelsHS(7) = Me.LabelHS789
    labelsHS(8) = Me.LabelHS802
    labelsHS(9) = Me.LabelHS803
    labelsHS(10) = Me.LabelHS835
    labelsHS(11) = Me.LabelHS812
    labelsHS(12) = Me.LabelHS813
    labelsHS(13) = Me.LabelHS814
    labelsHS(14) = Me.LabelHS816
    labelsHS(15) = Me.LabelHS831
    labelsHS(16) = Me.LabelHS832
    labelsHS(17) = Me.LabelHS833
    labelsHS(18) = Me.LabelREG1_QUEIMADOR_VS
    labelsHS(19) = Me.LabelREG1_QUEIMADOR_LD
    labelsHS(20) = Me.LabelHS755
    labelsHS(21) = Me.LabelHS757
    labelsHS(22) = Me.LabelHS759
    labelsHS(23) = Me.LabelHS760
    labelsHS(24) = Me.LabelHS772
    labelsHS(25) = Me.LabelHS773
    labelsHS(26) = Me.LabelHS777
    labelsHS(27) = Me.LabelHS790
    labelsHS(28) = Me.LabelHS792
    labelsHS(29) = Me.LabelHS794
    labelsHS(30) = Me.LabelHS796
    labelsHS(31) = Me.LabelHS820
    labelsHS(32) = Me.LabelREG2_QUEIMADOR_VS
    labelsHS(33) = Me.LabelREG2_QUEIMADOR_LD
    labelsHS(34) = Me.LabelHS007
    labelsHS(35) = Me.LabelHS008
    labelsHS(36) = Me.LabelHS011
    labelsHS(37) = Me.LabelHS740
    labelsHS(38) = Me.LabelHS742
    labelsHS(39) = Me.LabelHS745
    labelsHS(40) = Me.LabelHS778
    labelsHS(41) = Me.LabelHS808
    labelsHS(42) = Me.LabelHS810
    labelsHS(43) = Me.LabelHS822
    labelsHS(44) = Me.LabelHS827
    labelsHS(45) = Me.LabelHS830
    labelsHS(46) = Me.LabelHS834
    labelsHS(47) = Me.LabelREG3_QUEIMADOR_VS
    labelsHS(48) = Me.LabelREG3_QUEIMADOR_LD
    labelsHS(49) = Me.LabelHS548
    labelsHS(50) = Me.LabelHS549
    labelsHS(51) = Me.LabelHS774
    labelsHS(52) = Me.LabelREG4_QUEIMADOR_LE
    labelsHS(53) = Me.LabelREG4_QUEIMADOR_LD
    labelsHS(54) = Me.LabelHS775
    
    textBoxsHS(0) = Me.TextBoxHS001
    textBoxsHS(1) = Me.TextBoxHS002
    textBoxsHS(2) = Me.TextBoxHS003
    textBoxsHS(3) = Me.TextBoxHS004
    textBoxsHS(4) = Me.TextBoxHS736
    textBoxsHS(5) = Me.TextBoxHS737
    textBoxsHS(6) = Me.TextBoxHS769
    textBoxsHS(7) = Me.TextBoxHS789
    textBoxsHS(8) = Me.TextBoxHS802
    textBoxsHS(9) = Me.TextBoxHS803
    textBoxsHS(10) = Me.TextBoxHS835
    textBoxsHS(11) = Me.TextBoxHS812
    textBoxsHS(12) = Me.TextBoxHS813
    textBoxsHS(13) = Me.TextBoxHS814
    textBoxsHS(14) = Me.TextBoxHS816
    textBoxsHS(15) = Me.TextBoxHS831
    textBoxsHS(16) = Me.TextBoxHS832
    textBoxsHS(17) = Me.TextBoxHS833
    textBoxsHS(18) = Me.TextBoxREG1_QUEIMADOR_VS
    textBoxsHS(19) = Me.TextBoxREG1_QUEIMADOR_LD
    textBoxsHS(20) = Me.TextBoxHS755
    textBoxsHS(21) = Me.TextBoxHS757
    textBoxsHS(22) = Me.TextBoxHS759
    textBoxsHS(23) = Me.TextBoxHS760
    textBoxsHS(24) = Me.TextBoxHS772
    textBoxsHS(25) = Me.TextBoxHS773
    textBoxsHS(26) = Me.TextBoxHS777
    textBoxsHS(27) = Me.TextBoxHS790
    textBoxsHS(28) = Me.TextBoxHS792
    textBoxsHS(29) = Me.TextBoxHS794
    textBoxsHS(30) = Me.TextBoxHS796
    textBoxsHS(31) = Me.TextBoxHS820
    textBoxsHS(32) = Me.TextBoxREG2_QUEIMADOR_VS
    textBoxsHS(33) = Me.TextBoxREG2_QUEIMADOR_LD
    textBoxsHS(34) = Me.TextBoxHS007
    textBoxsHS(35) = Me.TextBoxHS008
    textBoxsHS(36) = Me.TextBoxHS011
    textBoxsHS(37) = Me.TextBoxHS740
    textBoxsHS(38) = Me.TextBoxHS742
    textBoxsHS(39) = Me.TextBoxHS745
    textBoxsHS(40) = Me.TextBoxHS778
    textBoxsHS(41) = Me.TextBoxHS808
    textBoxsHS(42) = Me.TextBoxHS810
    textBoxsHS(43) = Me.TextBoxHS822
    textBoxsHS(44) = Me.TextBoxHS827
    textBoxsHS(45) = Me.TextBoxHS830
    textBoxsHS(46) = Me.TextBoxHS834
    textBoxsHS(47) = Me.TextBoxREG3_QUEIMADOR_VS
    textBoxsHS(48) = Me.TextBoxREG3_QUEIMADOR_LD
    textBoxsHS(49) = Me.TextBoxHS548
    textBoxsHS(50) = Me.TextBoxHS549
    textBoxsHS(51) = Me.TextBoxHS774
    textBoxsHS(52) = Me.TextBoxREG4_QUEIMADOR_LE
    textBoxsHS(53) = Me.TextBoxREG4_QUEIMADOR_LD
    textBoxsHS(54) = Me.TextBoxHS775
    
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

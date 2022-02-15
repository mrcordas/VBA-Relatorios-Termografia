VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameImgsCTForm 
   Caption         =   "Renomeia Imagens"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   OleObjectBlob   =   "RenameImgsCTForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameImgsCTForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private labels(0 To 9) As String
Private textBoxs(0 To 9) As String

Private Sub CommandButtonOk_Click()
    Dim txcontrol As Control
  
    ' Verifica se tem algum campo vazio
    For Each txcontrol In Me.Controls
        If txcontrol.Tag = "R" And txcontrol = "" Then
            MsgBox "prencha todos campos Seu Burro!!!"
            Exit Sub
        End If
    Next
    
    'executa de fato as operaçoes
    
    preencheLabelAndTxtbox
    
    'Verificar se tem repetidos prenchidos
    Dim i As Integer, j As Integer, repetido As Boolean
    For i = 0 To UBound(textBoxs)
        For j = (i + 1) To UBound(textBoxs)
            If (Format(textBoxs(i), "0000") = Format(textBoxs(j), "0000")) Then
                MsgBox "Algum nome repetido!!!" & textBoxs(i)
                Exit Sub
            End If
        Next j
    Next i
    
    'Renomeia as imagens
    If WriteLabelAndBoxToFiles(textBoxs, labels) Then
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If MsgBox("Deseja mesmo Fechar? Voce  ainda não renomeou todos!!!", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal) = vbNo Then
            Cancel = True
        Else
            End
        End If
    End If
End Sub
'renomeia dos arrays labels txtxbox pra os arquivos
Private Function WriteLabelAndBoxToFiles(ByRef arrTxtBox() As String, ByRef arrLabel() As String) As Integer

    Dim fs As Object, ctFolder As Object, ctAllFiles As Object
    Dim file As Object
    
    Static contRenamed As Integer, contRenamed2 As Integer
    Dim i As Integer
    Dim nameFormated As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ctFolder = fs.GetFolder(".\IR\" & Me.LabelCT)
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
    
    If contRenamed <> (UBound(arrTxtBox) + 1) Then
        MsgBox "Não possivel renomear algumas imagens!Não Feche, confira os campos e na pasta IR"
        WriteLabelAndBoxToFiles = contRenamed
        Exit Function
    End If
    
    Set ctFolder = fs.GetFolder(".\Tratadas\" & Me.LabelCT)
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
    
    If contRenamed2 <> (UBound(arrTxtBox) + 1) Then
        MsgBox "Não possivel renomear algumas imagens!Não Feche, confira os campos e na pasta Tratada"
        WriteLabelAndBoxToFiles = contRenamed2
        Exit Function
    End If
    
    WriteLabelAndBoxToFiles = 0 ' ok
    
End Function
Private Sub preencheLabelAndTxtbox()
        
        'Array string para labels recebe os caption das labels
        
        labels(0) = Me.LabelConeLaEncosta
        labels(1) = Me.LabelCilindroEncosta
        labels(2) = Me.LabelConeLlEncosta
        labels(3) = Me.LabelConeLaRio
        labels(4) = Me.LabelCilindroRio
        labels(5) = Me.LabelConeLlRio
        labels(6) = Me.LabelCalotaLa
        labels(7) = Me.LabelCalotaLl
        labels(8) = Me.LabelTetoLa
        labels(9) = Me.LabelTetoLl
        
        'Array textBoxs recebe os texto value das TextBox
        textBoxs(0) = Me.TextBoxConeLaEncosta
        textBoxs(1) = Me.TextBoxCilindroEncosta
        textBoxs(2) = Me.TextBoxConeLlEncosta
        textBoxs(3) = Me.TextBoxConeLaRio
        textBoxs(4) = Me.TextBoxCilindroRio
        textBoxs(5) = Me.TextBoxConeLlRio
        textBoxs(6) = Me.TextBoxCalotaLa
        textBoxs(7) = Me.TextBoxCalotaLl
        textBoxs(8) = Me.TextBoxTetoLa
        textBoxs(9) = Me.TextBoxTetoLl
End Sub



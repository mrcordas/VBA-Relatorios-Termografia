Attribute VB_Name = "Comuns"
Option Explicit
'======================
'Variaveis Globais
'===================
Public Started As Boolean
Public successFinal As Boolean

'==============================================================================================================
''Esta função recebe uma string nome do relatorio(ex. Document.Name) para verificar se é o nome procurado
'==============================================================================================================

Function whatDoc(reportNameSearch As String, nameSearch As String) As Boolean
    
    If Len(reportNameSearch) > 0 And Len(nameSearch) > 0 Then
        If (InStr(1, reportNameSearch, nameSearch, vbTextCompare)) > 0 Then
            whatDoc = True
            Exit Function
        Else
            MsgBox "O documento ativo não é " & nameSearch
        End If
    End If
    whatDoc = False
End Function

'===================================================================================
'Esta função lê um arquivo de texto e retorna a primeira linha do mesmo como String
'===================================================================================

Function getTemperaturas(aquivoDeTemperatura As String)
    Dim fs, file, op As Object
    Dim temp As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set file = fs.getFile(aquivoDeTemperatura)
    Set op = file.OpenAsTextStream(1, -2) ' leitura, txto do sistema
    
    'temp = Split(op.readline, Chr(9)) 'chr(9) é tab
    
    temp = op.readline
    
    op.Close
    
    getTemperaturas = temp
  
End Function

Sub changeLocal()

    '......Define o local....................

    Dim strDiretorioDoDocumentoAtual As String
    strDiretorioDoDocumentoAtual = ActiveDocument.Path
   
    ChDrive (strDiretorioDoDocumentoAtual) 'usei chdrive porque se tiver em driver diferente
    ChDir (strDiretorioDoDocumentoAtual)
    
    MsgBox "Configurado diretorio Atual ---> " & CurDir

'.............................................................................
    
End Sub
'======================================================================
'Esta Função ver se tem algum processso excel aberto no sistema
'======================================================================
Public Function WarningTask() As Integer
    Dim str As String
    Dim app As Task
    Dim taskFound As Integer
    taskFound = 0 ' nenhum processo
    For Each app In Tasks
        str = app.name
        If InStr(1, str, "Excel", vbTextCompare) > 0 Then
           taskFound = taskFound + 1
           MsgBox "tarefa em execução: " & app.name
        End If
    Next app
    WarningTask = taskFound
    
End Function

Public Function argStrToMyArray(ByRef myArray() As String, ParamArray values() As Variant)
    
    Dim i As Integer
    If UBound(myArray) <> UBound(values) Then
        MsgBox "ERRO: numeros de argumentos diferente do tamanho do Array"
        End
    End If
    
    For i = 0 To UBound(values)
        myArray(i) = values(i)
    Next i
    
End Function
Public Function checkIRtratFolders()

    Dim fs As Object
    Dim name As Variant
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
        If Not (fs.FolderExists("IR") And fs.FolderExists("Tratadas")) Then
            MsgBox "Pastas não encontradas!!! Verifique Pastas IR e Tratadas", vbCritical
            End
        End If
       
End Function
'retorna -1 se a pasta é invalida ou 0 para vazio ou o numero de arquivos
Public Function IsEmptyFolder(ByRef strName As String) As Integer

    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not fs.FolderExists(strName) Then
        MsgBox "pasta " & strName & " não encontrada"
        IsEmptyFolder = -1
        Exit Function
    End If
    
    IsEmptyFolder = fs.GetFolder(strName).Files.Count
    
End Function

Attribute VB_Name = "ConjPortaVento"
Option Explicit
'=================================================================================================
'Esta função checa se existe as pastas e os arquivos nescessarios para o relatório, Recebe um array
'Das pastas pos IR/Tratadas e um Array com o nome dos Arquivos
'=================================================================================================
Private Function checkFilesAndFolders(ByRef strClasseEquipament() As String, ByRef strFileNameList() As String) As Boolean
    '............... Checa pasta principais IR e Tratadas ........................

    Dim ObjSistemaAquivos As Object
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    '=========== verificação 1====
    If Not (ObjSistemaAquivos.FolderExists("IR") And ObjSistemaAquivos.FolderExists("Tratadas")) Then
    
        MsgBox "Pastas IR e/ou Tratadas não encontradas:", vbCritical, "ERRO"
        
        checkFilesAndFolders = False ' retorno da função
        Exit Function
    
    End If
'.....................................................................................


    Dim msgErroNoFoundDiretorio As String
    Dim indice As Integer, i As Integer ' uso geral
    Dim flagErro As Boolean
    Dim contErro As Integer
    
    '=========== verificação 2 ====
    flagErro = False
    contErro = 0
    msgErroNoFoundDiretorio = "Não encontradas seguintes subpastas de IR e/ou Tradadas: " & vbCrLf
    
    
    For indice = 0 To UBound(strClasseEquipament)
        flagErro = ((ObjSistemaAquivos.FolderExists(".\IR\" & strClasseEquipament(indice))) And _
                (ObjSistemaAquivos.FolderExists(".\Tratadas\" & strClasseEquipament(indice))))

        If (Not flagErro) Then
            msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strClasseEquipament(indice) & vbCrLf
            contErro = contErro + 1
        End If
        
    Next indice
    
    If (contErro > 0) Then
            checkFilesAndFolders = False
            MsgBox msgErroNoFoundDiretorio
            Exit Function
    End If
'........................................................................................
    
   '=========== verificação 3====
    flagErro = False
    contErro = 0
    msgErroNoFoundDiretorio = "Arquivos não encontrados " & vbCrLf
    
    For indice = 0 To UBound(strClasseEquipament)
        For i = 0 To UBound(strFileNameList)
            
            flagErro = ObjSistemaAquivos.FileExists(".\IR\" & strClasseEquipament(indice) & "\" & strFileNameList(i) & "_LD" & ".jpg") _
            And ObjSistemaAquivos.FileExists(".\Tratadas\" & strClasseEquipament(indice) & "\" & strFileNameList(i) & "_LD" & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strClasseEquipament(indice) & "\" & strFileNameList(i) & "_LD" & ".jpg" & vbCrLf
                contErro = contErro + 1
            End If
            
            flagErro = ObjSistemaAquivos.FileExists(".\IR\" & strClasseEquipament(indice) & "\" & strFileNameList(i) & "_LE" & ".jpg") _
            And ObjSistemaAquivos.FileExists(".\Tratadas\" & strClasseEquipament(indice) & "\" & strFileNameList(i) & "_LE" & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strClasseEquipament(indice) & "\" & strFileNameList(i) & "_LE" & ".jpg" & vbCrLf
                contErro = contErro + 1
            End If
        Next i
    Next indice
    
     If (contErro > 0) Then
            checkFilesAndFolders = False
            MsgBox msgErroNoFoundDiretorio
            Exit Function
    End If
    checkFilesAndFolders = True
End Function

'*****************************************************************************************
'Esta função é chamada inumera vezes, insere todas informações e a imagem de cada Shape
'strEquipament -> é o diretório após IR ou Tratadas
'filename -> É cada arquivo .jpg dentro de strEquipament
'temperatura -> cada temperatura do arquivo de texto de temperatura
'*****************************************************************************************
Private Function getAndWriteImgDataCPV(ByRef strEquipament As String, ByRef strFileNameList() As String)

    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer
    Dim lados() As Variant, lado As Variant
    lados = Array("_LD", "_LE")
    
    For i = 0 To UBound(strFileNameList)
        
        For Each lado In lados
            ' insere e formata a imagem tradada LD
            nameShapeObj = UCase(strEquipament) & "_" & UCase(strFileNameList(i) & lado)
            fullPathFile = ".\Tratadas\" & strEquipament & "\" & strFileNameList(i) & lado & ".jpg"
            With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
                .Select
                If (.InlineShapes.Count <> 0) Then
                    .InlineShapes.Item(1).Delete
                End If
                .InlineShapes.AddPicture (fullPathFile)
                .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Width
                .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Height
            End With
            
            'coloca as informações de data das imagens originais LD
            fullPathFile = ".\IR\" & strEquipament & "\" & strFileNameList(i) & lado & ".jpg"
    
            str = Split(CStr(ObjSistemaAquivos.getFile(fullPathFile).DateLastModified))
            ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = str(0)
            ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = str(1)
            
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1
            
         Next lado
         
        DoEvents
    
    Next i
   
End Function
'*****************************************************************************************
Private Function getAndWriteGrafDataCPV(ByRef index As Integer, ByRef strEquipament As String, ByRef strFileNameList() As String)

    Dim i As Integer
    Dim ValueOfColuns As Integer ' linha pra achar a ultima inserida, coluna dos valores de cada lado
    Dim pathWorkbook As String, workBookName As String
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook
    
    Dim lados() As Variant, lado As Variant
    Dim strCharts() As Variant, strchart As Variant
    'Dim nameShapeObj As String
    
    workBookName = Choose(index, "Gráfico Saída Porta Vento.xlsx", "Gráfico DowLeg.xlsx", "Gráfico Joelho.xlsx", "Gráfico Nariz.xlsx")
    pathWorkbook = ActiveDocument.Path & "\Gráfico\" & workBookName
    
    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = False
    
    lados = Array("_LD", "_LE")
    strCharts = Array("VT-01~04", "VT-05~08", "VT-09~12", "VT-13~16", "VT-17~20", "VT-21~22")
    
    myWorkbook.Sheets("TEMPERATURA-LD~LE").Activate
    With myWorkbook.Sheets("TEMPERATURA-LD~LE")
        ValueOfColuns = 2
        appExcel.Cells(3, ValueOfColuns).Select ' linha 3, coluna 2 (B3)
        
        '======== Inicio Insere as Temperaturas ===========
        For i = 0 To UBound(strFileNameList)
            For Each lado In lados
                With ActiveDocument.Shapes(UCase(strEquipament) & "_" & UCase(strFileNameList(i) & lado)).GroupItems("Temp").TextFrame
                    .TextRange.Select
                    .TextRange.Delete
                    .TextRange.Text = "MAX= " & appExcel.Cells(3, ValueOfColuns) & "ºC"
                    .VerticalAnchor = msoAnchorBottom
                End With
                ValueOfColuns = ValueOfColuns + 1
            Next lado
        Next i
        '======== Fim Insere as Temperaturas =========
    End With
    
    
        '======== Inicio Insere o Grafico ===========
    For Each strchart In strCharts
    
        myWorkbook.Charts(strchart).Activate
        myWorkbook.Charts(strchart).ChartArea.Select
        myWorkbook.Charts(strchart).ChartArea.Copy
        With ActiveDocument.Shapes(UCase(strEquipament) & "_" & strchart & "_GRAFICO").TextFrame.TextRange
            .Select
            If .InlineShapes.Count > 0 Then
                .InlineShapes(1).Delete
            End If
            Selection.Paste
        End With
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
        '======== fim Insere o Grafico ===========
    Next strchart
    
    appExcel.CutCopyMode = False ' excel
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing
    
End Function
Sub mainCPV()
    
    successFinal = False
    
    Dim classeEquipament(0 To 3) As String
    Dim fileNameListTypeStr(0 To 21) As String
    Dim i As Integer
    
    argStrToMyArray classeEquipament, "Saida", "DownLeg", "Joelho", "Nariz"
    argStrToMyArray fileNameListTypeStr, "vt01", "vt02", "vt03", "vt04", "vt05", "vt06", "vt07", "vt08", "vt09", "vt10", "vt11", _
                         "vt12", "vt13", "vt14", "vt15", "vt16", "vt17", "vt18", "vt19", "vt20", "vt21", "vt22"
        
    
    If Not (checkFilesAndFolders(classeEquipament, fileNameListTypeStr)) Then
        Exit Sub
    End If
    
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If
    
    
    On Error GoTo trata
    
    '================= Barra de progresso =========================
    Dim maxProgBar As Integer
    maxProgBar = (UBound(fileNameListTypeStr) + 1) * 2 * (UBound(classeEquipament) + 1) + (UBound(classeEquipament) + 1) * 6 ' 2 de lado e 6 graficos cada equipamento
    StartForm.ProgressBar1.Value = 0
    StartForm.ProgressBar1.Max = maxProgBar
    '================= Barra de progresso =========================
    
    For i = 0 To UBound(classeEquipament)
                
        Call getAndWriteImgDataCPV(classeEquipament(i), fileNameListTypeStr)
        Call getAndWriteGrafDataCPV(i + 1, classeEquipament(i), fileNameListTypeStr)
    Next i
    
    ActiveDocument.Protect wdAllowOnlyReading, , "01552375609"
    
    successFinal = True
    
    Exit Sub

trata:
     MsgBox Err.Description & Err.Number
     End
'    Resume Next
    
End Sub

Sub startCPV()

    If WarningTask Then
        MsgBox "Feche os tarefas do excel em execução, Tente novamente"
        Exit Sub
    End If
    
    Call changeLocal
    
    
    If Not (whatDoc(ActiveDocument.name, "CPV")) Then Exit Sub
    
    StartForm.caption = "Conjunto Porta Vento"
    StartForm.Show
    
    ActiveDocument.SaveAs2 filename:=CurDir & "\RT-CPV-AFA 2021-XX"
    
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If
    
    ActiveDocument.Save
    
End Sub
'======================================================================
'                       PRA DEBUG
'======================================================================
Private Function deleteImgsCPV(ByRef strEquipament As String, ByRef strFileNameList() As String)


    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim i As Integer
    Dim lados() As Variant, lado As Variant
    Dim strCharts() As Variant, strchart As Variant
    
    lados = Array("_LD", "_LE")
    strCharts = Array("VT-01~04", "VT-05~08", "VT-09~12", "VT-13~16", "VT-17~20", "VT-21~22")
    
    For i = 0 To UBound(strFileNameList)
        
        For Each lado In lados
            nameShapeObj = UCase(strEquipament) & "_" & UCase(strFileNameList(i) & lado)
            With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
                .Select
                If (.InlineShapes.Count <> 0) Then
                    .InlineShapes.Item(1).Delete
                End If
            End With
            ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = "00/00/0000"
            ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = "00:00:00"
            
            With ActiveDocument.Shapes(nameShapeObj).GroupItems("Temp").TextFrame
                .TextRange.Select
                .TextRange.Delete
                .TextRange.Text = "MAX= " & "---" & "ºC"
                .VerticalAnchor = msoAnchorBottom
            End With
            
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1
            
         Next lado
         
        DoEvents
    
    Next i
   
    For Each strchart In strCharts
    
        With ActiveDocument.Shapes(UCase(strEquipament) & "_" & strchart & "_GRAFICO").TextFrame.TextRange
            .Select
            If .InlineShapes.Count > 0 Then
                .InlineShapes(1).Delete
            End If
        End With
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
        '======== fim Insere o Grafico ===========
    Next strchart
   
End Function

Attribute VB_Name = "CarroTorpedos"
Option Explicit
Private nEquipaments As Integer
Private classeEquipament() As String
Private pathWorkbook As String
''============================================================================================================
''Chamada pela botão ok do form "selectCtForm" escolhe os carros para prencher o array privado classeEquipament
''define os a quantidade e os carros que vão no relatório
''============================================================================================================
'Public Function changeCTs(strCaption As String, sel As Boolean)
'
'    If sel Then
'        nEquipaments = nEquipaments + 1
'        ReDim Preserve classeEquipament(nEquipaments - 1) As String ' n-1 por que lower to upper 0 a superior
'        classeEquipament(nEquipaments - 1) = strCaption
'    End If
'End Function

'============================================================================================================
'
'
'============================================================================================================
Private Function changeFoldersCTs()
    
    nEquipaments = 0
    
    Dim fs As Object, ctFolders As Object, ctFolder As Object, irTratFolder As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set irTratFolder = fs.GetFolder(".\Tratadas") ' aqui basta só um pois a pasta IR deve ter as mesmas
    Set ctFolders = irTratFolder.Subfolders
    For Each ctFolder In ctFolders
        nEquipaments = nEquipaments + 1
        ReDim Preserve classeEquipament(nEquipaments - 1) As String ' n-1 por que lower to upper 0 a superior
        classeEquipament(nEquipaments - 1) = ctFolder.name
    Next ctFolder
    
'    Dim i As Integer
'
'    For i = 0 To UBound(classeEquipament)
'        If Not fs.FolderExists(".\IR\" & classeEquipament(i)) Then
'            MsgBox ""
'        End If
'    Next i
    
End Function
'=========================================
'
'============================================
Public Function renameImgs()

    Dim response As VbMsgBoxResult
    Dim i As Integer
    
    response = MsgBox("Deseja Renomear as imagens? Lógico se vc já renomeou não precisa JUMENTO", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal)
    
    If response = vbYes Then
        For i = 0 To UBound(classeEquipament)
            RenameImgsCTForm.LabelCT.caption = classeEquipament(i)
            RenameImgsCTForm.Show
        Next i
   End If
   
End Function
'======================================================================
'Abre a janela de busca do arquivo de grafico
'========================================================================
Private Function searchChart() As Integer
    Dim filetest As FileDialog
    Set filetest = Application.FileDialog(msoFileDialogFilePicker)
    filetest.Filters.Add "Pasta de Trabalho Excel", "*.xlsx; *.xlsm", 1
    With filetest
        .AllowMultiSelect = False
        .InitialFileName = ActiveDocument.Path
        If .Show = -1 Then
            'MsgBox .SelectedItems(1)
            pathWorkbook = .SelectedItems(1)
        Else
            MsgBox "Tente Novamente, você deve selecionar o Grafico Imbecil!!!", vbCritical
            searchChart = 0
            Exit Function
        End If
    End With
    
    Set filetest = Nothing
    searchChart = 1 ' pra tudo ok
End Function
'=================================================================================================
'Esta função checa se existe as pastas e os arquivos nescessarios para o relatório, Recebe um array
'Das pastas pos IR/Tratadas e um Array com o nome dos Arquivos
'=================================================================================================
Private Function checkFilesAndFolders(ByRef strClasseEquipament() As String, ByRef strFileNameList() As String) As Boolean
    '............... Checa pasta principais IR e Tratadas ........................

    Dim ObjSistemaAquivos As Object
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    

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
        For i = 0 To UBound(strFileNameList) - 1
            flagErro = ObjSistemaAquivos.FileExists(".\IR\" & strClasseEquipament(indice) & "\" & strFileNameList(i) & ".jpg") And _
            ObjSistemaAquivos.FileExists(".\Tratadas\" & strClasseEquipament(indice) & "\" & strFileNameList(i) & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strClasseEquipament(indice) & "\" & strFileNameList(i) & ".jpg" & vbCrLf
                contErro = contErro + 1
            End If
        Next i
        
        flagErro = ObjSistemaAquivos.FileExists(".\Tratadas\" & strClasseEquipament(indice) & "\ESCALA.jpg")
        If (Not flagErro) Then
            msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strClasseEquipament(indice) & "\" & strFileNameList(i) & ".jpg" & vbCrLf
            contErro = contErro + 1
        End If
    Next indice
    
     If (contErro > 0) Then
            checkFilesAndFolders = False
            MsgBox msgErroNoFoundDiretorio
            Exit Function
    End If
    
    checkFilesAndFolders = True
    
End Function
'============================================================================================
'Esta Função salva o range dos marcadores, modifica o texto nos mesmos e então os recriam
'pois os marcadores são deletados quando escrevemos nos ranges, então são recriados
'Obs: Todo n vai corresponder ao carro certo
'============================================================================================
Private Function writeInBookMarks(n As Integer, nameCaption As String)
    StartForm.Label1.caption = "Processando Bookmarks " & nameCaption
    Dim ran As Range
    
    Set ran = ActiveDocument.Bookmarks("CTX" & n & "_NAME").Range
    ran.Text = Right(nameCaption, 2)
    ActiveDocument.Bookmarks.Add ("CTX" & n & "_NAME"), ran
    
    Set ran = ActiveDocument.Bookmarks("CTX" & n & "_NAME2").Range
    ran.Text = Right(nameCaption, 2)
    ActiveDocument.Bookmarks.Add ("CTX" & n & "_NAME2"), ran
    
    Set ran = ActiveDocument.Bookmarks("CTX" & n & "_NAME3").Range
    ran.Text = Right(nameCaption, 2)
    ActiveDocument.Bookmarks.Add ("CTX" & n & "_NAME3"), ran
    
    Set ran = ActiveDocument.Bookmarks("CTX" & n & "_NAME4").Range
    ran.Text = Right(nameCaption, 2)
    ActiveDocument.Bookmarks.Add ("CTX" & n & "_NAME4"), ran
    
End Function
'===============================================================================================
'Esta função insere as Imagens Tratadas, coleta e colocas as infos de datas e hora das originais
'obs: Todo n vai corresponder ao carro certo
'===============================================================================================
Private Function getAndWriteImgDataCT(n As Integer, ByVal strEquipament As String, ByRef strFileNameList() As String)

    StartForm.Label1.caption = "Processando imagens " & strEquipament
    
    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer
    
    'Coloca as imagens
    For i = 0 To (UBound(strFileNameList) - 1) ' cada imagem, colocado - 1 pra tira a "ESCALA" o ultimo indice
        nameShapeObj = "CTX" & n & "_" & UCase(strFileNameList(i))
        fullPathFile = ".\Tratadas\" & strEquipament & "\" & strFileNameList(i) & ".jpg"
        
        With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
            .Select
            If (.InlineShapes.Count <> 0) Then
                .InlineShapes.Item(1).Delete
            End If
            .InlineShapes.AddPicture (fullPathFile)
            .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Width
            .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Height
        End With
        
        'coloca as informações de data das imagens originais
        fullPathFile = ".\IR\" & strEquipament & "\" & strFileNameList(i) & ".jpg"
        
        str = Split(CStr(ObjSistemaAquivos.getFile(fullPathFile).DateLastModified))
        ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = str(0)
        ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = str(1)
        
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
         
        DoEvents
        
    Next i
    
    
    'ESCALA
    nameShapeObj = "CTX" & n & "_" & strFileNameList(UBound(strFileNameList))
    fullPathFile = ".\Tratadas\" & strEquipament & "\" & strFileNameList(UBound(strFileNameList)) & ".jpg"
    
    With ActiveDocument.Shapes(nameShapeObj).TextFrame.TextRange
            .Select
            If (.InlineShapes.Count <> 0) Then
                .InlineShapes.Item(1).Delete
            End If
            .InlineShapes.AddPicture (fullPathFile)
            .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).Width
            .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).Height
    End With
    StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
    
End Function
'======================================================================
'
'Todo n vai corresponder ao carro certo
'======================================================================
Private Function getAndWriteGrafDataCT(ByRef strEquipament() As String, ByRef strFileNameList() As String)

    Dim h As Integer, i As Integer
    
    Dim linhaPlan As Integer, ValueOfColuns As Integer ' linha pra achar a ultima inserida, coluna dos valores de cada lado
    Dim strchart As String
    'Dim pathWorkbook As String, workbookName As String '-
    
    'workbookName = "GR-CT-AFA 2020-22.xlsm" '-
    'pathWorkbook = ActiveDocument.Path & "\" & workbookName '-
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook

    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = False
    
    For h = 0 To UBound(strEquipament)
        StartForm.Label1.caption = "Processando Grafico e Tabela" & strEquipament(h)
        'insere a tabela, temperatuara e grafico
        myWorkbook.Sheets("CT" & "-" & Right(strEquipament(h), 2)).Activate
        With myWorkbook.Sheets("CT" & "-" & Right(strEquipament(h), 2))
            appExcel.Cells(5, 2).Select
            linhaPlan = 5
            Do Until appExcel.Cells(linhaPlan, 2).Value = ""
                linhaPlan = linhaPlan + 1
            Loop
    
            linhaPlan = linhaPlan - 1 ' Volta pra linha interessada
            ValueOfColuns = 11 ' 11 é a primeira coluna com os valores das temperaturas é variavel
            
            '======== Inicio Insere as Temperaturas ===========
            For i = 0 To (UBound(strFileNameList) - 1)
                With ActiveDocument.Shapes("CTX" & h + 1 & "_" & UCase(strFileNameList(i))).GroupItems("Temp").TextFrame
                .TextRange.Select
                .TextRange.Delete
                .TextRange.Text = "MAX= " & appExcel.Cells(linhaPlan, ValueOfColuns) & "ºC"
                .VerticalAnchor = msoAnchorBottom
                End With
                ValueOfColuns = ValueOfColuns + 1
            Next i
            '======== Fim Insere as Temperaturas =========
            
            
            '======== Inicio Insere a tabela ===========
            appExcel.Range("B2", appExcel.Cells(linhaPlan, 10)).Select
            
            appExcel.Selection.Copy
            
            With ActiveDocument.Shapes("CTX" & h + 1 & "_TABELA").TextFrame.TextRange
                .Select
                If .Tables.Count > 0 Then
                    .Tables(1).Delete
                End If
                Selection.Paste
                .Tables(1).AutoFitBehavior (wdAutoFitWindow)
            End With
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
            '======== Fim Insere a tabela ===========
        
            '======== Inicio Insere o Grafico ===========
            
            strchart = "Grafico_" & "CT" & "-" & Right(strEquipament(h), 2)
            myWorkbook.Charts(strchart).Activate
            myWorkbook.Charts(strchart).ChartArea.Select
            myWorkbook.Charts(strchart).ChartArea.Copy
            
            With ActiveDocument.Shapes("CTX" & h + 1 & "_GRAFICO").TextFrame.TextRange
                .Select
                If .InlineShapes.Count > 0 Then
                    .InlineShapes(1).Delete
                End If
                Selection.Paste
                
            End With
            
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
            
            '======== fim Insere o Grafico ===========
        
        End With
    Next h
    appExcel.CutCopyMode = False ' excel
   
    'appExcel.Workbooks.Count
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing

End Function
'======================================================================
'                               MAIN
'======================================================================
Sub mainCT()

'    On Error GoTo trata
    
    successFinal = False
    
    Dim fileNameList(0 To 10) As String
    
    argStrToMyArray fileNameList, "CONE_LA_ENCOSTA", "CILINDRO_ENCOSTA", "CONE_LL_ENCOSTA", "CONE_LA_RIO", "CILINDRO_RIO", _
                         "CONE_LL_RIO", "CALOTA_LA", "CALOTA_LL", "TETO_LA", "TETO_LL", "ESCALA"
                         
   
    
    If Not (checkFilesAndFolders(classeEquipament, fileNameList)) Then
        Exit Sub
    End If
    

    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If
    
    '================= Barra de progresso =========================
    Dim maxProgBar
    maxProgBar = (UBound(fileNameList) + 1 + 2) * nEquipaments ' +1 corrigi o array(por causa "ESCALA"), + 2 do grafico e tabela
    StartForm.ProgressBar1.Value = 0
    StartForm.ProgressBar1.Max = maxProgBar
    '================= Barra de progresso =========================
    
    'executa as tarefas
    Dim i As Integer
    For i = 0 To UBound(classeEquipament) ' cada carro
        Call writeInBookMarks(i + 1, classeEquipament(i))
        Call getAndWriteImgDataCT(i + 1, classeEquipament(i), fileNameList)
        'Call deleteImgs(i + 1, classeEquipament(i), fileNameListTypeStr)
    Next i
    
    
    'Essa ja contem o for acima de cada carro internamente para evitar abertura do excel muitas vezes
    Call getAndWriteGrafDataCT(classeEquipament, fileNameList)

    
    'oculta os os textos das paginas com os bookmaks que não compõem o relatorio
    If nEquipaments < 7 Then
        For i = 1 To nEquipaments
            ActiveDocument.Bookmarks("CTX" & i).Select
            Selection.Font.Hidden = False
        Next i
        
        For i = (nEquipaments + 1) To 7
            ActiveDocument.Bookmarks("CTX" & i).Select
            Selection.Font.Hidden = True
        Next i
    Else
        For i = 1 To nEquipaments
            ActiveDocument.Bookmarks("CTX" & i).Select
            Selection.Font.Hidden = False
        Next i
    End If
    
    ActiveDocument.Protect wdAllowOnlyReading, , "01552375609"
    
    successFinal = True
    Exit Sub
    
'trata:
'    MsgBox Err.Description & Err.Number
    'Resume Next
End Sub
'====================================================================
'Chamada pela barra de opçoes no word é a PRINCIPAL que inicia tudo
'====================================================================
Sub startCT()
    
    If WarningTask Then
        MsgBox "Feche os tarefas do excel em execução, Tente novamente"
        Exit Sub
    End If
    
    changeLocal
    
    If Not (whatDoc(ActiveDocument.name, "CT")) Then Exit Sub
    
    checkIRtratFolders
    
    changeFoldersCTs
    
    If nEquipaments > 0 Then
        
        renameImgs
        
        MsgBox "Selecione o grafico! Se não você será ofendido"
        If searchChart = 1 Then
            StartForm.caption = "Carro Torpedos"
            StartForm.Show
        End If
    Else
        MsgBox "Sem equipamento identificado"
    End If
    Erase classeEquipament
    
    ActiveDocument.SaveAs2 filename:=CurDir & "\RT-CT-AFA 2021-XX"
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If
    
    ActiveDocument.Save
End Sub
'======================================================================
'                       PRA DEBUG
'======================================================================
Private Function deleteImgsCT(n As Integer, ByVal strEquipament As String, ByRef strFileNameList() As String)

    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim i As Integer
    
    'IMAGENS LADOS
    For i = 0 To (UBound(strFileNameList) - 1) ' cada imagem, colocado - 1 pra tira a "ESCALA" o ultimo indice
        nameShapeObj = "CTX" & n & "_" & UCase(strFileNameList(i))
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
        
        DoEvents
        
    Next i
    
    'ESCALA
    nameShapeObj = "CTX" & n & "_" & strFileNameList(UBound(strFileNameList))
    
    With ActiveDocument.Shapes(nameShapeObj).TextFrame.TextRange
            .Select
            If (.InlineShapes.Count <> 0) Then
                .InlineShapes.Item(1).Delete
            End If
    End With
    
    'TABELA
    With ActiveDocument.Shapes("CTX" & n & "_TABELA").TextFrame.TextRange
        .Select
        If .Tables.Count > 0 Then
            .Tables(1).Delete
        End If
    End With
    
    ' GRAFICO
    With ActiveDocument.Shapes("CTX" & n & "_GRAFICO").TextFrame.TextRange
        .Select
        If .InlineShapes.Count > 0 Then
            .InlineShapes(1).Delete
        End If
    End With

End Function

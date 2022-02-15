Attribute VB_Name = "Staves"
Option Explicit
Dim equipamentGroup(0 To 6) As String 'aneis
'===========================================
'Preenche o array global equipamentGroup  e verifica se as pasta estão no local
'============================================
Private Function changeFoldersNamesX()

    Dim fs As Object
    Dim name As Variant
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    argStrToMyArray equipamentGroup, "Anel13", "Anel11", "Anel10", "Anel09", "Anel08", "Anel06", "Anel04"
    
    For Each name In equipamentGroup
        If Not (fs.FolderExists("IR\" & name) And fs.FolderExists("Tratadas\" & name)) Then
            MsgBox "Pastas não encontradas!!! Verifique Pastas dos Aneis", vbCritical
            End
        End If
    Next name
       
End Function
'==============================================
'Lembra-o de renomear as imagens com script .bat
'==============================================
Private Sub renameImgs()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("As Imagens Ja foram renomeadas com Script CMD?", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal)
    If response = vbNo Then
        MsgBox "Renomeie os Arquivos primeiro e tente novamente"
        End
    End If
    
End Sub
'===========================================
'Identifica os arquivos de grafico variaveis globais workbook
'===========================================
Private Function searchChart()
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not (fs.FileExists(".\Gráfico\Gráfico STAVES 04° ANEL.xlsx") And fs.FileExists(".\Gráfico\Gráfico STAVES 06° ANEL.xlsx") _
        And fs.FileExists(".\Gráfico\Gráfico STAVES 08° ANEL.xlsx") And fs.FileExists(".\Gráfico\Gráfico STAVES 09° ANEL.xlsx") _
        And fs.FileExists(".\Gráfico\Gráfico STAVES 10° ANEL.xlsx") And fs.FileExists(".\Gráfico\Gráfico STAVES 11° ANEL.xlsx") _
        And fs.FileExists(".\Gráfico\Gráfico STAVES 13° ANEL.xlsx")) Then
        
        MsgBox "ERRO: Verifique os nomes na pasta .\Grafico"
        End
    End If
End Function
Private Function checkFilesAndFolders(ByRef strEquipamentGroup() As String) As Boolean
    
    Dim ObjSistemaAquivos As Object
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
   
    Dim msgErroNoFoundDiretorio As String
    Dim j As Integer, i As Integer ' uso geral
    Dim flagErro As Boolean
    Dim contErro As Integer
    
    '=========== verificação 2 ====
    flagErro = False
    contErro = 0
    msgErroNoFoundDiretorio = "Não encontradas seguintes subpastas de IR e/ou Tradadas: " & vbCrLf
    
    
    For i = 0 To UBound(strEquipamentGroup)
        flagErro = ((ObjSistemaAquivos.FolderExists(".\IR\" & strEquipamentGroup(i))) And _
                (ObjSistemaAquivos.FolderExists(".\Tratadas\" & strEquipamentGroup(i))))

        If (Not flagErro) Then
            msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strEquipamentGroup(i) & vbCrLf
            contErro = contErro + 1
        End If
        
    Next i
    
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
    
    'verifica os aquivos
    Dim fileNameListTypeStr() As String
    
    For i = 0 To UBound(strEquipamentGroup)
        Select Case i
        Case 0 'anel 13
            ReDim fileNameListTypeStr(7)
            argStrToMyArray fileNameListTypeStr, "st01", "st04", "st08", "st11", "st15", "st18", "st22", "st25"
        Case 1 'anel 11
            ReDim fileNameListTypeStr(13)
            argStrToMyArray fileNameListTypeStr, "st01", "st03", "st05", "st07", "st09", "st11", "st13", _
                                                 "st15", "st17", "st19", "st21", "st23", "st25", "st27"
        Case 2 'anel 10
            ReDim fileNameListTypeStr(13)
            argStrToMyArray fileNameListTypeStr, "st02", "st04", "st06", "st08", "st10", "st12", "st14", _
                                                 "st16", "st18", "st20", "st22", "st24", "st26", "st28"
        Case 3, 4, 5 'anel 09, 08, 06
            ReDim fileNameListTypeStr(15)
            argStrToMyArray fileNameListTypeStr, "st01", "st03", "st05", "st07", "st09", "st11", "st13", _
                                                 "st15", "st17", "st19", "st21", "st23", "st25", "st27", "st29", "st31"
        Case 6 'anel 04
            ReDim fileNameListTypeStr(21)
            argStrToMyArray fileNameListTypeStr, "st01", "st02", "st03", "st04", "st05", "st06", "st07", "st08", "st09", "st10", _
                                    "st11", "st12", "st13", "st14", "st15", "st16", "st17", "st18", "st19", "st20", "st21", "st22"
                                      
        End Select
        
        For j = 0 To UBound(fileNameListTypeStr)
            
            flagErro = ObjSistemaAquivos.FileExists(".\IR\" & strEquipamentGroup(i) & "\" & fileNameListTypeStr(j) & ".jpg") _
            And ObjSistemaAquivos.FileExists(".\Tratadas\" & strEquipamentGroup(i) & "\" & fileNameListTypeStr(j) & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strEquipamentGroup(i) & "\" & fileNameListTypeStr(j) & ".jpg" & vbCrLf
                contErro = contErro + 1
            End If
            
        Next j
        
    Next i
    
    If (contErro > 0) Then
        checkFilesAndFolders = False
        MsgBox msgErroNoFoundDiretorio
        Exit Function
    End If
    
    checkFilesAndFolders = True
    
End Function


Private Function getAndWriteImgDataST(ByRef strEquipamentGroup As String, ByRef strFileNameList() As String)

    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer

    For i = 0 To UBound(strFileNameList)
        
        nameShapeObj = UCase(strEquipamentGroup) & "_" & UCase(strFileNameList(i))
        fullPathFile = ".\Tratadas\" & strEquipamentGroup & "\" & strFileNameList(i) & ".jpg"
        With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
            .Select
            If (.InlineShapes.Count <> 0) Then
                .InlineShapes.Item(1).Delete
            End If
            .InlineShapes.AddPicture (fullPathFile)
            .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Width
            .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Height
        End With

        fullPathFile = ".\IR\" & strEquipamentGroup & "\" & strFileNameList(i) & ".jpg"

        str = Split(CStr(ObjSistemaAquivos.getFile(fullPathFile).DateLastModified))
        ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = str(0)
        ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = str(1)
        ''ActiveDocument.Shapes(nameShapeObj).GroupItems("Nome").TextFrame.TextRange.Text = UCase(strFileNameList(i))
        
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        DoEvents
    
    Next i
      
End Function
Private Function getAndWriteGrafDataST(ByRef index, ByRef strEquipamentGroup As String, ByRef strFileNameList() As String)
    
    Dim i As Integer, linhaPlan As Integer
    Dim ValueOfColuns As Integer ' linha pra achar a ultima inserida, coluna dos valores de cada lado
    Dim pathWorkbook As String, workBookName As String
    
    Dim nameShapeObj As String
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook
                            'corrigi choose não pode ser < 1
    workBookName = Choose(index + 1, "Gráfico STAVES 13° ANEL.xlsx", "Gráfico STAVES 11° ANEL.xlsx", "Gráfico STAVES 10° ANEL.xlsx", "Gráfico STAVES 09° ANEL.xlsx", _
                                 "Gráfico STAVES 08° ANEL.xlsx", "Gráfico STAVES 06° ANEL.xlsx", "Gráfico STAVES 04° ANEL.xlsx")
    
    pathWorkbook = ActiveDocument.Path & "\Gráfico\" & workBookName
    
    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = False
    
    Dim strCharts() As Variant, strchart As Variant
    
    
    myWorkbook.Sheets(UCase(strEquipamentGroup)).Activate
    
    With myWorkbook.Sheets(UCase(strEquipamentGroup))
        linhaPlan = 5
        appExcel.Cells(linhaPlan, 2).Select
        Do Until appExcel.Cells(linhaPlan, 2).Value = ""
            linhaPlan = linhaPlan + 1
        Loop

        linhaPlan = linhaPlan - 1 ' Volta pra linha interessada
        ValueOfColuns = 7 ' 7 é a primeira coluna com os valores das temperaturas é variavel
        
    '======== Inicio Insere as Temperaturas ===========
        For i = 0 To UBound(strFileNameList)
            nameShapeObj = UCase(strEquipamentGroup) & "_" & UCase(strFileNameList(i))
            With ActiveDocument.Shapes(nameShapeObj).GroupItems("Temp").TextFrame
                .TextRange.Select
                .TextRange.Delete
                .TextRange.Text = "MAX= " & appExcel.Cells(linhaPlan, ValueOfColuns) & "ºC"
                .VerticalAnchor = msoAnchorBottom
            End With

            ValueOfColuns = ValueOfColuns + 1

         Next i
    End With
        '======== Fim Insere as Temperaturas =========
        
        '======== Inicio Insere grafico ===========
    Select Case index
        Case 0 'anel 13
            strCharts = Array("ST-01~11", "ST-15~25")
        Case 1 'anel 11
            strCharts = Array("ST-01~07", "ST-09~15", "ST-17~23", "ST-25~27")
        Case 2 'anel 10
            strCharts = Array("ST-02~08", "ST-10~16", "ST-18~24", "ST-26~28")
        Case 3, 4, 5 'anel 09, 08, 06
            strCharts = Array("ST-01~07", "ST-09~15", "ST-17~23", "ST-25~31")
        Case 6 'anel 04
            strCharts = Array("ST-01~04", "ST-05~08", "ST-09~12", "ST-13~16", "ST-17~20", "ST-21~22")
    End Select
    
    For Each strchart In strCharts
        nameShapeObj = UCase(strEquipamentGroup) & "_" & strchart & "_GRAFICO"
        myWorkbook.Charts(strchart).Activate
        myWorkbook.Charts(strchart).ChartArea.Select
        myWorkbook.Charts(strchart).ChartArea.Copy
        
        With ActiveDocument.Shapes(nameShapeObj).TextFrame.TextRange
            .Select
            If .InlineShapes.Count > 0 Then
            .InlineShapes(1).Delete
            End If
        'Selection.Paste
        Selection.PasteSpecial DataType:=wdPasteBitmap
        End With
        
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
        DoEvents
            
    Next strchart
        '======== Fim Insere grafico =========
    
    Erase strCharts
    
    appExcel.CutCopyMode = False ' excel
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing
    
    
End Function
'*****************************************************************************************
'*****************************************************************************************

Sub mainStaves()

    successFinal = False
    
'    On Error GoTo trata

    If Not (checkFilesAndFolders(equipamentGroup)) Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim fileNameList() As String
    
    
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If
    
     '================= Barra de progresso =========================
    Dim maxProgBar, controlBar As Integer
    maxProgBar = 106 + 28 'numero de imagens + graficos
    StartForm.ProgressBar1.Value = 0
    StartForm.ProgressBar1.Max = maxProgBar
    '================= Barra de progresso =========================
    
    For i = 0 To UBound(equipamentGroup)
        
        Select Case i
        Case 0 'anel 13
            ReDim fileNameList(7)
            argStrToMyArray fileNameList, "st01", "st04", "st08", "st11", "st15", "st18", "st22", "st25"
        Case 1 'anel 11
            ReDim fileNameList(13)
            argStrToMyArray fileNameList, "st01", "st03", "st05", "st07", "st09", "st11", "st13", _
                                          "st15", "st17", "st19", "st21", "st23", "st25", "st27"
        Case 2 'anel 10
            ReDim fileNameList(13)
            argStrToMyArray fileNameList, "st02", "st04", "st06", "st08", "st10", "st12", "st14", _
                                         "st16", "st18", "st20", "st22", "st24", "st26", "st28"
        Case 3, 4, 5 'anel 09, 08, 06
            ReDim fileNameList(15)
            argStrToMyArray fileNameList, "st01", "st03", "st05", "st07", "st09", "st11", "st13", _
                                          "st15", "st17", "st19", "st21", "st23", "st25", "st27", "st29", "st31"
        Case 6 'anel 04
            ReDim fileNameList(21)
            argStrToMyArray fileNameList, "st01", "st02", "st03", "st04", "st05", "st06", "st07", "st08", "st09", "st10", _
                                    "st11", "st12", "st13", "st14", "st15", "st16", "st17", "st18", "st19", "st20", "st21", "st22"
                                      
        End Select
        
        Call getAndWriteImgDataST(equipamentGroup(i), fileNameList)
        Call getAndWriteGrafDataST(i, equipamentGroup(i), fileNameList)
       'Call deleteImgsST(equipamentGroup(i), fileNameList)
    Next i
    
    successFinal = True
    
    Exit Sub
    
'trata:
'   MsgBox "Erro: " & Err.Number & "  " & Err.Description & vbNewLine & _
'           "Fonte: " & Err.Source
'    End
    
End Sub

Sub startStaves()
    
    If WarningTask Then
        MsgBox "Feche os tarefas do excel em execução, Tente novamente"
        Exit Sub
    End If
    
    Call changeLocal
    
    If Not (whatDoc(ActiveDocument.name, "STAVES")) Then Exit Sub
    
    checkIRtratFolders
    
    changeFoldersNamesX
    
    renameImgs
    
    searchChart
    
    StartForm.caption = "Staves"
    StartForm.Show
    
    ActiveDocument.SaveAs2 filename:=CurDir & "\RT-STAVES-AFA 2021-XX"
    
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If

    ActiveDocument.Save
    
End Sub

Private Function deleteImgsST(ByRef strEquipamentGroup As String, ByRef strFileNameList() As String)

    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer

    For i = 0 To UBound(strFileNameList)
    
        nameShapeObj = UCase(strEquipamentGroup) & "_" & UCase(strFileNameList(i))
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
      
End Function

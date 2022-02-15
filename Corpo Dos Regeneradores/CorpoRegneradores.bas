Attribute VB_Name = "CorpoRegneradores"
Option Explicit
'===========================================
'Chama o form RenameImgsDOMOSForm
'===========================================
Private Sub renameImgs()

    Dim response As VbMsgBoxResult
    
    response = MsgBox("Deseja Renomear as imagens? Lógico se vc já renomeou não precisa JUMENTO", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal)
    
    If response = vbYes Then RenameImgsCorpRegForm.Show
    
End Sub
'===========================================
'Identifica os arquivos de grafico variaveis globais workbook
'===========================================
Private Function searchChart(ByVal grafName As String)

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not fs.FileExists(grafName) Then
        MsgBox "ERRO: Verifique o nome do grafico: """ & grafName & """ e se ele esta nesta pasta!!!"
        End
    End If
End Function
Private Function checkFilesAndFolders(ByRef strFileNameList() As String) As Boolean
    
    Dim ObjSistemaAquivos As Object
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim msgErroNoFoundDiretorio As String
    Dim i As Integer ' uso geral
    Dim flagErro As Boolean
    Dim contErro As Integer
    
    
    flagErro = False
    contErro = 0
    msgErroNoFoundDiretorio = "Arquivos não encontrados " & vbCrLf
    
    For i = 0 To UBound(strFileNameList)
            
            flagErro = ObjSistemaAquivos.FileExists(".\IR\" & strFileNameList(i) & ".jpg") _
            And ObjSistemaAquivos.FileExists(".\Tratadas\" & strFileNameList(i) & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strFileNameList(i) & ".jpg" & vbCrLf
                contErro = contErro + 1
            End If
    Next i
     If (contErro > 0) Then
            checkFilesAndFolders = False
            MsgBox msgErroNoFoundDiretorio
            Exit Function
    End If
    
    checkFilesAndFolders = True

End Function
'============================================================
'Coloca as imagens tratadas, data e hora das originais dos HS
'============================================================
Private Function getAndWriteImgDataReg(ByRef strFileNameList() As String)

    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer
    
    For i = 0 To UBound(strFileNameList)
        
         nameShapeObj = UCase(strFileNameList(i))
         fullPathFile = ".\Tratadas\" & strFileNameList(i) & ".jpg"
         With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
             .Select
             If (.InlineShapes.Count <> 0) Then
                 .InlineShapes.Item(1).Delete
             End If
             .InlineShapes.AddPicture (fullPathFile)
             .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Width
             .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Height
         End With

         fullPathFile = ".\IR\" & strFileNameList(i) & ".jpg"
 
         str = Split(CStr(ObjSistemaAquivos.getFile(fullPathFile).DateLastModified))
         ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = str(0)
         ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = str(1)
      
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
        DoEvents
    
    Next i
    
End Function
Private Function getAndWriteGrafDataReg(ByRef strFileNameList() As String)
    Dim i As Integer
    Dim linhaPlan As Integer, ValueOfColuns As Integer
    Dim pathWorkbook As String, workBookName As String
    Dim strchart As String
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook
    
    
    workBookName = "Corpo Regeneradores-2021.xlsx"
    pathWorkbook = ActiveDocument.Path & "\" & workBookName
    
    
    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = False
    
    myWorkbook.Sheets("Regeneradores").Activate
    With myWorkbook.Sheets("Regeneradores")
        linhaPlan = 5
        appExcel.Cells(linhaPlan, 2).Select
        Do Until appExcel.Cells(linhaPlan, 2).Value = ""
            linhaPlan = linhaPlan + 1
        Loop

        linhaPlan = linhaPlan - 1 ' Volta pra linha interessada
        ValueOfColuns = 7 ' 7 é a primeira coluna com os valores das temperaturas é variavel
        
    '======== Inicio Insere as Temperaturas ===========
        For i = 0 To UBound(strFileNameList) ' tirando os queimadores
            If ValueOfColuns = 18 Or ValueOfColuns = 23 Or ValueOfColuns = 48 Then
                ValueOfColuns = ValueOfColuns + 1
            End If
            
            With ActiveDocument.Shapes(UCase(strFileNameList(i))).GroupItems("Temp").TextFrame
                .TextRange.Select
                .TextRange.Delete
                .TextRange.Text = "MAX= " & appExcel.Cells(linhaPlan, ValueOfColuns) & "ºC"
                .VerticalAnchor = msoAnchorBottom
            End With
            
            ValueOfColuns = ValueOfColuns + 1
          
         Next i
        '======== Fim Insere as Temperaturas =========
        
    End With
    
'======== Inicio Insere grafico ===========
    For i = 0 To UBound(strFileNameList) - 8
    
        strchart = strFileNameList(i)
        
        'If strchart = "HS-835" Then GoTo pulahs
        
        myWorkbook.Charts(strchart).Activate
        myWorkbook.Charts(strchart).ChartArea.Select
        myWorkbook.Charts(strchart).ChartArea.Copy
        
        With ActiveDocument.Shapes(UCase(strFileNameList(i)) & "_GRAFICO").TextFrame.TextRange
            .Select
            If .InlineShapes.Count > 0 Then
            .InlineShapes(1).Delete
            End If
        Selection.PasteSpecial DataType:=wdPasteBitmap
        
        End With
        
''pulahs:
        
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
        DoEvents
        
    Next i
        
        '======== Fim Insere grafico =========
        
    appExcel.CutCopyMode = False ' excel
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing
            
End Function

Sub mainCorpReg()

    successFinal = False
    
    On Error GoTo trata

    Dim fileNameList(0 To 54) As String


    argStrToMyArray fileNameList, "HS-001", "HS-002", "HS-003", "HS-004", "HS-736", "HS-737", "HS-769", "HS-775", "HS-789", "HS-802", "HS-803", _
                                  "HS-812", "HS-813", "HS-814", "HS-816", "HS-831", "HS-832", "HS-833", "HS-835", "HS-755", "HS-757", "HS-759", _
                                  "HS-760", "HS-772", "HS-773", "HS-777", "HS-790", "HS-792", "HS-794", "HS-796", "HS-820", "HS-007", "HS-008", _
                                  "HS-011", "HS-740", "HS-742", "HS-745", "HS-778", "HS-808", "HS-810", "HS-822", "HS-827", "HS-830", "HS-834", _
                                  "HS-548", "HS-549", "HS-774", _
                                  "REG1_QUEIMADOR_VS", "REG1_QUEIMADOR_LD", "REG2_QUEIMADOR_VS", "REG2_QUEIMADOR_LD", "REG3_QUEIMADOR_VS", "REG3_QUEIMADOR_LD", "REG4_QUEIMADOR_LE", "REG4_QUEIMADOR_LD"



    If Not (checkFilesAndFolders(fileNameList)) Then
        Exit Sub
    End If
    

    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If

'    ================= Barra de progresso =========================
    Dim maxProgBar
    maxProgBar = (UBound(fileNameList) + 1) * 2 - 8 ' - 8 graficos dos queimadores
    StartForm.ProgressBar1.Value = 0
    StartForm.ProgressBar1.Max = maxProgBar
'    ================= Barra de progresso =========================
    
    
    Call getAndWriteImgDataReg(fileNameList)
    Call getAndWriteGrafDataReg(fileNameList)


   successFinal = True

    Exit Sub

trata:
     MsgBox Err.Description & Err.Number
    'Resume Next

End Sub

'====================================================================
'Chamada pela barra de opçoes no word é a PRINCIPAL que inicia tudo
'====================================================================
Sub startCorpReg()
    
    If WarningTask Then
        MsgBox "Feche os tarefas do excel em execução, Tente novamente"
        Exit Sub
    End If
    
    changeLocal
    
    If Not (whatDoc(ActiveDocument.name, "REGENERADORE")) Then Exit Sub
    
    checkIRtratFolders
    
    If Not (IsEmptyFolder("IR") Or IsEmptyFolder("Tratadas")) Then
        MsgBox "Nenhum Arquivo nas pastas IR ou Tratadas"
        End
    End If
    
    renameImgs
    
    Call searchChart("Corpo Regeneradores-2021.xlsx")
    
    StartForm.caption = "Corpo dos Regeneradores"
    StartForm.Show


    ActiveDocument.SaveAs2 filename:=CurDir & "\RT-CORPO REGENERADORES-AFA 2021-XX"
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If

    ActiveDocument.Save
End Sub
Sub tt()
    changeLocal
    Call searchChart("Corpo Regeneradores-2021.xlsx")
End Sub

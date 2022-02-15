Attribute VB_Name = "Domos"
Option Explicit
Dim equipamentGroup(0 To 4) As String
Private workbookPOS As String
Private workbookHS As String
'===========================================
'Preenche o array global equipamentGroup  e verifica se as pasta estão no local
'============================================
Private Function changeFoldersNamesX()

    Dim fs As Object
    Dim name As Variant
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    argStrToMyArray equipamentGroup, "COWPER1", "COWPER2", "COWPER3", "COWPER4", "HS"
    
    For Each name In equipamentGroup
        If Not (fs.FolderExists("IR\" & name) And fs.FolderExists("Tratadas\" & name)) Then
            MsgBox "Pastas não encontradas!!! Verifique Pastas dos equipamentos e HS", vbCritical
            End
        End If
    Next name
       
End Function
'===========================================
'Chama o form RenameImgsDOMOSForm
'===========================================
Private Sub renameImgs()

    Dim i As Integer
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Deseja Renomear as imagens? Lógico se vc já renomeou não precisa JUMENTO", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal)
    
    If response = vbNo Then Exit Sub
    
    For i = 0 To (UBound(equipamentGroup) - 1) ' são as os 4 cowpers, por isso tira o ultimoidici
        With RenameImgsDOMOSForm
            .Height = 140
            .Width = 309
            .Frame1.Left = 6
            .Frame1.Top = 36
            .LabelEquiGroup.caption = equipamentGroup(i)
'            If i = 0 Then
'                .LabelPOS3.caption = "POS4"
'            End If
        End With
        RenameImgsDOMOSForm.Show
    Next i
    
    With RenameImgsDOMOSForm
        .LabelEquiGroup.caption = equipamentGroup(UBound(equipamentGroup)) ' indice HS
        .Height = 222
        .Width = 375
        .Frame1.Visible = False
        RenameImgsDOMOSForm.Show
    End With
End Sub
'===========================================
'Identifica os arquivos de grafico variariveis globais workbook
'===========================================
Private Function searchChart() As Integer

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")

    workbookPOS = "Grafico Domos-Posiçoes-2021.xlsx"
    workbookHS = "Grafico HS-2021.xlsx"
    
    
    If Not (fs.FileExists(workbookHS) And fs.FileExists(workbookPOS)) Then
        MsgBox "ERRO: Verifique os graficos:" & vbNewLine _
        & "Grafico das Posiçoes: " & workbookPOS _
        & "Gráfico dos HS: " & workbookHS
        searchChart = 0
        End
    End If
    searchChart = 1
End Function
'==========================================================================
'Faz uma nova checagem das pasta e aruivos completos após renomeados
'=========================================================================
Private Function checkFilesAndFolders(ByRef strEquipamentGroup() As String, ByRef strFileNameListPOS() As String, ByRef strFileNameListHS() As String) As Boolean
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
    
    
    For indice = 0 To UBound(strEquipamentGroup)
        flagErro = ((ObjSistemaAquivos.FolderExists(".\IR\" & strEquipamentGroup(indice))) And _
                (ObjSistemaAquivos.FolderExists(".\Tratadas\" & strEquipamentGroup(indice))))

        If (Not flagErro) Then
            msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strEquipamentGroup(indice) & vbCrLf
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
    
    'verifica as posiçoes
    For indice = 0 To (UBound(strEquipamentGroup) - 1)
        For i = 0 To UBound(strFileNameListPOS)
            
            flagErro = ObjSistemaAquivos.FileExists(".\IR\" & strEquipamentGroup(indice) & "\" & strFileNameListPOS(i) & ".jpg") _
            And ObjSistemaAquivos.FileExists(".\Tratadas\" & strEquipamentGroup(indice) & "\" & strFileNameListPOS(i) & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & strEquipamentGroup(indice) & "\" & strFileNameListPOS(i) & ".jpg" & vbCrLf
                contErro = contErro + 1
            End If
            
        Next i
    Next indice
        
    'verifica os HS
    For i = 0 To UBound(strFileNameListHS)
            
            flagErro = ObjSistemaAquivos.FileExists(".\IR\HS\" & strFileNameListHS(i) & ".jpg") _
            And ObjSistemaAquivos.FileExists(".\Tratadas\HS\" & strFileNameListHS(i) & ".jpg")
            
            If (Not flagErro) Then
                msgErroNoFoundDiretorio = msgErroNoFoundDiretorio & "HS\" & strFileNameListHS(i) & ".jpg" & vbCrLf
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
'=====================================================
'Coloca as imagens tratadas, data e hora das originais das POS
'====================================================
Private Function getAndWriteImgDataCowperPOS(ByRef strEquipamentGroup() As String, ByRef strFileNameListPOS() As String)

    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer, j As Integer
        
     For i = 0 To (UBound(strEquipamentGroup) - 1)
     
        For j = 0 To UBound(strFileNameListPOS)
         
            nameShapeObj = UCase(strEquipamentGroup(i)) & "_" & UCase(strFileNameListPOS(j))
            fullPathFile = ".\Tratadas\" & strEquipamentGroup(i) & "\" & strFileNameListPOS(j) & ".jpg"
            With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
                .Select
                If (.InlineShapes.Count <> 0) Then
                    .InlineShapes.Item(1).Delete
                End If
                .InlineShapes.AddPicture (fullPathFile)
                .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Width
                .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Height
            End With
   
            fullPathFile = ".\IR\" & strEquipamentGroup(i) & "\" & strFileNameListPOS(j) & ".jpg"
    
            str = Split(CStr(ObjSistemaAquivos.getFile(fullPathFile).DateLastModified))
            ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = str(0)
            ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = str(1)
            
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
            
            DoEvents
        
        Next j
        
    Next i
   
End Function
'=====================================================
'Coloca os Graficos e temperaturas das POS
'====================================================
Private Function getAndWriteGrafDataCowperPOS(ByRef strEquipamentGroup() As String, ByRef strFileNameListPOS() As String)
    
    Dim j As Integer, i As Integer
    Dim linhaPlan As Integer, ValueOfColuns As Integer
    Dim pathWorkbook As String
    
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook

    pathWorkbook = ActiveDocument.Path & "\" & workbookPOS
    
    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = False
    
    Dim strchart As String
    
    'localiza o ultimo preenchimento na grafico excel
    For i = 0 To (UBound(strEquipamentGroup) - 1) ' sempre tirando o indice HS
        myWorkbook.Sheets(strEquipamentGroup(i)).Activate
        With myWorkbook.Sheets(strEquipamentGroup(i))
            linhaPlan = 5
            appExcel.Cells(linhaPlan, 2).Select
            Do Until appExcel.Cells(linhaPlan, 2).Value = ""
                linhaPlan = linhaPlan + 1
            Loop
    
            linhaPlan = linhaPlan - 1 ' Volta pra linha interessada
            ValueOfColuns = 7 ' 7 é a primeira coluna com os valores das temperaturas é variavel
            
            
        '======== Inicio Insere as Temperaturas ===========
            For j = 0 To (UBound(strFileNameListPOS) - 1) ' tira o pirometro por enquanto, tem que ser manual pois não tem grafico
                With ActiveDocument.Shapes(UCase(strEquipamentGroup(i)) & "_" & UCase(strFileNameListPOS(j))).GroupItems("Temp").TextFrame
                    .TextRange.Select
                    .TextRange.Delete
                    .TextRange.Text = "MAX= " & appExcel.Cells(linhaPlan, ValueOfColuns) & "ºC"
                    .VerticalAnchor = msoAnchorBottom
                End With

                ValueOfColuns = ValueOfColuns + 1

             Next j
            '======== Fim Insere as Temperaturas =========
            
            
            '======== Inicio Insere grafico ===========
            For j = 0 To (UBound(strFileNameListPOS) - 1)
            
                strchart = "COW" & i + 1 & "-" & strFileNameListPOS(j)
                
                myWorkbook.Charts(strchart).Activate
                myWorkbook.Charts(strchart).ChartArea.Select
                myWorkbook.Charts(strchart).ChartArea.Copy
                
                With ActiveDocument.Shapes(UCase(strEquipamentGroup(i)) & "_" & UCase(strFileNameListPOS(j)) & "_GRAFICO").TextFrame.TextRange
                    .Select
                    If .InlineShapes.Count > 0 Then
                    .InlineShapes(1).Delete
                    End If
                'Selection.Paste
                Selection.PasteSpecial DataType:=wdPasteBitmap
                End With
                
                StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
                
                DoEvents
                
            Next j
        '======== Fim Insere grafico =========
        
        End With
    Next i
    
    appExcel.CutCopyMode = False ' excel
   
    'appExcel.Workbooks.Count
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing
    
End Function
'============================================================
'Coloca as imagens tratadas, data e hora das originais dos HS
'============================================================
Private Function getAndWriteImgDataCowperHS(ByRef strEquipamentGroup As String, ByRef strFileNameListHS() As String)
    
    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPathFile As String 'nome e caminho do arquivo pra obter as proriedades e/ou colocar nos shapes
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer
    
    For i = 0 To UBound(strFileNameListHS)
        
         nameShapeObj = UCase(strFileNameListHS(i))
         fullPathFile = ".\Tratadas\" & strEquipamentGroup & "\" & strFileNameListHS(i) & ".jpg"
         With ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").TextFrame.TextRange
             .Select
             If (.InlineShapes.Count <> 0) Then
                 .InlineShapes.Item(1).Delete
             End If
             .InlineShapes.AddPicture (fullPathFile)
             .InlineShapes.Item(1).Width = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Width
             .InlineShapes.Item(1).Height = ActiveDocument.Shapes(nameShapeObj).GroupItems("Img").Height
         End With

         fullPathFile = ".\IR\" & strEquipamentGroup & "\" & strFileNameListHS(i) & ".jpg"
 
         str = Split(CStr(ObjSistemaAquivos.getFile(fullPathFile).DateLastModified))
         ActiveDocument.Shapes(nameShapeObj).GroupItems("Data").TextFrame.TextRange.Text = str(0)
         ActiveDocument.Shapes(nameShapeObj).GroupItems("Hora").TextFrame.TextRange.Text = str(1)
      
        StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        DoEvents
    
    Next i
End Function
'===========================================================
'Coloca as imagens tratadas, data e hora das originais dos HS
'============================================================
Private Function getAndWriteGrafDataCowperHS(ByRef strEquipamentGroup As String, ByRef strFileNameListHS() As String)
    
    Dim j As Integer
    Dim linhaPlan As Integer, ValueOfColuns As Integer
    Dim pathWorkbook As String
    
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook
    
    Dim strchart As String
    
    pathWorkbook = ActiveDocument.Path & "\" & workbookHS
    
    
    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = False
    
        myWorkbook.Sheets("DOMOS").Activate
        With myWorkbook.Sheets("DOMOS")
            linhaPlan = 5
            appExcel.Cells(linhaPlan, 2).Select
            Do Until appExcel.Cells(linhaPlan, 2).Value = ""
                linhaPlan = linhaPlan + 1
            Loop
    
            linhaPlan = linhaPlan - 1 ' Volta pra linha interessada
            ValueOfColuns = 7 ' 7 é a primeira coluna com os valores das temperaturas é variavel
            
            
        '======== Inicio Insere as Temperaturas ===========
            For j = 0 To UBound(strFileNameListHS)
                With ActiveDocument.Shapes(UCase(strFileNameListHS(j))).GroupItems("Temp").TextFrame
                    .TextRange.Select
                    .TextRange.Delete
                    .TextRange.Text = "MAX= " & appExcel.Cells(linhaPlan, ValueOfColuns) & "ºC"
                    .VerticalAnchor = msoAnchorBottom
                End With
                
                ValueOfColuns = ValueOfColuns + 1
              
             Next j
            '======== Fim Insere as Temperaturas =========
            
            
            '======== Inicio Insere grafico ===========
            For j = 0 To UBound(strFileNameListHS)
            
                strchart = strFileNameListHS(j)
                
                myWorkbook.Charts(strchart).Activate
                myWorkbook.Charts(strchart).ChartArea.Select
                myWorkbook.Charts(strchart).ChartArea.Copy
                
                With ActiveDocument.Shapes(UCase(strFileNameListHS(j)) & "_GRAFICO").TextFrame.TextRange
                    .Select
                    If .InlineShapes.Count > 0 Then
                    .InlineShapes(1).Delete
                    End If
                Selection.PasteSpecial DataType:=wdPasteBitmap
                
                End With
                
                StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
                DoEvents
                
            Next j
        '======== Fim Insere grafico =========
        
        End With
    
    appExcel.CutCopyMode = False ' excel
   
    'appExcel.Workbooks.Count
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing
    
End Function

'====================================================================
'Chamada pela barra de opçoes no word é a PRINCIPAL que inicia tudo
'====================================================================
Sub startCowper()
    
    If WarningTask Then
        MsgBox "Feche os tarefas do excel em execução, Tente novamente"
        Exit Sub
    End If
    
    changeLocal
    
    If Not (whatDoc(ActiveDocument.name, "DOMOS")) Then Exit Sub
    
    checkIRtratFolders
    
    changeFoldersNamesX
    
    If searchChart = 1 Then
        StartForm.caption = "Domos"
        StartForm.Show
    End If

    ActiveDocument.SaveAs2 filename:=CurDir & "\RT-DOMOS-AFA 2021-XX"
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If

    ActiveDocument.Save
End Sub

Sub mainCowper()

    successFinal = False
    
    On Error GoTo trata
    
    Dim fileNameListPOS(0 To 3) As String
    Dim fileNameListHS(0 To 20) As String
    Dim i As Integer
    
    argStrToMyArray fileNameListPOS, "POS1", "POS2", "POS3", "PIROMETRO" 'cowper 1 é POS4, mas ta como POS3 pra facilitar
    argStrToMyArray fileNameListHS, "HS-744", "HS-745", "HS-779", "HS-824", "HS-825", "HS-796", "HS-798", "HS-801", "HS-805", "HS-826", _
                                    "HS-748", "HS-802", "HS-807", "HS-820", "HS-823", "HS-799", "HS-808", "HS-822", "HS-827", "HS-828", "HS-829"

    
    If Not (checkFilesAndFolders(equipamentGroup, fileNameListPOS, fileNameListHS)) Then
        Exit Sub
    End If

    If ActiveDocument.ProtectionType <> wdNoProtection Then
        ActiveDocument.Unprotect ("01552375609")
    End If
    
    '================= Barra de progresso =========================
    Dim maxProgBar
    maxProgBar = (UBound(fileNameListPOS) + 1) * UBound(equipamentGroup) * 1.75 + (UBound(fileNameListHS) + 1) * 2
    StartForm.ProgressBar1.Value = 0
    StartForm.ProgressBar1.Max = maxProgBar
    '================= Barra de progresso =========================
    
    'writePOS
    Call getAndWriteImgDataCowperPOS(equipamentGroup, fileNameListPOS)
    Call getAndWriteGrafDataCowperPOS(equipamentGroup, fileNameListPOS)
    
    'write HS
    Call getAndWriteImgDataCowperHS(equipamentGroup(UBound(equipamentGroup)), fileNameListHS)
    Call getAndWriteGrafDataCowperHS(equipamentGroup(UBound(equipamentGroup)), fileNameListHS)
    
'    Call deleteImgsPOS(equipamentGroup, fileNameListPOS)
'    Call deleteImgsHS(equipamentGroup(UBound(equipamentGroup)), fileNameListHS)
    
    successFinal = True
    
    Exit Sub
    
trata:
     MsgBox Err.Description & Err.Number
     End
'    Resume Next

End Sub
Private Function deleteImgsPOS(ByRef strEquipamentGroup() As String, ByRef strFileNameListXX() As String)

    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer, j As Integer
        
    For i = 0 To (UBound(strEquipamentGroup) - 1)
     
        For j = 0 To UBound(strFileNameListXX)
         
            nameShapeObj = UCase(strEquipamentGroup(i)) & "_" & UCase(strFileNameListXX(j))
            
            'Apaga as imagens
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
            
            'Apaga o grafico
            If Not (strFileNameListXX(j) = "PIROMETRO") Then
                With ActiveDocument.Shapes(nameShapeObj & "_GRAFICO").TextFrame.TextRange
                    .Select
                    If (.InlineShapes.Count <> 0) Then
                        .InlineShapes.Item(1).Delete
                    End If
                End With
            End If
            
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
            DoEvents
        Next j
    Next i
    
End Function

Private Function deleteImgsHS(ByRef strEquipamentGroup As String, ByRef strFileNameListXX() As String)
    Dim ObjSistemaAquivos As Object
    
    Set ObjSistemaAquivos = CreateObject("Scripting.FileSystemObject")
    
    Dim nameShapeObj As String 'nome dos shape maiusculos agrupados no documento
    Dim str As Variant ' sera o array derivado da função split da formatação da data e hora
    Dim i As Integer, j As Integer
     
        For j = 0 To UBound(strFileNameListXX)
         
            nameShapeObj = UCase(strFileNameListXX(j))
            
            'Apaga as imagens
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
            
            'Apaga o grafico
            With ActiveDocument.Shapes(nameShapeObj & "_GRAFICO").TextFrame.TextRange
                .Select
                If (.InlineShapes.Count <> 0) Then
                    .InlineShapes.Item(1).Delete
                End If
            End With
            
            StartForm.ProgressBar1.Value = StartForm.ProgressBar1.Value + 1 '=== Atualiza Barra de progress ====
        
            DoEvents
        Next j
    
End Function

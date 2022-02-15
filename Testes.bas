Attribute VB_Name = "Testes"
Option Explicit
Function bordasAndOther(ByVal nomeDosdiretorio As String, ByVal filename_ As String)
    
    Dim nameShapeObj As String
    Dim teste As Integer
    'teste = 0
    
    nameShapeObj = UCase(nomeDosdiretorio) & "_" & UCase(filename_)
    
    Dim myshape As Shape
    
    For Each myshape In ActiveDocument.Shapes(nameShapeObj).GroupItems 'configurarar as bordas
      'myshape.Line.Visible = Not myshape.Line.Visible
      myshape.Line.Visible = msoFalse
    Next myshape

End Function

Sub testtemp()

    Dim txt() As String
    Dim str As String
    Dim test As Variant
    
    MsgBox TypeName(getTemperaturas(".\tempAnel13.txt"))
    txt = Split(getTemperaturas(".\tempAnel13.txt"), Chr(9))
    
    
    With ActiveDocument.Shapes("ANEL13_ST15").GroupItems("Temp").TextFrame.TextRange
        .Select
        .Delete
       ' Selection.Move Unit:=wdLine
        '.Text = vbCrLf & UCase("st 08") & Space(30) & "MAX= " & txt(4) & "ºC"
        
        Selection.Style = ActiveDocument.Styles("Título 1")
        
'        Selection.Move Unit:=wdLine
        Selection.Font.Size = 8
        .InsertAfter (vbCrLf)
        .InsertAfter (UCase("st 11") & Space(30) & "MAX= " & txt(4) & "ºC")
        
        Selection.Move Unit:=wdLine
        Selection.EndOf Unit:=wdLine, Extend:=wdExtend
        Selection.Font.Size = 10
        Selection.Font.Bold = True
        
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
'        Selection.Font.Size = 10
'
        
        
        
'        Selection.Move Unit:=wdLine
'        Selection.StartOf Unit:=wdLine, Extend:=wdExtend
'        Selection.Font.Size = 10
    End With
    
    Erase txt
End Sub

Sub tee()
    
    ChDrive ActiveDocument.Path
    ChDir ActiveDocument.Path
    
    ActiveDocument.Shapes("ANEL13_ST01").GroupItems("Img").TextFrame.TextRange.Select
    
    With ActiveDocument.Shapes("ANEL13_ST01").GroupItems("Img").TextFrame.TextRange
       ' MsgBox .Count
        .Select
        .InlineShapes.Item(1).Delete
        .InlineShapes.AddPicture (".\Tratadas\HS\117.jpg")
        .InlineShapes.Item(1).Width = ActiveDocument.Shapes("ANEL13_ST01").GroupItems("Img").Width
        .InlineShapes.Item(1).Height = ActiveDocument.Shapes("ANEL13_ST01").GroupItems("Img").Height
    End With
     
End Sub

Sub ggg()

    ChDrive ActiveDocument.Path
    ChDir ActiveDocument.Path
    
    
    Dim fileNameList, name As Variant
    Dim nomeDiretorioAneis As Variant
    
    Dim temperaturasList() As String
    
    nomeDiretorioAneis = Array("Saida", "Downleg", "Joelho", "Nariz")
    
    fileNameList = Array("vt01", "vt02", "vt03")
    
    temperaturasList = Split(getTemperaturas(".\tempSaida.txt"), Chr(9))
    
    
    Call getAndWriteInfo2(nomeDiretorioAneis(0), fileNameList(0) & "_LD", temperaturasList(0))
    'Call getAndWriteInfo2(nomeDiretorioAneis(0), fileNameList(0) & "_LE", temperaturasList(0 + 1))
    
    
    'MsgBox ActiveDocument.Shapes("SAIDA_VT01_LD").GroupItems("Temp").TextFrame.TextRange.Text
    
    
End Sub

Sub gdd()
    
    Dim i  As Integer
    Dim nameShapeObj As String, nameshaepeMod As String
    
   
    If (MsgBox(Selection.ShapeRange.Count, vbOKCancel)) = vbCancel Then
        Exit Sub
    End If
    
    For i = 1 To 22
        nameShapeObj = "JOELHO_VT" & Format(i, "00")
        'nameshaepeMod = "VT" & Format(i, "00") & " - JL XX" '-

        With Selection.ShapeRange(nameShapeObj & "_LD").GroupItems("Nome").TextFrame
            '.TextRange.Select
            nameshaepeMod = .TextRange.Text '+
            .TextRange.Delete
            '.TextRange.Text = nameshaepeMod & " / LD"
            .TextRange.Text = Replace(nameshaepeMod, "JL", "NR") '+
            .VerticalAnchor = msoAnchorBottom
        End With
        
        Selection.ShapeRange(nameShapeObj & "_LD").name = "NARIZ_VT" & Format(i, "00") & "_LD"

        
        With Selection.ShapeRange(nameShapeObj & "_LE").GroupItems("Nome").TextFrame
            '.TextRange.Select
            nameshaepeMod = .TextRange.Text '+
            .TextRange.Delete
           ' .TextRange.Text = nameshaepeMod & " / LE"
           .TextRange.Text = Replace(nameshaepeMod, "JL", "NR") '+
            .VerticalAnchor = msoAnchorBottom
        End With
        Selection.ShapeRange(nameShapeObj & "_LE").name = "NARIZ_VT" & Format(i, "00") & "_LE"
        
    Next i
    myResize
End Sub

Sub myResize()

    Dim n As Shape
    Dim str As String
    For Each n In Selection.ShapeRange
    
        'n.GroupItems("Nome").Width = 107.68
        str = n.GroupItems("Nome").TextFrame.TextRange.Text
        MsgBox str
        MsgBox Replace(str, "JL", "NR")
    Next
    
End Sub

Sub hy()
    On Error GoTo trata
    Dim a As Integer
    
    a = "cu"
    MsgBox a
    
    MsgBox "test"
    Exit Sub
    
trata:
    MsgBox Err.Description & vbNewLine & Err.Number
    Resume Next
End Sub

Sub aaa()

    Dim pathWorkbook As String, workBookName As String
    
    workBookName = "GR-CT-AFA 2020-22.xlsm"
    
    pathWorkbook = ActiveDocument.Path & "\" & workBookName
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.workbook

    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(pathWorkbook)
    appExcel.Visible = True
    
    
    'insere a tabela, temperatuara e grafico
    myWorkbook.Sheets("CT-01").Activate
    With myWorkbook.Sheets("CT-01")
        .Range("A5").Select
        Do Until appExcel.ActiveCell = ""
            appExcel.ActiveCell.Offset(1).Select
        Loop
        
        appExcel.ActiveCell.Offset(-1).Select ' volta pra celula com conteudo de fato
        
        
       '======== Inicio Insere as Temperaturas ===========
        With ActiveDocument.Shapes("CT01_CONE_LA_ENCOSTA").GroupItems("Temp").TextFrame
        .TextRange.Select
        .TextRange.Delete
        .TextRange.Text = "MAX= " & appExcel.Cells(appExcel.ActiveCell, 11) & "ºC"
        .VerticalAnchor = msoAnchorBottom

        End With
        '======== Fim Insere as Temperaturas ===========
     
        
        '======== Inicio Insere a tabela ===========
        appExcel.Range("B2", appExcel.Cells(appExcel.ActiveCell.Row, 10)).Select
        
        appExcel.Selection.Copy
        
        With ActiveDocument.Shapes("CT01_TABELA").TextFrame.TextRange
            .Select
            If .Tables.Count > 0 Then
                .Tables(1).Delete
            End If
            Selection.Paste
            .Tables(1).AutoFitBehavior (wdAutoFitWindow)
        End With
        
        '======== Fim Insere a tabela ===========
    End With
    
    
    
    myWorkbook.Charts("Grafico_CT-01").Activate
    myWorkbook.Charts("Grafico_CT-01").ChartArea.Select
    myWorkbook.Charts("Grafico_CT-01").ChartArea.Copy
    With ActiveDocument.Shapes("CT01_GRAFICO").TextFrame.TextRange
            .Select
            If .InlineShapes.Count > 0 Then
                .InlineShapes(1).Delete
            End If
            Selection.Paste
            '.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    End With
    
    appExcel.CutCopyMode = False ' excel
   
    'appExcel.Workbooks.Count
    
    myWorkbook.Close (False) ' descarta alteraçoes na planilha, evita travamento por mensagem clipboard
    
    appExcel.Quit
        
    Set myWorkbook = Nothing
    Set appExcel = Nothing
    
End Sub

Sub aaa2()
''    If ActiveDocument.ProtectionType <> wdNoProtection Then
''        ActiveDocument.Unprotect
''    End If

'    Dim ran As Range
    'ActiveDocument.Bookmarks("CTX1_NAME").Select
    'Selection.TypeText ("11")
'    Set ran = ActiveDocument.Bookmarks("CTX1_NAME").Range
'    ran.Text = "50"
'    ActiveDocument.Bookmarks.Add ("CTX1_NAME"), ran
    
    ActiveDocument.Bookmarks("CTX7").Select
    Selection.Font.Hidden = False
    
'    ActiveDocument.Bookmarks("CTX3").Select
'    Selection.Font.Hidden = False
    
'
'    'ActiveDocument.Shapes("CT02_TABELA").TextFrame.TextRange.Tables.Count
    'MsgBox ActiveDocument.Shapes("CT01_GRAFICO").TextFrame.TextRange.InlineShapes(1).Type
    
''    ActiveDocument.Protect (wdAllowOnlyReading)
    
End Sub
Sub var1()
    Dim filetest As FileDialog
    Set filetest = Application.FileDialog(msoFileDialogFilePicker)
    'filetest.Filters.Add "Pasta de Trabalho Excel", "*.xlsx; *.xlsm", 1
    With filetest
        .AllowMultiSelect = False
        .InitialFileName = ActiveDocument.Path
        .Show
        MsgBox .SelectedItems(1)
    End With
    
    Set filetest = Nothing
End Sub
Sub var2()
    Dim st As String, stre As String
    st = "CT08"
    stre = Right(st, 2)
    MsgBox Len(stre)
End Sub

Sub subgfds()
    Dim str As String
    Dim app As Task
    
    For Each app In Tasks
        str = app.name
        If InStr(1, str, "Excel", vbTextCompare) > 0 Then
            MsgBox app.name
            app.Close
        End If
    Next app
End Sub
Sub language()
    
    changeLocal
    
    Dim fs As Object, ctFolders As Object, ctFolder As Object, irTratFolder As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not (fs.FolderExists("IR") And fs.FolderExists("Tratadas")) Then
        MsgBox "Pasta IR ou Tradadas Não encontrada"
    End If
    
    Set irTratFolder = fs.GetFolder(".\IR") ' aqui basta só um pois a pasta tratadas deve ter as mesmas
    Set ctFolders = irTratFolder.Subfolders
    For Each ctFolder In ctFolders
        MsgBox TypeName(ctFolder.name)
    Next ctFolder
    
End Sub

Public Sub CHES()
Dim idtask As Double
idtask = Shell("C:\Users\mrcordas\Desktop\VBARealatorios\Template - PortaVento\teste.bat", vbNormalFocus)
End Sub

Sub test123()
    changeLocal
    MsgBox IsEmptyFolder("IR")
End Sub

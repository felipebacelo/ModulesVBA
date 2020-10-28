Attribute VB_Name = "ModPreferences"
Option Explicit

'Macro utilizada para desabilitar alguns recursos do Excel deixando-o com uma cara de executável'

Sub TelaMenu()
    
    With Application
        
        .ScreenUpdating = False
        .EnableEvents = False
        
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .Caption = "Programe aqui"
        
        With ActiveWindow
            
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayHeadings = False
            .DisplayWorkbookTabs = False
            .DisplayGridlines = False
            
        End With
        
        .ScreenUpdating = True
        .EnableEvents = True
        
    End With
    
End Sub

'Macro utilizada para retornar os recursos padrões de exibição do Excel'

Sub TelaNormal()
    
    With Application
        
        .ScreenUpdating = False
        .EnableEvents = False
        
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .Caption = Empty
        
        With ActiveWindow
            
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayHeadings = True
            .DisplayWorkbookTabs = True
            .DisplayGridlines = True
            
        End With
        
        .ScreenUpdating = True
        .EnableEvents = True
        
    End With
    
End Sub

'Macro utilizada para criar pasta'

Sub CriarPasta(Pasta)
    
    If VBA.Dir(Pasta, vbDirectory) = "" Then
        
        Shell ("cmd /c mkdir """ & Pasta & """")
        
    End If
    
End Sub

'Macro utilizada para salvar o arquivo em PDF'

Sub SalvarPDF(Plan As String, Caminho As String, NomeArquivo As String)
    
On Error GoTo Erro
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ChDir Caminho
    
    Sheets(Plan).Visible = True
    
    Sheets(Plan).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        Caminho & "\" & NomeArquivo & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
     
    Sheets(Plan).Visible = False
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Exit Sub
Erro:
    
    Sheets(Plan).Visible = False
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

'Macro utilizada para limpar a planilha'

Sub LimparPlanilha(Aba As String, ColI As String, LinI As Integer, ColF As String)
    
    With Sheets(Aba)
        
        If .Range(ColI & LinI).Value = Empty Then Exit Sub
        
        Dim Lin As Integer
        Lin = .Range(ColI & ":" & ColI).Find("", .Range(ColI & LinI)).Row
        
        .Range(ColI & LinI & ":" & ColF & Lin - 1).ClearContents
    
    End With
    
End Sub

'Função utilizada para abrir a caixa de seleção de arquivos'

Function SelecionaArquivo(Optional Filtro As String = "", Optional Extensao As String = "", _
Optional Titulo As String = "", Optional Email As Boolean = False) As String
    
    Dim Caixa As FileDialog
    
    Set Caixa = Application.FileDialog(msoFileDialogOpen)
    
    With Caixa
        
        .InitialView = msoFileDialogViewDetails
        
        .InitialFileName = "C:\"
        
        .AllowMultiSelect = Email
        
        If Filtro <> Empty Then
            .Filters.Clear
            .Filters.Add Filtro, Extensao
        End If
        
    End With
    
    Caixa.Show
    
    SelecionaArquivo = ""
    
    On Error Resume Next
        SelecionaArquivo = Caixa.SelectedItems(1)
    
End Function

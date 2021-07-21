<p align="center">
<a href= "https://img.shields.io/github/repo-size/felipebacelo/ModulesVBA?style=for-the-badge"><img src="https://img.shields.io/github/repo-size/felipebacelo/ModulesVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/github/languages/count/felipebacelo/ModulesVBA?style=for-the-badge"><img src="https://img.shields.io/github/languages/count/felipebacelo/ModulesVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/github/forks/felipebacelo/ModulesVBA?style=for-the-badge"><img src="https://img.shields.io/github/forks/felipebacelo/ModulesVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/bitbucket/pr-raw/felipebacelo/ModulesVBA?style=for-the-badge"><img src="https://img.shields.io/bitbucket/pr-raw/felipebacelo/ModulesVBA?style=for-the-badge"/></a>
<a href= "https://img.shields.io/bitbucket/issues/felipebacelo/ModulesVBA?style=for-the-badge"><img src="https://img.shields.io/bitbucket/issues/felipebacelo/ModulesVBA?style=for-the-badge"/></a>
</p>

# ModulesVBA

Módulos de Formatação e Preferências em VBA Excel.

### Desenvolvimento

Desenvolvido em Microsoft VBA Excel.
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor

### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
* Microsoft Forms 2.0 Object Library

### Compatibilidade

Estes módulos foram desenvolvidos no Excel 2019 (64 bits) e testados no Excel 2016 (64 bits). Sua compatibilidade é garantida para a versão 2016 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento do mesmo.

### Usabilidade

Para utilizar os módulos o usuário deverá:

* Realizar o download do arquivo ZIP: __ModulesVBA__.
* Abrir o Excel.
* Importar através do VBA os arquivos __ModControls.bas__ e __ModPreferences.bas__.
***

### Descrição dos Módulos

#### ModControls

* FormatData (utilizada para formatar datas)
* FormatMoeda (utilizada para formatar moeda)
* FormatCEP (utilizada para formatar CEP)
* FormatCPF (utilizada para formatar CPF)
* FormatCNPJ (utilizada para formatar CNPJ)
* FormatCelular (utilizada para formatar celular)
* FormatTelefone (utilizada para formatar telefone)
***

#### ModPreferences

* TelaMenu (utilizada para desabilitar alguns recursos do Excel deixando-o com uma cara de executável)
* TelaNormal (utilizada para retornar os recursos padrões de exibição do Excel)
* CriarPasta (utilizada para criar pasta)
* SalvarPDF (utilizada para salvar o arquivo em PDF)
* LimparPlanilha (utilizada para limpar a planilha)
* SelecionaArquivo (utilizada para abrir a caixa de seleção de arquivos)
***

### Exemplo de Função Utilizada

```vba
Option Explicit

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
```

***
### Licenças

_MIT License_
_Copyright   ©   2020 Felipe Bacelo Rodrigues_

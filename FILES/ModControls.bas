Attribute VB_Name = "ModControls"
Option Explicit

'Macro utilizada para formatar datas'

Sub FormatData(Controle As Control)
    
    If Controle.Text = Empty Then Exit Sub
    
    If VBA.IsDate(Controle.Text) = False Then
    
        Exit Sub
        
    End If
    
    Controle.Text = VBA.Format(Controle.Text, "dd/mm/yyyy")
    
End Sub

'Macro utilizada para formatar moeda'

Sub FormatMoeda(Controle As Control)
    
    If Controle.Text = Empty Then Exit Sub
    
    If VBA.IsNumeric(Controle.Text) = False Then
    
        Exit Sub
        
    End If
    
    Controle.Text = VBA.Format(Controle.Text, "Currency")
    
End Sub

'Macro utilizada para formatar CEP'

Sub FormatCEP(ByVal TeclaPressionada As MSForms.ReturnInteger, Controle As Control)

    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim CEP As String
            CEP = Controle.Text
            
            Dim x As Integer
            x = VBA.Len(CEP)
            
            If x = 9 Then TeclaPressionada = 0
            
            If x = 5 Then CEP = CEP & "-"
            
            Controle.Text = CEP
            
        Case Else
        
            TeclaPressionada = 0
            
    End Select
    
End Sub

'Macro utilizada para formatar CPF'

Sub FormatCPF(ByVal TeclaPressionada As MSForms.ReturnInteger, Controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim CPF As String
            CPF = Controle.Text
            
            Dim x As Integer
            x = VBA.Len(CPF)
            
            If x = 14 Then TeclaPressionada = 0
            
            If x = 3 Or x = 7 Then CPF = CPF & "."
            If x = 11 Then CPF = CPF & "-"
            
            Controle.Text = CPF
            
        Case Else
        
            TeclaPressionada = 0
            
    End Select
    
End Sub

'Macro utilizada para formatar CNPJ'

Sub FormatCNPJ(ByVal TeclaPressionada As MSForms.ReturnInteger, Controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim CNPJ As String
            CNPJ = Controle.Text
            
            Dim x As Integer
            x = VBA.Len(CNPJ)
            
            If x = 18 Then TeclaPressionada = 0
            
            If x = 2 Or x = 6 Then CNPJ = CNPJ & "."
            If x = 10 Then CNPJ = CNPJ & "/"
            If x = 15 Then CNPJ = CNPJ & "-"
            
            Controle.Text = CNPJ
            
        Case Else
        
            TeclaPressionada = 0
            
    End Select
    
End Sub

'Macro utilizada para formatar celular'

Sub FormatCelular(ByVal TeclaPressionada As MSForms.ReturnInteger, Controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57
            
            Dim Celular As String
            Celular = Controle.Text
            
            Dim x As Integer
            x = VBA.Len(Celular)
            
            If x = 15 Then TeclaPressionada = 0
            
            If x = 0 Then Celular = Celular & "("
            If x = 3 Then Celular = Celular & ") "
            If x = 10 Then Celular = Celular & "-"
            
            Controle.Text = Celular
            
        Case Else
            
            TeclaPressionada = 0
    
    End Select
    
End Sub

'Macro utilizada para formatar telefone'

Sub FormatTelefone(ByVal TeclaPressionada As MSForms.ReturnInteger, Controle As Control)
    
    Select Case TeclaPressionada
        
        Case 48 To 57

            Dim Telefone As String
            Telefone = Controle.Text
            
            Dim x As Integer
            x = VBA.Len(Telefone)
            
            If x = 14 Then TeclaPressionada = 0
            
            If x = 0 Then Telefone = Telefone & "("
            If x = 3 Then Telefone = Telefone & ") "
            If x = 9 Then Telefone = Telefone & "-"
            
            Controle.Text = Telefone
            
        Case Else
        
            TeclaPressionada = 0
    
    End Select
    
End Sub

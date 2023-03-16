Attribute VB_Name = "Módulo1"
Sub arregla_datos()

Dim UF As Long
Dim FILA As Long

'ULTIMA CELDA DINAMICA
UF = Worksheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row

'CREACIÓN DE COLUMNA AUXILIAR
Cells(1, 79).Value = "AUX_COLUMN"

'NUMERACIÓN ORDENADA DE REPETIDOS
For FILA = 2 To UF

'CREACIÓN DE LA CONDICIÓN INICIAL (SOLO SE NUMERA LAS ACTAS QUE SU CUENTA ES MAYOR QUE 1)
datos1 = Worksheets("Hoja1").Cells(FILA, 1)
suma = Application.WorksheetFunction.CountIf(Worksheets("Hoja1").Range("A2:A" & UF), datos1)
  
    If suma > 1 Then
    
        If Cells(FILA, 70).Value Like "*FOB*" Or Cells(FILA, 70).Value Like "*VALOR*" _
                                              Or Cells(FILA, 70).Value Like "*VALO*" _
                                              Or Cells(FILA, 70).Value Like "*US $*" _
                                              Or Cells(FILA, 70).Value Like "*DECLARADO*" _
                                              Or Cells(FILA, 70).Value Like "*INDICADO*" _
                                              Or Cells(FILA, 70).Value Like "*$*" Then
                                              
            Cells(FILA, 79).Value = "FOB"
        
        End If
        
    End If
    

Next

End Sub

Sub arregla_datos2()

Dim UF As Long
Dim FILA As Long

Call arregla_datos
'ULTIMA CELDA DINAMICA
UF = Worksheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row

'CREACIÓN DE COLUMNA AUXILIAR
Cells(1, 80).Value = "AUX_COLUMN(1)"

'NUMERACIÓN ORDENADA DE REPETIDOS
For FILA = 2 To UF

'CREACIÓN DE LA CONDICIÓN INICIAL (SOLO SE NUMERA LAS ACTAS QUE SU CUENTA ES MAYOR QUE 1)
datos1 = Worksheets("Hoja1").Cells(FILA, 1)
suma = Application.WorksheetFunction.CountIf(Worksheets("Hoja1").Range("A2:A" & UF), datos1)
    
'NUMERACIÓN (AÑADIR MAS SI HAY MAS DE 8 REPES)

    
'11111111111111111111111111111111111111111111111111111111111
    If suma = 1 Then
    
    Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value)
    
    End If
    
    
    
    
    
'22222222222222222222222222222222222222222222222222222222222

    If suma = 2 And Cells(FILA, 1) = Cells(FILA + 1, 1) Then
        
            Cells(FILA, 80).Value = Trim(Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
        
    End If
    
    
    
    
    
 '33333333333333333333333333333333333333333333333333333333333
 
    If suma = 3 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                And Cells(FILA + 1, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
    
    Else
    
        If suma = 3 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
    
    
        Else
        
             If suma = 3 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
    
             End If
             
        End If
        
    End If
    
    
    
    
    
 '4444444444444444444444444444444444444444444444444444444444444
 
    If suma = 4 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                And Cells(FILA + 2, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
    
    Else
    
        If suma = 4 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 1, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
    
        
        Else
        
            If suma = 4 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
        
            Else
            
                If suma = 4 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                        And Cells(FILA + 3, 79).Value = "FOB" Then
                    
                    Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
                                        
                End If
                
            End If
            
        End If
        
    End If
    
    
    
    
    
 '5555555555555555555555555555555555555555555555555555555555555555
 
    If suma = 5 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                And Cells(FILA + 3, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
    
    Else
    
        If suma = 5 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 2, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
    
        
        Else
        
            If suma = 5 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                        And Cells(FILA + 1, 79).Value = "FOB" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
        
            Else
            
                If suma = 5 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
        
                Else
                
                    If suma = 5 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
                
                    End If
                    
                End If
                
            End If
            
        End If
        
    End If
    
    
    
    
    
 '666666666666666666666666666666666666666666666666666666666666666666666
    
 If suma = 6 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
             And Cells(FILA + 5, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value)
    
    Else
    
        If suma = 6 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                    And Cells(FILA + 4, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value)
    
        
        Else
        
             If suma = 6 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                         And Cells(FILA + 3, 79).Value = "FOB" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
        
            Else
            
                If suma = 6 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                            And Cells(FILA + 2, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
        
                Else
                
                    If suma = 6 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                And Cells(FILA + 1, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                
                    Else
                    
                        If suma = 6 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                
                        End If
                    
                    End If
                    
                End If
                
            End If
            
        End If
        
 End If
    
    
    
 '7777777777777777777777777777777777777777777777777777777777777777777777777777777
 
 If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
             And Cells(FILA + 6, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
    
    Else
    
        If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                    And Cells(FILA + 5, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
    
        
        Else
        
             If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                         And Cells(FILA + 4, 79).Value = "FOB" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
        
            Else
            
                If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                            And Cells(FILA + 3, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
        
                Else
                
                    If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                And Cells(FILA + 2, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                
                    Else
                    
                        If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                    And Cells(FILA + 1, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                
                        Else
                    
                            If suma = 7 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                    And Cells(FILA, 79).Value = "FOB" Then
        
                  
                              Cells(FILA, 80).Value = Trim(Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                
                           End If
                        
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If
 
 
 
 
 
 '8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
 
 If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
             And Cells(FILA + 7, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
    
    Else
    
        If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                    And Cells(FILA + 6, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
    
        
        Else
        
             If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                         And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                         And Cells(FILA + 5, 79).Value = "FOB" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
        
            Else
            
                If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                            And Cells(FILA + 4, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
        
                Else
                
                    If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                And Cells(FILA + 3, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
                
                    Else
                    
                        If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                    And Cells(FILA + 2, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                
                        Else
                    
                            If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                        And Cells(FILA + 1, 79).Value = "FOB" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 8 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                            And Cells(FILA, 79).Value = "FOB" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If
    
    
    
 '999999999999999999999999999999999999999999999999999999999999999999999999999999999999
 
 
 
 
 If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
             And Cells(FILA + 8, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value)
    
    Else
    
        If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                    And Cells(FILA + 7, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
    
        
        Else
        
             If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                         And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                         And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                         And Cells(FILA + 6, 79).Value = "FOB" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
        
            Else
            
                If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                            And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                            And Cells(FILA + 5, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
        
                Else
                
                    If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                And Cells(FILA + 4, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
                
                    Else
                    
                        If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                    And Cells(FILA + 3, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
                
                        Else
                    
                            If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                        And Cells(FILA + 2, 79).Value = "FOB" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                            And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                            And Cells(FILA + 1, 79).Value = "FOB" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                    
                                Else
                    
                                If suma = 9 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                            And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                            And Cells(FILA, 79).Value = "FOB" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                    
                                End If
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If




'10101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010



 If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
              And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
              And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
              And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
              And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
              And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
              And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
              And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
              And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
              And Cells(FILA + 9, 79).Value = "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value)
    
    Else
    
        If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                     And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                     And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                     And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                     And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                     And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                     And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                     And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                     And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                     And Cells(FILA + 8, 79).Value = "FOB" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value)
    
        
        Else
        
             If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                          And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                          And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                          And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                          And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                          And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                          And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                          And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                          And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                          And Cells(FILA + 7, 79).Value = "FOB" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
        
            Else
            
                If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                             And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                             And Cells(FILA + 6, 79).Value = "FOB" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
        
                Else
                
                    If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                 And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                 And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                 And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                 And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                 And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                 And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                 And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                 And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                 And Cells(FILA + 5, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
                
                    Else
                    
                        If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                     And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                     And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                     And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                     And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                     And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                     And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                     And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                     And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                     And Cells(FILA + 4, 79).Value = "FOB" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
                
                        Else
                    
                            If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                         And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                         And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                         And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                         And Cells(FILA + 3, 79).Value = "FOB" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                             And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                             And Cells(FILA + 2, 79).Value = "FOB" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                    
                                Else
                    
                                   If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 1, 79).Value = "FOB" Then
                                                
                      
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                    
                                
                                
                                    Else
                    
                                      If suma = 10 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                   And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                   And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                   And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                   And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                   And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                   And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                   And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                   And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                   And Cells(FILA, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                     
                     
                                   End If
                                
                                End If
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If
    
Next

rezagados
arregla_datos4
arregla_datos5
arregla_datos3
EviarHojaEmail

Worksheets("Hoja1").Range("CA1:CB" & UF).ClearContents



End Sub

Private Sub arregla_datos3()

Dim UF As Long
Dim UF2 As Long
Dim FILA As Long
Dim exist As Boolean
Dim AEV As Range
Dim BEV As Range
Dim EEV As Range
Dim FEV As Range
Dim HEV As Range
Dim AIEV As Range
Dim BJEV As Range
Dim BKEV As Range
Dim BOEV As Range
Dim BREV As Range
Dim BTEV As Range
Dim BZEV As Range
Dim UNION As Range
Dim RESULETO As Range



Set AEV = Worksheets("Hoja1").Range("A:A")
Set BEV = Worksheets("Hoja1").Range("B:B")
Set FEV = Worksheets("Hoja1").Range("F:F")
Set EEV = Worksheets("Hoja1").Range("E:E")
Set HEV = Worksheets("Hoja1").Range("H:H")
Set AIEV = Worksheets("Hoja1").Range("AI:AI")
Set BJEV = Worksheets("Hoja1").Range("BJ:BJ")
Set BKEV = Worksheets("Hoja1").Range("BK:BK")
Set BOEV = Worksheets("Hoja1").Range("BO:BO")
Set BREV = Worksheets("Hoja1").Range("BR:BR")
Set BTEV = Worksheets("Hoja1").Range("BT:BT")
Set BZEV = Worksheets("Hoja1").Range("BZ:BZ")
Set RESUELTO = Worksheets("Hoja1").Range("CB:CB")

Set UNION = Application.UNION(AEV, BEV, EEV, FEV, HEV, AIEV, BJEV, BKEV, BOEV, BREV, BTEV, BZEV, RESUELTO)

UF = Worksheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row
UNION.Copy

On Error Resume Next
exist = Worksheets("HECHO").Name <> ""

If Not exist Then
    Worksheets.Add(after:=Worksheets("Hoja1")).Name = "HECHO"
    Worksheets("HECHO").Range("A1").PasteSpecial
    Worksheets("HECHO").Range("A1:M" & UF).RemoveDuplicates Columns:=1, Header:=xlYes
    
    UF2 = Worksheets("HECHO").Range("A" & Rows.Count).End(xlUp).Row
    Worksheets("HECHO").Range("M2:M" & UF2).Select
    Selection.Copy
    Worksheets("HECHO").Range("J2:J" & UF2).PasteSpecial
    Worksheets("HECHO").Range("M1:M" & UF2).ClearContents
    Worksheets(2).Cells(1, 1) = "AÑO/ACTA TRASLADO"
    Worksheets(2).Cells(1, 2) = "DEP"
    Worksheets(2).Cells(1, 3) = "TPESO_BRUT"
    Worksheets(2).Cells(1, 4) = "AVISO/AÑO"
    Worksheets(2).Cells(1, 5) = "USUARIO"
    Worksheets(2).Cells(1, 6) = "FECH_MODSI"
    Worksheets(2).Cells(1, 7) = "FEMOD"
    Worksheets(2).Cells(1, 8) = "CODIMOD"
    Worksheets(2).Cells(1, 9) = "ITEM"
    Worksheets(2).Cells(1, 10) = "DESCRIPCION DE LA MERCADERIA"
    Worksheets(2).Cells(1, 11) = "NRO_ORDEN"
    Worksheets(2).Cells(1, 12) = "DPO"
    

Else

    Worksheets.Add(after:=Worksheets("Hoja1")).Name = "HECHO" & Worksheets.Count
    Worksheets(2).Range("A1").PasteSpecial
    Worksheets(2).Range("A1:M" & UF).RemoveDuplicates Columns:=1, Header:=xlYes
    
    UF2 = Worksheets(2).Range("A" & Rows.Count).End(xlUp).Row
    Worksheets(2).Range("M2:M" & UF2).Select
    Selection.Copy
    Worksheets(2).Range("J2:J" & UF2).PasteSpecial
    Worksheets(2).Range("M1:M" & UF2).ClearContents
    Worksheets(2).Cells(1, 1) = "AÑO/ACTA TRASLADO"
    Worksheets(2).Cells(1, 2) = "DEP"
    Worksheets(2).Cells(1, 3) = "TPESO_BRUT"
    Worksheets(2).Cells(1, 4) = "AVISO/AÑO"
    Worksheets(2).Cells(1, 5) = "USUARIO"
    Worksheets(2).Cells(1, 6) = "FECH_MODSI"
    Worksheets(2).Cells(1, 7) = "FEMOD"
    Worksheets(2).Cells(1, 8) = "CODIMOD"
    Worksheets(2).Cells(1, 9) = "ITEM"
    Worksheets(2).Cells(1, 10) = "DESCRIPCION DE LA MERCADERIA"
    Worksheets(2).Cells(1, 11) = "NRO_ORDEN"
    Worksheets(2).Cells(1, 12) = "DPO"
    
   
    
End If



End Sub


Private Sub arregla_datos4()


Dim UF As Long
Dim FILA As Long

'ULTIMA CELDA DINAMICA
UF = Worksheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row

'CREACIÓN DE COLUMNA AUXILIAR
Cells(1, 80).Value = "AUX_COLUMN(1)"

'NUMERACIÓN ORDENADA DE REPETIDOS
For FILA = 2 To UF

'CREACIÓN DE LA CONDICIÓN INICIAL (SOLO SE NUMERA LAS ACTAS QUE SU CUENTA ES MAYOR QUE 1)
datos1 = Worksheets("Hoja1").Cells(FILA, 1)
suma = Application.WorksheetFunction.CountIf(Worksheets("Hoja1").Range("A2:A" & UF), datos1)
    
'NUMERACIÓN (AÑADIR MAS SI HAY MAS DE 8 REPES)


'11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11_11
 
 
 
 If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
              And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
              And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
              And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
              And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
              And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
              And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
              And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
              And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
              And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
              And Cells(FILA + 10, 79).Value Like "*FOB*" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value)
    
    Else
    
        If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                     And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                     And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                     And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                     And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                     And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                     And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                     And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                     And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                     And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                     And Cells(FILA + 9, 79).Value Like "*FOB*" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value)
    
        
        Else
        
             If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                          And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                          And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                          And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                          And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                          And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                          And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                          And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                          And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                          And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                          And Cells(FILA + 8, 79).Value Like "*FOB*" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value)
        
            Else
            
                If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                             And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                             And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                             And Cells(FILA + 7, 79).Value Like "*FOB*" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
        
                Else
                
                    If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                 And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                 And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                 And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                 And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                 And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                 And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                 And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                 And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                 And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                 And Cells(FILA + 6, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
                
                    Else
                    
                        If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                     And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                     And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                     And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                     And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                     And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                     And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                     And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                     And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                     And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                     And Cells(FILA + 5, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
                
                        Else
                    
                            If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                         And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                         And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                         And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                         And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                         And Cells(FILA + 4, 79).Value Like "*FOB*" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                             And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                             And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                             And Cells(FILA + 3, 79).Value Like "*FOB*" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
                    
                                Else
                    
                                   If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                And Cells(FILA + 2, 79).Value Like "*FOB*" Then
                                                
                      
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                    
                                
                                
                                    Else
                    
                                      If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                   And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                   And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                   And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                   And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                   And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                   And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                   And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                   And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                   And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                   And Cells(FILA + 1, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                     
                                     Else
                    
                                      If suma = 11 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                   And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                   And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                   And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                   And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                   And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                   And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                   And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                   And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                   And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                   And Cells(FILA, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                     
                                    End If
                     
                                   End If
                                
                                End If
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If



'12121212121212121212121222121212121212121212121212121212121212121212121212121212121212121212121212121212

 If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
              And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
              And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
              And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
              And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
              And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
              And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
              And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
              And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
              And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
              And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
              And Cells(FILA + 11, 79).Value Like "*FOB*" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 11, 70).Value)
    
    Else
    
        If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                     And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                     And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                     And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                     And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                     And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                     And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                     And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                     And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                     And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                     And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                     And Cells(FILA + 10, 79).Value Like "*FOB*" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value)
    
        
        Else
        
             If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                         And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                         And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                         And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                         And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                         And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                         And Cells(FILA + 9, 79).Value Like "*FOB*" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70))
        
            Else
            
                If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                             And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                             And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                             And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                             And Cells(FILA + 8, 79).Value Like "*FOB*" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value)
        
                Else
                
                    If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                 And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                 And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                 And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                 And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                 And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                 And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                 And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                 And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                 And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                 And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                 And Cells(FILA + 7, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
                
                    Else
                    
                        If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                     And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                     And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                     And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                     And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                     And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                     And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                     And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                     And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                     And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                     And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                     And Cells(FILA + 6, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
                
                        Else
                    
                            If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                         And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                         And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                         And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                         And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                         And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                         And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                         And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                         And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                         And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                         And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                         And Cells(FILA + 5, 79).Value Like "*FOB*" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                             And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                             And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                             And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                             And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                             And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                             And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                             And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                             And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                             And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                             And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                             And Cells(FILA + 4, 79).Value Like "*FOB*" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
                    
                                Else
                    
                                   If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                And Cells(FILA + 3, 79).Value Like "*FOB*" Then
                                                
                      
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70))
                    
                                
                                
                                    Else
                    
                                      If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                   And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                   And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                   And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                   And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                   And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                   And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                   And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                   And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                   And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                   And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                   And Cells(FILA + 2, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                     
                                     Else
                    
                                      If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                   And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                   And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                   And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                   And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                   And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                   And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                   And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                   And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                   And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                   And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                   And Cells(FILA + 1, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                     
                     
                                      Else
                    
                                       If suma = 12 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                   And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                   And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                   And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                   And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                   And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                   And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                   And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                   And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                   And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                   And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                   And Cells(FILA, 79).Value Like "*FOB*" Then
                                                  
                        
                                          Cells(FILA, 80).Value = Trim(Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                     
                                
                                        End If
                                        
                                     End If
                     
                                   End If
                                
                                End If
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If


'1313131313131313131313131313131313131313131313131313131313131313131313131313131313131313131313131313131313131313




Next

End Sub


Private Sub arregla_datos5()


Dim UF As Long
Dim FILA As Long

'ULTIMA CELDA DINAMICA
UF = Worksheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row

'CREACIÓN DE COLUMNA AUXILIAR
Cells(1, 80).Value = "AUX_COLUMN(1)"

'NUMERACIÓN ORDENADA DE REPETIDOS
For FILA = 2 To UF

'CREACIÓN DE LA CONDICIÓN INICIAL (SOLO SE NUMERA LAS ACTAS QUE SU CUENTA ES MAYOR QUE 1)
datos1 = Worksheets("Hoja1").Cells(FILA, 1)
suma = Application.WorksheetFunction.CountIf(Worksheets("Hoja1").Range("A2:A" & UF), datos1)

If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
              And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
              And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
              And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
              And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
              And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
              And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
              And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
              And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
              And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
              And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
              And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
              And Cells(FILA + 12, 79).Value Like "*FOB*" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 12, 70).Value)
    
    Else
    
        If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                    And Cells(FILA + 11, 79).Value Like "*FOB*" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 11, 70).Value)
    
        
        Else
        
             If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                        And Cells(FILA + 10, 79).Value Like "*FOB*" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70) & " " & Cells(FILA + 10, 70).Value)
        
            Else
            
                If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                            And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                            And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                            And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                            And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                            And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                            And Cells(FILA + 9, 79).Value Like "*FOB*" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value)
        
                Else
                
                    If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                And Cells(FILA + 8, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value)
                
                    Else
                    
                        If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                    And Cells(FILA + 7, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
                
                        Else
                    
                            If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                        And Cells(FILA + 6, 79).Value Like "*FOB*" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                And Cells(FILA + 5, 79).Value Like "*FOB*" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
                    
                                Else
                    
                                   If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                And Cells(FILA + 4, 79).Value Like "*FOB*" Then
                                                
                      
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value)
                    
                                
                                
                                    Else
                    
                                      If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                    And Cells(FILA + 3, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value)
                     
                                     Else
                    
                                      If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                    And Cells(FILA + 2, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                     
                     
                                      Else
                    
                                       If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                    And Cells(FILA + 1, 79).Value Like "*FOB*" Then
                                                  
                        
                                          Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                     
                                        Else
                    
                                           If suma = 13 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                        And Cells(FILA, 79).Value Like "*FOB*" Then
                                                      
                            
                                              Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                                    
                                
                                            End If
                                                                         
                                        End If
                                        
                                     End If
                     
                                   End If
                                
                                End If
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If



'1414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141441141414141414141414141414141414141414141414141


If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
              And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
              And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
              And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
              And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
              And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
              And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
              And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
              And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
              And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
              And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
              And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
              And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
              And Cells(FILA + 13, 79).Value Like "*FOB*" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70) & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 13, 70).Value)
    
    Else
    
        If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                    And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                    And Cells(FILA + 12, 79).Value Like "*FOB*" Then
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 12, 70).Value)
    
        
        Else
        
             If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                        And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                        And Cells(FILA + 11, 79).Value Like "*FOB*" Then
                
                    Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70) & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 11, 70).Value)
        
            Else
            
                If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                            And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                            And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                            And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                            And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                            And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                            And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                            And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                            And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                            And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                            And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                            And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                            And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                            And Cells(FILA + 10, 79).Value Like "*FOB*" Then
        
                  
                        Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 10, 70).Value)
        
                Else
                
                    If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                And Cells(FILA + 9, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 12, 70).Value)
                
                    Else
                    
                        If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                    And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                    And Cells(FILA + 8, 79).Value Like "*FOB*" Then
        
                  
                           Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 8, 70).Value)
                
                        Else
                    
                            If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                        And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                        And Cells(FILA + 7, 79).Value Like "*FOB*" Then
        
                  
                               Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 7, 70).Value)
                
                                                   
                           Else
                    
                                If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                And Cells(FILA + 6, 79).Value Like "*FOB*" Then
            
                      
                                   Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 6, 70).Value)
                    
                                Else
                    
                                   If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                And Cells(FILA + 5, 79).Value Like "*FOB*" Then
                                                
                      
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70) & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 5, 70).Value)
                    
                                
                                
                                    Else
                    
                                      If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                    And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                    And Cells(FILA + 4, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 4, 70).Value)
                     
                                     Else
                    
                                      If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                    And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                    And Cells(FILA + 3, 79).Value Like "*FOB*" Then
                                                  
                        
                                        Cells(FILA, 80).Value = Trim(Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                     
                     
                                      Else
                    
                                       If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                    And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                    And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                    And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                    And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                    And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                    And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                    And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                    And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                    And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                    And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                    And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                    And Cells(FILA + 2, 79).Value Like "*FOB*" Then
                                                  
                        
                                          Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA + 2, 70).Value)
                     
                                        Else
                    
                                           If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                        And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                        And Cells(FILA + 1, 79).Value Like "*FOB*" Then
                                                      
                            
                                              Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA, 70).Value & " " & Cells(FILA + 1, 70).Value)
                                    
                                
                                            Else
                                            
                                                If suma = 14 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                                                        And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                                                        And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                                                        And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                                                        And Cells(FILA + 4, 1).Value = Cells(FILA + 5, 1).Value _
                                                        And Cells(FILA + 5, 1).Value = Cells(FILA + 6, 1).Value _
                                                        And Cells(FILA + 6, 1).Value = Cells(FILA + 7, 1).Value _
                                                        And Cells(FILA + 7, 1).Value = Cells(FILA + 8, 1).Value _
                                                        And Cells(FILA + 8, 1).Value = Cells(FILA + 9, 1).Value _
                                                        And Cells(FILA + 9, 1).Value = Cells(FILA + 10, 1).Value _
                                                        And Cells(FILA + 10, 1).Value = Cells(FILA + 11, 1).Value _
                                                        And Cells(FILA + 11, 1).Value = Cells(FILA + 12, 1).Value _
                                                        And Cells(FILA + 12, 1).Value = Cells(FILA + 13, 1).Value _
                                                        And Cells(FILA, 79).Value Like "*FOB*" Then
                                                      
                            
                                                Cells(FILA, 80).Value = Trim(Cells(FILA + 13, 70).Value & " " & Cells(FILA + 12, 70).Value & " " & Cells(FILA + 11, 70).Value & " " & Cells(FILA + 10, 70).Value & " " & Cells(FILA + 9, 70).Value & " " & Cells(FILA + 8, 70).Value & " " & Cells(FILA + 7, 70).Value & " " & Cells(FILA + 6, 70).Value & " " & Cells(FILA + 5, 70).Value & " " & Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
                                    
                                              End If
                                
                                            End If
                                                                         
                                        End If
                                        
                                     End If
                     
                                   End If
                                
                                End If
                    
                              End If
                                                                             
                           End If
                                                   
                        End If
                    
                    End If
                    
                End If

            End If
            
        End If
        
 End If




Next


End Sub


Private Sub rezagados()

Dim UF As Long
Dim FILA As Long

'ULTIMA CELDA DINAMICA
UF = Worksheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row

'CREACIÓN DE COLUMNA AUXILIAR
Cells(1, 80).Value = "AUX_COLUMN(1)"

'NUMERACIÓN ORDENADA DE REPETIDOS
For FILA = 2 To UF

'CREACIÓN DE LA CONDICIÓN INICIAL (SOLO SE NUMERA LAS ACTAS QUE SU CUENTA ES MAYOR QUE 1)
datos1 = Worksheets("Hoja1").Cells(FILA, 1)
suma = Application.WorksheetFunction.CountIf(Worksheets("Hoja1").Range("A2:A" & UF), datos1)

    If suma = 3 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                    And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                    And Cells(FILA + 1, 79).Value <> "FOB" _
                    And Cells(FILA + 2, 79).Value <> "FOB" Then
                   
                
                Cells(FILA, 80).Value = Trim(Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)
        
    End If



If suma = 4 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                And Cells(FILA + 1, 79).Value <> "FOB" _
                And Cells(FILA + 2, 79).Value <> "FOB" _
                And Cells(FILA + 3, 79).Value <> "FOB" Then
            
            Cells(FILA, 80).Value = Trim(Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)

End If


If suma = 5 And Cells(FILA, 1).Value = Cells(FILA + 1, 1).Value _
                And Cells(FILA + 1, 1).Value = Cells(FILA + 2, 1).Value _
                And Cells(FILA + 2, 1).Value = Cells(FILA + 3, 1).Value _
                And Cells(FILA + 3, 1).Value = Cells(FILA + 4, 1).Value _
                And Cells(FILA + 1, 79).Value <> "FOB" _
                And Cells(FILA + 2, 79).Value <> "FOB" _
                And Cells(FILA + 3, 79).Value <> "FOB" _
                And Cells(FILA + 4, 79).Value <> "FOB" Then

            
            Cells(FILA, 80).Value = Trim(Cells(FILA + 4, 70).Value & " " & Cells(FILA + 3, 70).Value & " " & Cells(FILA + 2, 70).Value & " " & Cells(FILA + 1, 70).Value & " " & Cells(FILA, 70).Value)

End If

Next
End Sub

Private Sub EXTRAE_DATA()

Dim LIBRO_ACTAS As Workbook
Dim LIBRO_ARREGLA_ACTAS As Workbook
Dim RUTA As String
Dim UF3

Set LIBRO_ARREGLA_ACTAS = ThisWorkbook
RUTA = Application.GetOpenFilename(Title:="Escoja el libro Excel donde se encuentra la data obtenida del Toad For Oracle (Ranita) sin haberle aplicado cambios")

If RUTA = "FALSE" Then
    Exit Sub
End If

Set LIBRO_ACTAS = Workbooks.Open(RUTA)

UF3 = LIBRO_ACTAS.Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row

LIBRO_ACTAS.Worksheets(1).Cells.Copy Destination:=LIBRO_ARREGLA_ACTAS.Worksheets(1).Range("A1")

LIBRO_ACTAS.Close

End Sub

Private Sub borrar_contenido()
Worksheets("Hoja1").Range("A:BZ").ClearContents

End Sub


'
Private Sub EviarHojaEmail()
'
Dim NombreArchivo As String
Dim RutaTemporal As String
Dim Mensaje As String
    '
    On Error Resume Next
    '
    Mensaje = "Coloca el nombre con el que enviarás las actas de traslado ordenadas para el encargado de enviarselas a la OSA (DEBES TENER INICIADA TU SESION EN OUTLOOK). Verifica rapidamente que el ordenamiento sea correcto; en caso de no serlo, usa formulas anidadas. Contraseña del proyecto: Sunat123456"
    NombreArchivo = InputBox(Mensaje, "DIVISIÓN DE ENVIOS POSTALES")
    '
    If NombreArchivo = "" Then NombreArchivo = ActiveSheet.Name
    '
    RutaTemporal = Environ("temp") & "\"
    NombreArchivo = RutaTemporal & NombreArchivo & ".xlsx"

    ActiveWorkbook.ActiveSheet.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs NombreArchivo
    Application.DisplayAlerts = True
    CommandBars.ExecuteMso ("FileSendAsAttachment")
    ActiveWorkbook.Close False
    Kill NombreArchivo
    '
    On Error GoTo 0
    '
End Sub



Sub NPSCompleto()
'Variables
Dim contadorCursos As Integer
Dim rango As String
Dim fila As Integer
Dim vacio As Boolean
Dim CourseID() As String
Dim ID As String
Dim ContadorFor As Integer
Dim programa, curso, idioma, trimestre, trimestre2 As String
Dim sumaPromoter, sumaDetractor, sumaNeutral, sumaTotal As Integer
trimestre = ""

'Contar Cursos
contadorCursos = 0
fila = 18
rango = "B" + Trim(Str(fila))
vacio = IsEmpty(Range(rango))

While vacio = False
contadorCursos = contadorCursos + 1
fila = fila + 1
rango = "B" + Trim(Str(fila))
vacio = IsEmpty(Range(rango))
Wend
'Fin de Contar Cursos

'For para recorrer cada CourseID

For ContadorFor = 0 To contadorCursos - 1 Step 1

Sheets("Presentación").Select
Sheets("Presentación").Activate
rango = "C" + Trim(Str(ContadorFor + 18))
ID = Range(rango).Value
'If ID = "APSK.CIHE.PT.ON.V1.201501.1A01" Then
    'Next
'End If


CourseID() = Split(ID, ".")
programa = CourseID(0)
curso = CourseID(1)
idioma = CourseID(2)
trimestre = CourseID(6)
trimestre2 = Mid(trimestre, 1, 2)
Sheets(ID).Select
Sheets(ID).Activate




'Calcular NPS de cada uno
Dim total, promoter, detractor, neutral, nps, contador, valor As Integer
Dim bool As Boolean
Dim rangoo, porcentaje As String


'inicializacion
total = 0
promoter = 0
detractor = 0
neutral = 0
contador = 9
valor = 0

rangoo = "B" + Trim(Str(contador))
bool = IsEmpty(Range(rangoo))

'while
While bool = False
    valor = Range(rangoo).Value

    If valor > 8 Then
        promoter = promoter + 1
        total = total + 1
    ElseIf valor < 7 Then
        detractor = detractor + 1
        total = total + 1
    Else
        neutral = neutral + 1
        total = total + 1
    End If


    contador = contador + 1
    rangoo = "B" + Trim(Str(contador))
    bool = IsEmpty(Range(rangoo))
Wend
'Termina While


'Escribir en celdas
Range("E8").Value = "Promoter: "
Range("E9").Value = "Neutral: "
Range("E10").Value = "Detractor: "
Range("E11").Value = "Total: "
Range("E12").Value = "NPS: "

Range("F8").Value = promoter
Range("F9").Value = neutral
Range("F10").Value = detractor
Range("F11").Value = total


Range("F8").HorizontalAlignment = xlLeft
Range("F9").HorizontalAlignment = xlLeft
Range("F10").HorizontalAlignment = xlLeft
Range("F11").HorizontalAlignment = xlLeft

nps = (promoter / total) - (detractor / total)
porcentaje = FormatPercent(nps, 2)
Range("F12").Value = porcentaje
Range("F12").HorizontalAlignment = xlLeft
'Termina de escribir'
'Fin de calculo de NPS




Range("F8:F11").Copy




'Abrir libro
'Academic Professional Skills
If programa = "APSK" Then


    Dim wb1 As Excel.Workbook
    Set wb1 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Academic Professional Skills.xlsx")


    If curso = "ACRE" Then
        wb1.Sheets("Academic Research").Select
        wb1.Sheets("Academic Research").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If


    ElseIf curso = "CIHE" Then
        wb1.Sheets("Change and Innovation in Higher Education").Select
        wb1.Sheets("Change and Innovation in Higher Education").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "CBLE" Then
        wb1.Sheets("Competency Based Learning").Select
        wb1.Sheets("Competency Based Learning").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If


    ElseIf curso = "PRS1" Then
        wb1.Sheets("Foundations of Oral Communicati").Select
        wb1.Sheets("Foundations of Oral Communicati").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If
    End If
    wb1.Save
    wb1.Close


ElseIf programa = "FACD" Then
    Dim wb2 As Excel.Workbook
    Set wb2 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Faculty Induction.xlsx ")



    If curso = "STRE" Then
        wb2.Sheets("Student Readiness").Select
        wb2.Sheets("Student Readiness").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "LFTC" Then
        wb2.Sheets("Laureate Faculty in the Twenty").Select
        wb2.Sheets("Laureate Faculty in the Twenty").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If
    End If

    wb2.Save
    wb2.Close

ElseIf programa = "COHB" Then
    Dim wb3 As Excel.Workbook
    Set wb3 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Laureate Certificate in Online, Hybrid and Blended Education.xlsx ")

    If curso = "IOHB" Then
        wb3.Sheets("Module 1").Select
        wb3.Sheets("Module 1").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "DPED" Then

        wb3.Sheets("Module 2").Select
        wb3.Sheets("Module 2").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "DLMS" Then
        wb3.Sheets("Module 3").Select
        wb3.Sheets("Module 3").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "OEFB" Then
        wb3.Sheets("Module 4").Select
        wb3.Sheets("Module 4").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "CODE" Then
        wb3.Sheets("Module 5").Select
        wb3.Sheets("Module 5").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "OTAA" Then
        wb3.Sheets("Module 6").Select
        wb3.Sheets("Module 6").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "CPTC" Then
        wb3.Sheets("Module 7").Select
        wb3.Sheets("Module 7").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    End If


wb3.Save
wb3.Close
ElseIf programa = "LCTL" Then
    Dim wb4 As Excel.Workbook
    Set wb4 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Laureate Certificate in Teaching and Learning in Higher Education.xlsx ")

    If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "UURL" Then

        wb4.Sheets("Module 2").Select
        wb4.Sheets("Module 2").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If


    ElseIf curso = "TTOL" Then
        wb4.Sheets("Module 3").Select
        wb4.Sheets("Module 3").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "ATOL" Then
        wb4.Sheets("Module 4").Select
        wb4.Sheets("Module 4").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "TECT" Then
        wb4.Sheets("Module 5").Select
        wb4.Sheets("Module 5").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    
    End If

wb4.Save
wb4.Close

ElseIf programa = "CWAE" Then
    Dim wb5 As Excel.Workbook
    Set wb5 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Laureate Certificate in Working Adult Education.xlsx")

    If curso = "UWAL" Then
        wb5.Sheets("Module 1").Select
        wb5.Sheets("Module 1").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "LACM" Then

        wb5.Sheets("Module 2").Select
        wb5.Sheets("Module 2").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If


    ElseIf curso = "TLS1" Then
        wb5.Sheets("Module 3").Select
        wb5.Sheets("Module 3").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "TLS2" Then
        wb5.Sheets("Module 4").Select
        wb5.Sheets("Module 4").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "ANDA" Then
        wb5.Sheets("Module 5").Select
        wb5.Sheets("Module 5").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    
    End If

wb5.Save
wb5.Close

ElseIf programa = "TSHE" Then
    Dim wb6 As Excel.Workbook
    Set wb6 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Learning Methods.xlsx")

    If curso = "CAST" Then
        wb6.Sheets("Case Studies Methodology").Select
        wb6.Sheets("Case Studies Methodology").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "COLE" Then

        wb6.Sheets("Collaborative Learning").Select
        wb6.Sheets("Collaborative Learning").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If


    ElseIf curso = "PRBL" Then
        wb6.Sheets("Problem Based Learning").Select
        wb6.Sheets("Problem Based Learning").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    
    
    End If

wb6.Save
wb6.Close

ElseIf programa = "PJBL" Then
    Dim wb7 As Excel.Workbook
    Set wb7 = Workbooks.Open("C:\Users\wcarrasco\Dropbox (laureate)\NPS - Walther Carrasco\NPS - Project Based Learning.xlsx")

    If curso = "PYB1" Then
        wb7.Sheets("Project Based Learning I").Select
        wb7.Sheets("Project Based Learning I").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If

    ElseIf curso = "PYB2" Then

        wb7.Sheets("Project Based Learning II").Select
        wb7.Sheets("Project Based Learning II").Activate

        If trimestre2 = "1A" Then
        
                If idioma = "EN" Then
                    Range("B5").Value = Range("B5").Value + promoter
                    Range("B6").Value = Range("B5").Value + neutral
                    Range("B6").Value = Range("B5").Value + detractor
                    Range("B6").Value = Range("B5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("B11").Value = Range("B11").Value + promoter
                    Range("B12").Value = Range("B12").Value + neutral
                    Range("B13").Value = Range("B13").Value + detractor
                    Range("B14").Value = Range("B14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("B17").Value = Range("B17").Value + promoter
                    Range("B18").Value = Range("B18").Value + neutral
                    Range("B19").Value = Range("B19").Value + detractor
                    Range("B20").Value = Range("B20").Value + total
                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "1B" Then
        
                If idioma = "EN" Then
                    Range("C5").Value = Range("C5").Value + promoter
                    Range("C6").Value = Range("C5").Value + neutral
                    Range("C6").Value = Range("C5").Value + detractor
                    Range("C6").Value = Range("C5").Value + total

                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("C11").Value = Range("C11").Value + promoter
                    Range("C12").Value = Range("C12").Value + neutral
                    Range("C13").Value = Range("C13").Value + detractor
                    Range("C14").Value = Range("C14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("C17").Value = Range("C17").Value + promoter
                    Range("C18").Value = Range("C18").Value + neutral
                    Range("C19").Value = Range("C19").Value + detractor
                    Range("C20").Value = Range("C20").Value + total
                    
                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "1C" Then
        
                If idioma = "EN" Then
                    Range("D5").Value = Range("D5").Value + promoter
                    Range("D6").Value = Range("D5").Value + neutral
                    Range("D6").Value = Range("D5").Value + detractor
                    Range("D6").Value = Range("D5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("D11").Value = Range("D11").Value + promoter
                    Range("D12").Value = Range("D12").Value + neutral
                    Range("D13").Value = Range("D13").Value + detractor
                    Range("D14").Value = Range("D14").Value + total

                    'wb1.ActiveSheet.Paste
                Else
                    Range("D17").Value = Range("D17").Value + promoter
                    Range("D18").Value = Range("D18").Value + neutral
                    Range("D19").Value = Range("D19").Value + detractor
                    Range("D20").Value = Range("D20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "2A" Then
       
                If idioma = "EN" Then
                    Range("E5").Value = Range("E5").Value + promoter
                    Range("E6").Value = Range("E5").Value + neutral
                    Range("E6").Value = Range("E5").Value + detractor
                    Range("E6").Value = Range("E5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("E11").Value = Range("E11").Value + promoter
                    Range("E12").Value = Range("E12").Value + neutral
                    Range("E13").Value = Range("E13").Value + detractor
                    Range("E14").Value = Range("E14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("E17").Value = Range("E17").Value + promoter
                    Range("E18").Value = Range("E18").Value + neutral
                    Range("E19").Value = Range("E19").Value + detractor
                    Range("E20").Value = Range("E20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
       
        ElseIf trimestre2 = "2B" Then
                
                If idioma = "EN" Then
                    Range("F5").Value = Range("F5").Value + promoter
                    Range("F6").Value = Range("F5").Value + neutral
                    Range("F6").Value = Range("F5").Value + detractor
                    Range("F6").Value = Range("F5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("F11").Value = Range("F11").Value + promoter
                    Range("F12").Value = Range("F12").Value + neutral
                    Range("F13").Value = Range("F13").Value + detractor
                    Range("F14").Value = Range("F14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("F17").Value = Range("F17").Value + promoter
                    Range("F18").Value = Range("F18").Value + neutral
                    Range("F19").Value = Range("F19").Value + detractor
                    Range("F20").Value = Range("F20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        ElseIf trimestre2 = "2C" Then
       
                If idioma = "EN" Then
                    Range("G5").Value = Range("G5").Value + promoter
                    Range("G6").Value = Range("G5").Value + neutral
                    Range("G6").Value = Range("G5").Value + detractor
                    Range("G6").Value = Range("G5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("G11").Value = Range("G11").Value + promoter
                    Range("G12").Value = Range("G12").Value + neutral
                    Range("G13").Value = Range("G13").Value + detractor
                    Range("G14").Value = Range("G14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("G17").Value = Range("G17").Value + promoter
                    Range("G18").Value = Range("G18").Value + neutral
                    Range("G19").Value = Range("G19").Value + detractor
                    Range("G20").Value = Range("G20").Value + total

                    'wb1.ActiveSheet.Paste
                End If
        
        ElseIf trimestre2 = "3A" Then
               
                If idioma = "EN" Then
                    Range("H5").Value = Range("H5").Value + promoter
                    Range("H6").Value = Range("H5").Value + neutral
                    Range("H6").Value = Range("H5").Value + detractor
                    Range("H6").Value = Range("H5").Value + total
                    
                    'wb1.ActiveSheet.Paste
                ElseIf idioma = "SP" Then
                    Range("H11").Value = Range("H11").Value + promoter
                    Range("H12").Value = Range("H12").Value + neutral
                    Range("H13").Value = Range("H13").Value + detractor
                    Range("H14").Value = Range("H14").Value + total
                    
                    'wb1.ActiveSheet.Paste
                Else
                    Range("H17").Value = Range("H17").Value + promoter
                    Range("H18").Value = Range("H18").Value + neutral
                    Range("H19").Value = Range("H19").Value + detractor
                    Range("H20").Value = Range("H20").Value + total

                    'wb1.ActiveSheet.Paste
                End If

        End If
 
    End If


wb7.Save
wb7.Close
End If 'Terminan Programas



Next

End Sub

Attribute VB_Name = "CifrasEnLetras"
Option Explicit

Function GruposALetras(ByVal Num As Integer) As String
    Dim units As Variant, teens As Variant, tens As Variant, hundreds As Variant
    Dim Result As String
    
    Dim unidades As Integer, decenas As Integer, centenas As Integer, resto As Integer

    ' Arrays de palabras
    units = Array("", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve")
    
    ' Aunque se llame "teens", incluye 10 y 20, porque en Español 20 también se escribe distinto a la mera suma del nombre de las decenas y unidades.
    
    teens = Array("diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve", _
                    "veinte", "veintiún", "veintidós", "veintitrés", "veinticuatro", "veinticinco", "veintiseis", "veintisiete", "veintiocho", "veintinueve")
    tens = Array("", "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa")
    hundreds = Array("", "cien", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos")

    ' Descomponer Num en unidades, decenas y centenas
    unidades = Num Mod 10
    decenas = (Num \ 10) Mod 10
    centenas = (Num \ 100) Mod 10
    
    ' Centenas
    If centenas > 0 Then Result = hundreds(centenas) & " "
    If centenas = 1 And (decenas > 0 Or unidades > 0) Then Result = "ciento "

    ' Decenas y unidades
    If decenas = 1 Then
        Result = Result & teens(unidades)
    ElseIf decenas = 2 Then
        Result = Result & teens(unidades + 10)
    ElseIf decenas > 2 Then
        Result = Result & tens(decenas)
        If unidades > 0 Then Result = Result & " y " & units(unidades)
    ElseIf unidades > 0 Then
        Result = Result & units(unidades)
    End If

    
    GruposALetras = Trim(Result)

End Function


Function NumerosALetras(ByVal numero As Double) As String

Dim grupos(1 To 4) As Integer
Dim i As Integer
Dim palabras As String
Dim euros As Double, cents As Integer
Dim a As Double

' Si se pasa un numero mayor que 999.999 millones, devolver un error

If numero > 999999999999# Then

    MsgBox "El número es mayor que 999.999 millones. Selecciona un número inferior"
    Exit Function
    
End If


euros = Int(numero)

cents = (numero - euros) * 100


' Separa el número en grupos de tres (centenas, millares, millones y miles de millones)

For i = 1 To 4

    grupos(i) = Int(numero) - 1000 * (Int(numero / 1000))

    numero = numero / 1000

Next i

' Toma la parte entera (Euros) y va llamando a la función GruposALetras por cada grupo, en orden inverso y siempre que sea mayor que 0, y añadiendo los sufijos correspondientes

If grupos(4) > 0 Then

    If grupos(4) = 1 Then
    
        palabras = "mil "
    
    Else
    
        palabras = GruposALetras(grupos(4)) & " mil "
    
    End If
    
    If grupos(3) = 0 Then
    
        palabras = palabras & "millones "
    
    End If

End If

If grupos(3) > 0 Then
    
    palabras = palabras & GruposALetras(grupos(3)) & " millones "

End If

If grupos(2) > 0 Then

    If grupos(2) = 1 Then
    
        palabras = palabras & "mil "
    
    Else

        palabras = palabras & GruposALetras(grupos(2)) & " mil "
        
    End If

End If

If grupos(1) > 0 Then
    
    palabras = palabras & GruposALetras(grupos(1)) & " "
    
End If

' Añade "euros" (o "euros" si euros = 1). Si euros = 0  y cents = 0, asigna "cero euros"

If euros = 0 And cents = 0 Then

    palabras = "cero euros"

ElseIf euros = 1 Then

    palabras = palabras & "euro"

ElseIf euros > 0 Then

    palabras = palabras & "euros"
    
End If

' Añadir céntimos, si los hay (cents > 0)

If cents > 0 Then

    If euros > 0 Then

        palabras = palabras & " con "
    
    End If
    
    palabras = palabras & GruposALetras(cents) & " céntimo"

    If cents > 1 Then
    
        palabras = palabras & "s"
    
    End If

End If

''''''''''''''''''' OPCIONAL: Añadir numeros con formato ''''''''''''''''''''''''''''''''''''

palabras = palabras & " (€ " & Format(euros + cents / 100, "###,###,###,##0.00") & ")" ''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

NumerosALetras = palabras



End Function

Sub ConvertirCifrasEnPalabras()

Dim palabras As String
Dim numero As Double

'Si no se ha seleccionado un numero mostrar error y salir

If Not IsNumeric(Trim(Application.Selection)) Then

    MsgBox "No se ha seleccionado un número válido."
    
    Exit Sub

End If

' Toma el valor de la selección como numero y con él llama a NumerosALetras
' Trim() sirve para evitar problemas cuando se selecciona espacios en blanco al final o al principio del numero).

numero = Trim(Application.Selection)
palabras = NumerosALetras(numero)

'Sustituye la selección por las palabras del número

Application.Selection.Range.Text = (UCase(palabras))


End Sub

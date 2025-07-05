Attribute VB_Name = "MNUMEROS"
Option Explicit

Function CADENANUM(PNUM As Double, PMON As String, ptipo As String) As String
    
    Dim WDECIMAL    As String * 2
    Dim WENTERO     As String
    Dim WMONEDA     As String
    Dim WCADENA     As String
    Dim WCONT       As Integer
    Dim WSUBENT     As String

    WDECIMAL = Right(Format$(PNUM, "#0.00"), 2)
    WENTERO = Left(Format$(PNUM, "#0.00"), Len(Format$(PNUM, "#0.00")) - 3)
    WMONEDA = IIf(PMON = "S", "NUEVOS SOLES", "DOLARES AMERICANOS")
    
    WCADENA = ""
    WCONT = 0
    WSUBENT = WENTERO
    Do While WCONT < Len(WENTERO)
        WSUBENT = Right(WENTERO, Len(WENTERO) - WCONT)
        Select Case Len(WSUBENT)
        Case Is = 3, 6, 9: WCADENA = WCADENA & FCENTENA(Mid(WSUBENT, 1, 3))
        Case Is = 2, 5, 8
            If Val(Mid(WSUBENT, 1, 2)) > 15 Then
                WCADENA = WCADENA & FDECENA(Mid(WSUBENT, 1, 2))
            Else
                WCADENA = WCADENA & FUNIDAD(Mid(WSUBENT, 1, 2), Len(WSUBENT))
                WCONT = WCONT + 1
            End If
        Case Is = 1, 4, 7: WCADENA = WCADENA & FUNIDAD(Mid(WSUBENT, 1, 1), Len(WSUBENT))
        End Select
        WCONT = WCONT + 1
    Loop

    If ptipo = "*" Then '--- NO DEVUELVE LA MONEDA
        CADENANUM = WCADENA & " CON " & WDECIMAL & "/100 "
    Else
        CADENANUM = WCADENA & " CON " & WDECIMAL & "/100 " & WMONEDA
    End If

End Function

Function FCENTENA(PCAD As String)
    
    ReDim WUNI(10) As String
    Dim WSUBCAD     As String
    
    WUNI(0) = " "
    WUNI(1) = "CIENTO "
    WUNI(2) = "DOSCIENTOS "
    WUNI(3) = "TRESCIENTOS "
    WUNI(4) = "CUATROCIENTOS "
    WUNI(5) = "QUINIENTOS "
    WUNI(6) = "SEISCIENTOS "
    WUNI(7) = "SETECIENTOS "
    WUNI(8) = "OCHOCIENTOS "
    WUNI(9) = "NOVECIENTOS "

    If PCAD = "100" Then
        WSUBCAD = "CIEN"
    Else
        WSUBCAD = WUNI(Val(Left(PCAD, 1)))
    End If

    FCENTENA = WSUBCAD

End Function

Function FDECENA(PCAD As String) As String
Dim WCAD        As String

    ReDim WUNI(10) As String
    Dim WSUBCAD     As String
    
    WCAD = Left(PCAD, 2)
    WUNI(0) = " "
    WUNI(1) = "DIEZ "
    WUNI(2) = "VEINTE "
    WUNI(3) = "TREINTA "
    WUNI(4) = "CUARENTA "
    WUNI(5) = "CINCUENTA "
    WUNI(6) = "SESENTA "
    WUNI(7) = "SETENTA "
    WUNI(8) = "OCHENTA "
    WUNI(9) = "NOVENTA "

    If Right(PCAD, 1) = 0 Then
        WSUBCAD = WUNI(Val(Left(PCAD, 1)))
    Else
        WSUBCAD = WUNI(Val(Left(PCAD, 1))) & "Y "
    End If

    FDECENA = WSUBCAD

End Function


Function FUNIDAD(PCAD As String, PLEN As Integer) As String

    ReDim WUNI(16) As String
    Dim WSUBCAD     As String
    
    WUNI(0) = " "
    WUNI(1) = "UN "
    WUNI(2) = "DOS "
    WUNI(3) = "TRES "
    WUNI(4) = "CUATRO "
    WUNI(5) = "CINCO "
    WUNI(6) = "SEIS "
    WUNI(7) = "SIETE "
    WUNI(8) = "OCHO "
    WUNI(9) = "NUEVE "
    WUNI(10) = "DIEZ "
    WUNI(11) = "ONCE "
    WUNI(12) = "DOCE "
    WUNI(13) = "TRECE "
    WUNI(14) = "CATORCE "
    WUNI(15) = "QUINCE "
           
    Select Case PLEN
        Case Is = 1, 2: WSUBCAD = WUNI(Val(PCAD))
        Case Is = 4, 5: WSUBCAD = WUNI(Val(PCAD)) & "MIL "
        Case Is = 7, 8: WSUBCAD = WUNI(Val(PCAD)) & IIf(PCAD = "1", "MILLON", "MILLONES ")
    End Select

    FUNIDAD = WSUBCAD

End Function



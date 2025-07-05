Attribute VB_Name = "Module1"
Option Explicit

Global loc As Integer
Global GOC As String
Global TOC As String * 2
Global GSOL As String
Global num_solcomp As String
Global item_solcomp As Integer
Global selecc As Boolean
Global cprv As String
Global cod_soli As String
Global solit As String
Global wcodpag As String
Global wnompag As String
Global cod_s As String
Global sold As String
Global gcodalm          As String   ' código del almacén
Global gtippro          As Integer  'tipo de producto en slado iniciales 'fvg
Global dbcabecera    As DAO.Database
Global tbcabecera    As DAO.Recordset
Global filtrox As Integer

Sub writexy(pdata As Variant, PFILA As Integer, pcolu As Integer, ptipo As Integer)

    Dim wtemp   As String
    Dim ptama   As Integer
    
    Select Case ptipo
    
    Case 0      'string
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            Printer.Print "" & pdata
    Case 1      'entero
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            Printer.Print Format$(Format(pdata, "###,###,##0"), "@@@@@@@@@@")
    Case 2      'doble
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            If Val(Format(pdata, "#0.00")) >= 0# Then
                Printer.Print Format$(Format(pdata, "###,##0.00"), "@@@@@@@@@@@")
            Else
                pdata = Val(Format(pdata, "#0.00")) * -1
                Printer.Print Format$(Format("(" & pdata & ")", "###,##0.00"), "@@@@@@@@@@@")
            End If
    Case 3      'triple
            Printer.CurrentY = PFILA
            Printer.CurrentX = pcolu
            If Val(Format(pdata, "#0.000")) >= 0# Then
                Printer.Print Format$(Format(pdata, "###,##0.000"), "@@@@@@@@@@@")
            Else
                pdata = Val(Format(pdata, "#0.000")) * -1
                Printer.Print Format$(Format(pdata, "###,##0.000"), "@@@@@@@@@@@")
            End If
    Case 4     'memo
            ptama = 50
            Do While Len(Trim(pdata)) > 0
                Printer.CurrentY = PFILA
                Printer.CurrentX = pcolu
                If right(left(pdata, ptama), 1) = " " Or Len(pdata) < ptama Then
                    Printer.Print LTrim(left(pdata, ptama))
                    pdata = right(pdata, Len(pdata) - Len(LTrim(left(pdata, ptama))))
                Else
                    wtemp = left(pdata, ptama)
                    Do While right(wtemp, 1) <> " "
                        wtemp = left(wtemp, Len(wtemp) - 1)
                    Loop
                    Printer.Print LTrim(wtemp)
                    pdata = right(pdata, Len(pdata) - Len(LTrim(wtemp)))
                End If
                PFILA = PFILA + 1
            Loop

    End Select

End Sub


Sub cabeceras()
    
    Set dbcabecera = OpenDatabase(App.Path & "\" & "Control.MDB")
    Set tbcabecera = dbcabecera.OpenRecordset("SF1PARAIN")
    tbcabecera.Index = "IDCODEMP"
    tbcabecera.Seek "=", wempresa      '<--FVG
    
    Printer.ScaleMode = 4
    Printer.FontName = "Courier New" 'Printer.Fonts(tbcabecera.Fields("F1FONNAM"))
    Printer.FontSize = 12
    Printer.FontBold = True
    writexy Trim("" & tbcabecera.Fields("f1nomemp")), 2, 1, 0
    Printer.FontSize = 8
    Printer.FontBold = False

    writexy "Fecha: ", 2, 80, 0
    writexy Format(Now, "dd/mm/yyyy"), 2, 88, 0
    writexy "Intersys - Inventario", 3, 1, 0
    writexy "Página: ", 3, 80, 0
    writexy Format(Printer.Page, "###00"), 3, 88, 0

    tbcabecera.Close
    dbcabecera.Close

End Sub

Public Sub AbrirBases(strB As String, dbsB As ADODB.Connection, blS As Boolean, podbc As String)
On Error GoTo Errores

'    Select Case strBaseDatos
'    Case 0, 2 'SQL Y LINUX
'        cnSTR = "DSN=" & podbc '& ";UID=" & strUsuSql & ";DATABASE=" & strB & ""
        If dbsB.State = 1 Then dbsB.Close
        dbsB.Open "DSN=" & podbc ', strUsuSql, strB
Errores:
    If Err.Number Then
        MsgBox Err.Description
    End If

End Sub



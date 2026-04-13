- 👋 Hi, I’m @LucaFiliol
- 👀 I’m interested in code and cars
- 🌱 I’m currently learning C++
- ⚡ Fun fact: i'm the honored one

<!---
LucaFiliol/LucaFiliol is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
Sub CopierEtSousTotal()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rowSource As Long, rowDest As Long
    Dim valeurA As String
    Dim dictFeuilles As Object
    Dim key As Variant, feuille As Variant
    Dim rngToSort As Range

    Dim colMT As Long, colA As Long, colDate As Long, colH As Long
    Dim colPalSol As Long, colRolls As Long, colPA As Long
    Dim colCA As Long, colKmsCharge As Long, colKmsVide As Long, colKmsTotal As Long

    ' ===================== FEUILLE SOURCE =====================
    Set wsSource = ThisWorkbook.Sheets("Matrice61")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    lastRow = wsSource.Cells(wsSource.Rows.Count, 5).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    ' ===================== RECHERCHE COLONNES =====================
    For i = 1 To lastCol
        Select Case Trim(wsSource.Cells(1, i).Value)
            Case "A": colA = i
            Case "MT": colMT = i
            Case "Date": colDate = i
            Case "H": colH = i
            Case "Pal Sol": colPalSol = i
            Case "Rolls": colRolls = i
            Case "PA": colPA = i
            Case "CA": colCA = i
            Case "Kms charge": colKmsCharge = i
            Case "Kms vide": colKmsVide = i
            Case "Kms Total": colKmsTotal = i
        End Select
    Next i

    ' ===================== VERIFICATION ESSENTIELLE =====================
    If colA = 0 Or colMT = 0 Or colDate = 0 Or colH = 0 Then
        MsgBox "Colonnes essentielles manquantes (A, MT, Date, H)", vbCritical
        GoTo Fin
    End If

    ' ===================== TRI =====================
    With wsSource.Sort
        .SortFields.Clear
    
        .SortFields.Add key:=wsSource.Range(wsSource.Cells(2, colMT), wsSource.Cells(lastRow, colMT)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        .SortFields.Add key:=wsSource.Range(wsSource.Cells(2, colDate), wsSource.Cells(lastRow, colDate)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        .SortFields.Add key:=wsSource.Range(wsSource.Cells(2, colH), wsSource.Cells(lastRow, colH)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
        .SetRange wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
        .header = xlYes
        .Apply
    End With


    ' ===================== FEUILLES DEST =====================
    Set dictFeuilles = CreateObject("Scripting.Dictionary")
    dictFeuilles.Add "25", Sheets("25")
    dictFeuilles.Add "34", Sheets("34")
    dictFeuilles.Add "43", Sheets("43")
    dictFeuilles.Add "45", Sheets("45")
    dictFeuilles.Add "51", Sheets("51")
    dictFeuilles.Add "54", Sheets("54")
    dictFeuilles.Add "62", Sheets("62")
    dictFeuilles.Add "63", Sheets("63")
    dictFeuilles.Add "69", Sheets("69")
    dictFeuilles.Add "84", Sheets("84")

    ' ===================== NETTOYAGE + EN-TÊTES =====================
    For Each key In dictFeuilles.Keys
        Set wsDest = dictFeuilles(key)
        wsDest.Cells.Clear
        wsSource.Rows(1).Copy
        wsDest.Rows(1).PasteSpecial xlPasteValuesAndNumberFormats
    Next key

    ' ===================== COPIE DONNEES =====================
    For rowSource = 2 To lastRow
        valeurA = Trim(wsSource.Cells(rowSource, colA).Value)
        If dictFeuilles.Exists(valeurA) Then
            Set wsDest = dictFeuilles(valeurA)
            rowDest = wsDest.Cells(wsDest.Rows.Count, colA).End(xlUp).Row + 1
            wsDest.Rows(rowDest).Resize(1, lastCol).Value = wsSource.Rows(rowSource).Resize(1, lastCol).Value
            wsSource.Rows(rowSource).Copy
            wsDest.Rows(rowDest).PasteSpecial xlPasteFormats
        End If
    Next rowSource

    ' ===================== FORMAT HEURE =====================
    For Each feuille In dictFeuilles.Keys
        Set wsDest = dictFeuilles(feuille)
        lastRow = wsDest.Cells(wsDest.Rows.Count, colA).End(xlUp).Row
        If lastRow > 1 Then
            wsDest.Range(wsDest.Cells(2, colH), wsDest.Cells(lastRow, colH)).NumberFormat = "hh:mm"
        End If
    Next feuille

    ' ===================== SOUS-TOTAUX =====================
    For Each key In dictFeuilles.Keys
        Set wsDest = dictFeuilles(key)
        lastRow = wsDest.Cells(wsDest.Rows.Count, colA).End(xlUp).Row

        If lastRow > 1 Then

            ' Sécurité colonnes
            If colMT = 0 Or colPalSol = 0 Or colRolls = 0 Or colPA = 0 _
               Or colCA = 0 Or colKmsCharge = 0 Or colKmsVide = 0 Or colKmsTotal = 0 Then
                MsgBox "Colonnes de sous-total manquantes feuille " & wsDest.Name, vbCritical
                GoTo SuiteFeuille
            End If

            wsDest.Cells.RemoveSubtotal

            wsDest.Range("A1").CurrentRegion.Subtotal _
                GroupBy:=colMT, _
                Function:=xlSum, _
                TotalList:=Array(colPalSol, colRolls, colPA, colCA, colKmsCharge, colKmsVide, colKmsTotal), _
                Replace:=True, _
                PageBreaks:=False, _
                SummaryBelowData:=True
        End If

SuiteFeuille:
    Next key

Fin:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False

    MsgBox "Copie + tri + sous-totaux terminés avec succès", vbInformation

End Sub



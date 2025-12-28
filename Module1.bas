Attribute VB_Name = "Module1"
Option Explicit

Sub Projet_VBA_Ksenia_Lvova()
   
    Dim wsData As Worksheet
    Dim wsAnalyse As Worksheet
    Dim wsDash As Worksheet
    Dim lastRowData As Long, i As Long, c As Long
    Dim writeRow As Long
    Dim currentTicker As String
    Dim rawVal As String
    
    Application.ScreenUpdating = False
    
 
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Dashboard").Delete
    Sheets("Analyse").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    
    Set wsData = ActiveSheet
    
    ' ------------------------------------------------------------------------------------------
    ' PARTIE 1 - Q1
    ' ------------------------------------------------------------------------------------------
    On Error Resume Next
    wsData.Columns("A:A").TextToColumns Destination:=wsData.Range("A1"), _
        DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False
    On Error GoTo 0
    
    lastRowData = wsData.Cells(wsData.Rows.count, 1).End(xlUp).Row
    
    wsData.Range("B:F").NumberFormat = "@"
    
    For i = 2 To lastRowData
        For c = 2 To 6
            rawVal = CStr(wsData.Cells(i, c).Value)
            If InStr(rawVal, ".") > 0 Then
                rawVal = Replace(rawVal, ".", ",")
                wsData.Cells(i, c).Value = rawVal
            End If
            If IsNumeric(rawVal) Then
                wsData.Cells(i, c).Value = CDbl(rawVal)
            End If
        Next c
    Next i

    ' ------------------------------------------------------------------------------------------
    ' PARTIE 1 - Q2
    ' ------------------------------------------------------------------------------------------
    wsData.Range("B:F").NumberFormat = "#,##0.00 ÿ"
    wsData.Columns("A:A").NumberFormat = "dd/mm/yyyy"

    ' CREATION FEUILLE ANALYSE
    Set wsAnalyse = Sheets.Add
    wsAnalyse.Name = "Analyse"
    
    wsAnalyse.Cells(1, 1).Value = "Date"
    wsAnalyse.Cells(1, 2).Value = "Ticker"
    wsAnalyse.Cells(1, 3).Value = "Close"
    wsAnalyse.Cells(1, 4).Value = "Rendement Journalier"
    
    writeRow = 2
    currentTicker = ""
    
    ' ------------------------------------------------------------------------------------------
    ' PARTIE 1 - Q3
    ' ------------------------------------------------------------------------------------------
    For i = 1 To lastRowData
        Dim cellA As Variant
        cellA = wsData.Cells(i, 1).Value
        
        If InStr(1, CStr(cellA), "Ticker", vbTextCompare) > 0 Then
            currentTicker = wsData.Cells(i, 2).Value
        ElseIf IsDate(cellA) And currentTicker <> "" Then
            If IsNumeric(wsData.Cells(i, 2).Value) Then
                wsAnalyse.Cells(writeRow, 1).Value = cellA
                wsAnalyse.Cells(writeRow, 2).Value = currentTicker
                wsAnalyse.Cells(writeRow, 3).Value = wsData.Cells(i, 2).Value
                wsAnalyse.Cells(writeRow, 4).Value = ""
                writeRow = writeRow + 1
            End If
        End If
    Next i
    
    wsAnalyse.Columns(1).NumberFormat = "dd/mm/yyyy"
    wsAnalyse.Columns(3).NumberFormat = "#,##0.00 ÿ"
    wsAnalyse.Columns(4).NumberFormat = "0.00%"
    
    Dim lastRowAnalyse As Long
    lastRowAnalyse = wsAnalyse.Cells(wsAnalyse.Rows.count, 1).End(xlUp).Row
    
    wsAnalyse.Range("A1:C" & lastRowAnalyse).Sort Key1:=wsAnalyse.Range("B1"), Order1:=xlAscending, _
                                                  Key2:=wsAnalyse.Range("A1"), Order2:=xlAscending, Header:=xlYes

    ' CALCUL RENDEMENT
    For i = 2 To lastRowAnalyse
        If wsAnalyse.Cells(i, 2).Value = wsAnalyse.Cells(i - 1, 2).Value Then
            Dim pxToday As Double, pxYest As Double
            pxToday = wsAnalyse.Cells(i, 3).Value
            pxYest = wsAnalyse.Cells(i - 1, 3).Value
            If pxYest <> 0 Then wsAnalyse.Cells(i, 4).Value = (pxToday / pxYest) - 1
        End If
    Next i

    ' ------------------------------------------------------------------------------------------
    ' PARTIE 1 - Q4, Q5, Q6
    ' ------------------------------------------------------------------------------------------
    wsAnalyse.Cells(2, 7).Value = "Actif"
    wsAnalyse.Cells(2, 8).Value = "Rendement Annuel"
    wsAnalyse.Cells(2, 9).Value = "Volatilite Annuelle"
    wsAnalyse.Cells(2, 10).Value = "Rendement Mensuel"
    wsAnalyse.Cells(2, 11).Value = "Volatilite Mensuelle"
    
    Dim assets As Variant
    assets = Array("^FCHI", "MC,PA", "AI,PA")
    
    Dim r As Integer
    For r = 0 To 2
        Dim tName As String
        tName = assets(r)
        wsAnalyse.Cells(3 + r, 7).Value = tName
        
        Dim count As Long, j As Long
        Dim returns() As Double
        count = 0
        
        For j = 2 To lastRowAnalyse
            If wsAnalyse.Cells(j, 2).Value = tName And IsNumeric(wsAnalyse.Cells(j, 4).Value) And wsAnalyse.Cells(j, 4).Value <> "" Then
                ReDim Preserve returns(count)
                returns(count) = wsAnalyse.Cells(j, 4).Value
                count = count + 1
            End If
        Next j
        
        If count > 0 Then
            Dim avg As Double, stdev As Double
            avg = Application.WorksheetFunction.Average(returns)
            stdev = Application.WorksheetFunction.StDev_S(returns)
            
            wsAnalyse.Cells(3 + r, 8).Value = avg * 252
            wsAnalyse.Cells(3 + r, 9).Value = stdev * Sqr(252)
            wsAnalyse.Cells(3 + r, 10).Value = avg * 21
            wsAnalyse.Cells(3 + r, 11).Value = stdev * Sqr(21)
        End If
    Next r
    
    wsAnalyse.Range("H3:K5").NumberFormat = "0.00%"
    wsAnalyse.Columns("G:K").AutoFit
    
    ' ------------------------------------------------------------------------------------------
    ' PARTIE 1 - Q7: GRAPHIQUE
    ' ------------------------------------------------------------------------------------------
    On Error Resume Next
    Dim chrt As ChartObject
    Set chrt = wsAnalyse.ChartObjects.Add(Left:=350, Top:=300, Width:=600, Height:=350)
    chrt.Chart.ChartArea.ClearContents
    chrt.Chart.ChartType = xlLine
    chrt.Chart.HasTitle = True
    If Err.Number = 0 Then chrt.Chart.ChartTitle.Text = "Evolution des Prix (Cloture)"
    
    Do While chrt.Chart.SeriesCollection.count > 0
        chrt.Chart.SeriesCollection(1).Delete
    Loop
    
    For r = 0 To 2
        Dim rowStart As Long, rowEnd As Long
        rowStart = 0: rowEnd = 0
        tName = assets(r)
        
        For j = 2 To lastRowAnalyse
            If wsAnalyse.Cells(j, 2).Value = tName Then
                If rowStart = 0 Then rowStart = j
                rowEnd = j
            End If
        Next j
        
        If rowStart > 0 Then
            Dim s As Series
            Set s = chrt.Chart.SeriesCollection.NewSeries
            s.Name = tName
            s.XValues = wsAnalyse.Range(wsAnalyse.Cells(rowStart, 1), wsAnalyse.Cells(rowEnd, 1))
            s.Values = wsAnalyse.Range(wsAnalyse.Cells(rowStart, 3), wsAnalyse.Cells(rowEnd, 3))
        End If
    Next r
    On Error GoTo 0
    
    ' ------------------------------------------------------------------------------------------
    ' PARTIE 1 - Q8 (+PREPARATION PARTIE 2)
    ' ------------------------------------------------------------------------------------------
    Dim colStart As Integer
    colStart = 20 ' Colonne T
    
    wsAnalyse.Range("A2:A" & lastRowAnalyse).Copy Destination:=wsAnalyse.Cells(2, colStart)
    wsAnalyse.Range(wsAnalyse.Columns(colStart), wsAnalyse.Columns(colStart)).RemoveDuplicates Columns:=1, Header:=xlNo
    
    Dim lastRowDates As Long
    lastRowDates = wsAnalyse.Cells(wsAnalyse.Rows.count, colStart).End(xlUp).Row
    wsAnalyse.Range(wsAnalyse.Cells(2, colStart), wsAnalyse.Cells(lastRowDates, colStart)).Sort _
        Key1:=wsAnalyse.Cells(2, colStart), Order1:=xlAscending, Header:=xlNo
        
    wsAnalyse.Cells(1, colStart).Value = "Date (Alignee)"
    
    Dim rngDates As Range, rngTickers As Range, rngReturns As Range
    Set rngDates = wsAnalyse.Range("A2:A" & lastRowAnalyse)
    Set rngTickers = wsAnalyse.Range("B2:B" & lastRowAnalyse)
    Set rngReturns = wsAnalyse.Range("D2:D" & lastRowAnalyse)
    
    ' ALIGNEMENT
    For r = 0 To 2
        tName = assets(r)
        wsAnalyse.Cells(1, colStart + 1 + r).Value = tName
        
        For i = 2 To lastRowDates
            Dim currentDate As Double
            currentDate = wsAnalyse.Cells(i, colStart).Value
            
            wsAnalyse.Cells(i, colStart + 1 + r).Value = _
                Application.WorksheetFunction.SumIfs(rngReturns, rngDates, currentDate, rngTickers, tName)
        Next i
    Next r
    
    ' MATRICE DE CORRELATION
    wsAnalyse.Cells(10, 7).Value = "Matrice de Correlation"
    Dim rngAsset1 As Range, rngAsset2 As Range
    Dim x As Integer, y As Integer
    
    For x = 0 To 2
        wsAnalyse.Cells(11, 8 + x).Value = assets(x)
        wsAnalyse.Cells(12 + x, 7).Value = assets(x)
        
        For y = 0 To 2
            Set rngAsset1 = wsAnalyse.Range(wsAnalyse.Cells(2, colStart + 1 + x), wsAnalyse.Cells(lastRowDates, colStart + 1 + x))
            Set rngAsset2 = wsAnalyse.Range(wsAnalyse.Cells(2, colStart + 1 + y), wsAnalyse.Cells(lastRowDates, colStart + 1 + y))
            
            On Error Resume Next
            Dim corrVal As Double
            corrVal = Application.WorksheetFunction.Correl(rngAsset1, rngAsset2)
            If Err.Number = 0 Then
                wsAnalyse.Cells(12 + x, 8 + y).Value = corrVal
            Else
                wsAnalyse.Cells(12 + x, 8 + y).Value = 0
            End If
            On Error GoTo 0
            wsAnalyse.Cells(12 + x, 8 + y).NumberFormat = "0.00"
        Next y
    Next x
    
    ' ==========================================================================================
    ' PARTIE 2: DASHBOARD
    ' ==========================================================================================
    Set wsDash = Sheets.Add
    wsDash.Name = "Dashboard"
    
    ' LARGEURS COLONNES
    wsDash.Columns("A").ColumnWidth = 2
    wsDash.Columns("B").ColumnWidth = 35
    wsDash.Columns("C").ColumnWidth = 20
    wsDash.Columns("D").ColumnWidth = 20
    wsDash.Columns("I").ColumnWidth = 20
    wsDash.Columns("J").ColumnWidth = 10
    
    ActiveWindow.DisplayGridlines = False
    wsDash.Range("A1:K30").Interior.Color = RGB(250, 250, 250)
    
    ' TITRE
    wsDash.Range("B2").Value = "DASHBOARD D'INVESTISSEMENT - ANALYSE TECHNIQUE & RISQUE"
    wsDash.Range("B2").Font.Size = 16
    wsDash.Range("B2").Font.Bold = True
    wsDash.Range("B2").Font.Color = RGB(0, 50, 100)
    
    ' PARAMETRES
    Dim riskFree As Double
    riskFree = 0.03
    wsDash.Range("I2").Value = "Taux sans risque :"
    wsDash.Range("J2").Value = riskFree
    wsDash.Range("J2").NumberFormat = "0.00%"
    
    ' DONNEES POUR BETA (SHARPE)
    Dim rngCAC_Aligned As Range, rngLVMH_Aligned As Range, rngAI_Aligned As Range
    Set rngCAC_Aligned = wsAnalyse.Range(wsAnalyse.Cells(2, colStart + 1), wsAnalyse.Cells(lastRowDates, colStart + 1))
    Set rngLVMH_Aligned = wsAnalyse.Range(wsAnalyse.Cells(2, colStart + 2), wsAnalyse.Cells(lastRowDates, colStart + 2))
    Set rngAI_Aligned = wsAnalyse.Range(wsAnalyse.Cells(2, colStart + 3), wsAnalyse.Cells(lastRowDates, colStart + 3))
    
    Dim betaLVMH As Double, betaAI As Double
    Dim sharpeLVMH As Double, sharpeAI As Double
    Dim retAnnLVMH As Double, retAnnAI As Double
    Dim volAnnLVMH As Double, volAnnAI As Double
    
    ' STATS ANNUELLES
    retAnnLVMH = wsAnalyse.Cells(4, 8).Value
    volAnnLVMH = wsAnalyse.Cells(4, 9).Value
    retAnnAI = wsAnalyse.Cells(5, 8).Value
    volAnnAI = wsAnalyse.Cells(5, 9).Value
    
    ' BETA
    betaLVMH = Application.WorksheetFunction.Slope(rngLVMH_Aligned, rngCAC_Aligned)
    betaAI = Application.WorksheetFunction.Slope(rngAI_Aligned, rngCAC_Aligned)
    
    ' SHARPE
    sharpeLVMH = (retAnnLVMH - riskFree) / volAnnLVMH
    sharpeAI = (retAnnAI - riskFree) / volAnnAI
    
    ' TABLEAU
    wsDash.Cells(5, 3).Value = "LVMH (MC.PA)"
    wsDash.Cells(5, 4).Value = "Air Liquide (AI.PA)"
    wsDash.Range("C5:D5").Font.Bold = True
    wsDash.Range("C5:D5").HorizontalAlignment = xlCenter
    wsDash.Range("C5:D5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    wsDash.Cells(6, 2).Value = "Rendement Annuel"
    wsDash.Cells(7, 2).Value = "Volatilite (Risque)"
    wsDash.Cells(8, 2).Value = "Beta (Sensibilite CAC40)"
    wsDash.Cells(9, 2).Value = "Ratio de Sharpe"
    
    wsDash.Cells(6, 3).Value = retAnnLVMH: wsDash.Cells(6, 4).Value = retAnnAI
    wsDash.Cells(7, 3).Value = volAnnLVMH: wsDash.Cells(7, 4).Value = volAnnAI
    wsDash.Cells(8, 3).Value = betaLVMH: wsDash.Cells(8, 4).Value = betaAI
    wsDash.Cells(9, 3).Value = sharpeLVMH: wsDash.Cells(9, 4).Value = sharpeAI
    
    wsDash.Range("C6:D7").NumberFormat = "0.00%"
    wsDash.Range("C8:D9").NumberFormat = "0.00"
    
    wsDash.Range("B5:D9").Borders.LineStyle = xlContinuous
    wsDash.Range("B5:D9").RowHeight = 25
    
    ' DECISION
    Dim recommendation As String
    Dim explanation As String
    Dim winnerRange As Range
    
    If sharpeAI > sharpeLVMH Then
        recommendation = "ACHETER : AIR LIQUIDE"
        explanation = "Air Liquide offre un meilleur rendement ajuste au risque (Ratio de Sharpe de " & Format(sharpeAI, "0.00") & " vs " & Format(sharpeLVMH, "0.00") & " pour LVMH). " & _
                      "De plus, son Beta de " & Format(betaAI, "0.00") & " indique une meilleure resistance en cas de baisse du marche."
        Set winnerRange = wsDash.Range("D9")
    Else
        recommendation = "ACHETER : LVMH"
        explanation = "LVMH pr?sente un profil plus agressif mais r?mun?rateur. Avec un Sharpe de " & Format(sharpeLVMH, "0.00") & ", le rendement compense la volatilit? ?lev?e."
        Set winnerRange = wsDash.Range("C9")
    End If
    
    wsDash.Range("B12").Value = "CONCLUSION DE L'ANALYSE"
    wsDash.Range("B12").Font.Underline = True
    wsDash.Range("B12").Font.Bold = True
    
    wsDash.Range("B13").Value = recommendation
    wsDash.Range("B13").Font.Size = 14
    wsDash.Range("B13").Font.Color = vbRed
    wsDash.Range("B13").Font.Bold = True
    
    wsDash.Range("B14").Value = explanation
    wsDash.Range("B14").WrapText = True
    wsDash.Range("B14:H18").Merge
    wsDash.Range("B14").VerticalAlignment = xlTop
    
    winnerRange.Interior.Color = RGB(200, 255, 200)
    
    ' EXPORT PDF + MESSAGE FIN
    Dim generatePDF As Integer
    generatePDF = MsgBox("Analyse terminee. Voulez-vous generer un rapport PDF?", vbYesNo + vbQuestion, "Export")
    
    If generatePDF = vbYes Then
        Dim pdfPath As String
        pdfPath = ThisWorkbook.Path & Application.PathSeparator & "Rapport_Investissement_" & Format(Date, "yyyymmdd") & ".pdf"
        
        On Error Resume Next
        wsDash.ExportAsFixedFormat Type:=xlTypePDF, FileName:=pdfPath, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=True
        If Err.Number = 0 Then
            MsgBox "PDF genere : " & pdfPath, vbInformation
        Else
            MsgBox "Erreur lors de la creation du PDF.", vbExclamation
        End If
        On Error GoTo 0
    Else
    
        MsgBox "Le Dashboard interactif est pret et peut etre consulte directement dans Excel." & vbCrLf & _
               "Bonne lecture!", vbInformation, "Analyse Termin?e"
    End If
    
    ' Nettoyage colonnes temp
    wsAnalyse.Range(wsAnalyse.Columns(colStart), wsAnalyse.Columns(colStart + 5)).Clear
    
    wsDash.Activate
    Application.ScreenUpdating = True

End Sub

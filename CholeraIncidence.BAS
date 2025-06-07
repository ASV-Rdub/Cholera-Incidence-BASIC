Sub CreateAdjustedIncidenceChart_Colored
    Dim oDoc As Object, oSheet As Object, oCharts As Object
    Dim oChart As Object, oDiagram As Object
    Dim oRect As New com.sun.star.awt.Rectangle
    Dim i As Integer, lastRow As Integer
    Dim region As String, colorCode As Long
    Dim regionMap As Object

    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0)

    ' Step 1: Calculate adjusted incidence in Column F
    oSheet.getCellByPosition(5, 0).String = "Adjusted Incidence Rate"
    For i = 1 To 100
        If oSheet.getCellByPosition(0, i).String = "" Then Exit For
        Dim cases As Double, pop As Double, waterQ As Double
        cases = oSheet.getCellByPosition(1, i).Value
        waterQ = oSheet.getCellByPosition(2, i).Value
        pop = oSheet.getCellByPosition(3, i).Value
        If pop > 0 And waterQ < 100 Then
            oSheet.getCellByPosition(5, i).Value = ((cases / pop) * 100000) / (100 - waterQ)
        Else
            oSheet.getCellByPosition(5, i).String = "N/A"
        End If
    Next i
    lastRow = i - 1

    ' Step 2: Copy to columns G:H:I for sorting and chart
    For i = 1 To lastRow
        oSheet.getCellByPosition(6, i).String = oSheet.getCellByPosition(0, i).String ' Country
        oSheet.getCellByPosition(7, i).Value = oSheet.getCellByPosition(5, i).Value   ' Adjusted Rate
        oSheet.getCellByPosition(8, i).String = oSheet.getCellByPosition(4, i).String ' Region
    Next i
    oSheet.getCellByPosition(6, 0).String = "Country Sorted"
    oSheet.getCellByPosition(7, 0).String = "Adjusted Rate"
    oSheet.getCellByPosition(8, 0).String = "Region"

    ' Step 3: Apply color to output columns
    Dim countryRange, rateRange, regionRange As Object
    countryRange = oSheet.getCellRangeByPosition(6, 0, 6, lastRow)
    rateRange = oSheet.getCellRangeByPosition(7, 0, 7, lastRow)
    regionRange = oSheet.getCellRangeByPosition(8, 0, 8, lastRow)

    countryRange.CellBackColor = RGB(230, 230, 230) ' light gray
    rateRange.CellBackColor = RGB(255, 255, 200)    ' light yellow
    regionRange.CellBackColor = RGB(200, 230, 255)  ' light blue

    ' Step 4: Sort by adjusted rate ascending
    Dim sortRange As Object
    sortRange = oSheet.getCellRangeByPosition(6, 0, 8, lastRow)

    Dim sortFields(0) As New com.sun.star.table.TableSortField
    sortFields(0).Field = 1
    sortFields(0).IsAscending = True

    Dim sortDesc(1) As New com.sun.star.beans.PropertyValue
    sortDesc(0).Name = "SortFields"
    sortDesc(0).Value = sortFields()
    sortDesc(1).Name = "ContainsHeader"
    sortDesc(1).Value = True

    sortRange.sort(sortDesc())

    ' Step 5: Remove existing chart
    oCharts = oSheet.Charts
    If oCharts.hasByName("AdjustedIncidenceChart") Then
        oCharts.removeByName("AdjustedIncidenceChart")
    End If

    ' Step 6: Create chart
    oRect.X = 1000
    oRect.Y = 1000
    oRect.Width = 25000
    oRect.Height = 12000

    Dim chartRange(0) As New com.sun.star.table.CellRangeAddress
    chartRange(0).Sheet = 0
    chartRange(0).StartColumn = 6
    chartRange(0).StartRow = 0
    chartRange(0).EndColumn = 7
    chartRange(0).EndRow = lastRow

    oCharts.addNewByName("AdjustedIncidenceChart", oRect, chartRange(), True, True)
    oChart = oCharts.getByName("AdjustedIncidenceChart").EmbeddedObject

    oChart.HasMainTitle = True
    oChart.Title.String = "Adjusted Cholera Incidence (Per 1% Without Water)"
    oDiagram = oChart.createInstance("com.sun.star.chart.BarDiagram")
    oDiagram.Vertical = False
    oChart.setDiagram(oDiagram)

    oDiagram.HasXAxisTitle = True
    oDiagram.HasYAxisTitle = True
    oDiagram.XAxisTitle.String = "Adjusted Incidence Rate"
    oDiagram.YAxisTitle.String = "Country"
    oDiagram.XAxisTitle.CharHeight = 14
    oDiagram.YAxisTitle.CharHeight = 14
    oDiagram.XAxis.CharHeight = 10
    oDiagram.YAxis.CharHeight = 10
    oDiagram.DataCaption = com.sun.star.chart.ChartDataCaption.VALUE
    oChart.Area.FillStyle = com.sun.star.drawing.FillStyle.NONE
    oDiagram.Wall.FillStyle = com.sun.star.drawing.FillStyle.NONE

    ' Step 7: Region color map
    Set regionMap = CreateUnoService("com.sun.star.container.EnumerableMap")
    regionMap.put("Africa", RGB(0, 102, 204))
    regionMap.put("Asia", RGB(0, 153, 0))
    regionMap.put("Europe", RGB(204, 0, 0))
    regionMap.put("Americas", RGB(255, 153, 0))
    regionMap.put("Other", RGB(128, 128, 128))

    ' Step 8: Apply region colors with alternation
    Dim series As Object
    series = oDiagram.getDataRowProperties(0)
    For i = 0 To lastRow - 1
        region = oSheet.getCellByPosition(8, i + 1).String
        If regionMap.hasByName(region) Then
            colorCode = regionMap.get(region)
        Else
            colorCode = RGB(150, 150, 150)
        End If
        If i Mod 2 = 1 Then colorCode = colorCode + RGB(25, 25, 25)
        series.setPropertyValue("Color", colorCode)
    Next i
End Sub
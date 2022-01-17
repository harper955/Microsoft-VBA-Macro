# Chinmay report, Matthew's code for auto generate the report while incoporating new raw data from the individual reports.(17/01/2022)

Sub Insert_data()
'
' Insert_data Macro
' Inserts new row 2 then Copy H1 through L1  to A2 through E2 without copying formuals then updates the 4 graphs
'

'
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Font.Bold = False
    Range("H1:L1").Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection(1).Values = _
        "='12 month returns report'!$D$2:$D$14"
    ActiveChart.SeriesCollection(1).XValues = _
        "='12 month returns report'!$A$2:$A$14"
    ActiveSheet.ChartObjects("Chart 6").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).XValues = _
        "='12 month returns report'!$A$2:$A$14"
    ActiveChart.SeriesCollection(1).Values = _
        "='12 month returns report'!$E$2:$E$14"
    ActiveChart.SeriesCollection(2).Values = _
        "='12 month returns report'!$C$2:$C$14"
    ActiveChart.SeriesCollection(3).Values = _
        "='12 month returns report'!$B$2:$B$14"
    ActiveWindow.SmallScroll Down:=25
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.SeriesCollection(1).XValues = _
        "='12 month returns report'!$A$2:$A$14"
    ActiveChart.SeriesCollection(1).Values = _
        "='12 month returns report'!$B$2:$B$14"
    ActiveChart.SeriesCollection(2).Values = _
        "='12 month returns report'!$C$2:$C$14"
    ActiveWindow.SmallScroll Down:=35
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.SeriesCollection(1).XValues = _
        "='12 month returns report'!$A$2:$A$14"
    ActiveChart.SeriesCollection(1).Values = _
        "='12 month returns report'!$D$2:$D$14"
    Range("G1").Select
End Sub

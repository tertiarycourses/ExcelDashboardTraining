Attribute VB_Name = "VBA_format"
Option Explicit

Sub SizeAndAlignCharts()
    Dim W As Long, H As Long
    Dim TopPosition As Long, LeftPosition As Long
    Dim ChtObj As ChartObject
    Dim i As Long, NumCols As Long
    
    If ActiveChart Is Nothing Then
        MsgBox "Select a chart to be used as the base for the sizing"
        Exit Sub
    End If
    
    'G2t columns
    On Error Resume Next
    NumCols = InputBox("How many columns of charts?")
    If Err.Number <> 0 Then Exit Sub
    If NumCols < 1 Then Exit Sub
    On Error GoTo 0
    
    'Get size of active chart
    W = ActiveChart.Parent.Width
    H = ActiveChart.Parent.Height
    
    'Change starting positions, if necessary
    TopPosition = 100
    LeftPosition = 20
    For i = 1 To ActiveSheet.ChartObjects.Count
        With ActiveSheet.ChartObjects(i)
            .Width = W
            .Height = H
            .Left = LeftPosition + ((i - 1) Mod NumCols) * W
            .Top = TopPosition + Int((i - 1) / NumCols) * H
        End With
    Next i
End Sub


Sub FormatAllCharts()
    Dim ChtObj As ChartObject
    For Each ChtObj In ActiveSheet.ChartObjects
      With ChtObj.Chart
        .ChartType = xlBarStacked   'xlBarStacked xlLineStacked xlColumnStacked
        .ApplyLayout 3
        .ChartStyle = 12
        .ClearToMatchStyle
        .SetElement msoElementChartTitleAboveChart
        .SetElement msoElementLegendNone
        .SetElement msoElementPrimaryValueAxisTitleNone
        .SetElement msoElementPrimaryCategoryAxisTitleNone
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 100
      End With
    Next ChtObj
End Sub



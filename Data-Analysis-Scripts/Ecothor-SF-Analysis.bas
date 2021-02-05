Attribute VB_Name = "Module1"
Sub Average_Values()
Attribute Average_Values.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Average_Values Macro
'

'
    Range("H28").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("I28").Select
    ActiveCell.FormulaR1C1 = "Ba"
    Range("P28").Select
    ActiveCell.FormulaR1C1 = "Cu"
    Range("Q28").Select
    ActiveCell.FormulaR1C1 = "Fe"
    Range("T28").Select
    ActiveCell.FormulaR1C1 = "Mg"
    Range("U28").Select
    ActiveCell.FormulaR1C1 = "Mn"
    Range("AA28").Select
    ActiveCell.FormulaR1C1 = "S"
    Range("AD28").Select
    ActiveCell.FormulaR1C1 = "Si"
    Range("AE28").Select
    ActiveCell.FormulaR1C1 = "Sr"
    Range("AJ28").Select
    ActiveCell.FormulaR1C1 = "Zn"
    Range("AK28").Select
    ActiveCell.FormulaR1C1 = "Zr"
    Range("H29").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-19]C:R[-17]C)"
    Range("H30").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-17]C:R[-16]C)"
    Range("H31").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-16]C:R[-15]C)"
    Range("H32").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-15]C:R[-14]C)"
    Range("H33").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-14]C:R[-13]C)"
    Range("H34").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-13]C:R[-12]C)"
    Range("H35").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-12]C:R[-11]C)"
    Range("D28").Select
    ActiveCell.FormulaR1C1 = "Time (mins)"
    Range("D29").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D30").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("D31").Select
    ActiveCell.FormulaR1C1 = "20"
    Range("D32").Select
    ActiveCell.FormulaR1C1 = "30"
    Range("D33").Select
    ActiveCell.FormulaR1C1 = "40"
    Range("D34").Select
    ActiveCell.FormulaR1C1 = "50"
    Range("D35").Select
    ActiveCell.FormulaR1C1 = "60"
    Range("H29:H35").Select
    Selection.AutoFill Destination:=Range("H29:AK35"), Type:=xlFillDefault
    Range("H29:AK35").Select
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("J29:K35").Select
    Selection.ClearContents
    Range("L29:O35").Select
    Selection.ClearContents
    Range("R29:S35").Select
    Selection.ClearContents
    Range("V29:Z35").Select
    Range("Z35").Activate
    Selection.ClearContents
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    Range("AB29:AC35").Select
    Selection.ClearContents
    Range("AF29:AI35").Select
    Range("AI35").Activate
    Selection.ClearContents
    Range("AJ28:AK35").Select
    Selection.Cut Destination:=Range("AF28:AG35")
    Range("AD28:AG35").Select
    Range("AG35").Activate
    Selection.Cut Destination:=Range("AB28:AE35")
    Range("AA28:AE35").Select
    Range("AE35").Activate
    Selection.Cut Destination:=Range("V28:Z35")
    Range("V28:Z35").Select
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("T28:Z35").Select
    Range("Z35").Activate
    Selection.Cut Destination:=Range("R28:X35")
    Range("P28:X35").Select
    Range("X35").Activate
    Selection.Cut Destination:=Range("J28:R35")
    Range("H28:R35").Select
    Range("R35").Activate
    Selection.Cut Destination:=Range("E28:O35")
    Range("E28:F35").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16751103
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I28:I35").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764159
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("E28:F35").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764159
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K28:M35").Select
    Range("M35").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764159
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G28:H35").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("J28:J35").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("N28:O35").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=6
    Range("E27").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Average Concentration"
    Range("E27:O27").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveWindow.SmallScroll Down:=-3
End Sub
Sub SF_and_Graph()
Attribute SF_and_Graph.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SF_and_Graph Macro
'

'
    Range("D28:D35").Select
    Selection.Copy
    Range("D38").Select
    ActiveSheet.Paste
    Range("E28:O28").Select
    Range("O28").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Range("E38").Select
    ActiveSheet.Paste
    Range("E39").Select
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=1
    Range("E40").Select
    ActiveCell.FormulaR1C1 = "=(R[-11]C-R[-10]C)/R[-11]C"
    Range("E41").Select
    ActiveCell.FormulaR1C1 = "=(R[-12]C-R[-10]C)/R[-12]C"
    Range("E42").Select
    ActiveCell.FormulaR1C1 = "=(R[-13]C-R[-10]C)/R[-13]C"
    Range("E43").Select
    ActiveCell.FormulaR1C1 = "=(R[-14]C-R[-10]C)/R[-14]C"
    Range("E44").Select
    ActiveCell.FormulaR1C1 = "=(R[-15]C-R[-10]C)/R[-15]C"
    Range("E45").Select
    ActiveCell.FormulaR1C1 = "=(R[-16]C-R[-10]C)/R[-16]C"
    Range("E40:E45").Select
    Selection.AutoFill Destination:=Range("E40:O45"), Type:=xlFillDefault
    Range("E40:O45").Select
    Rows("39:39").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Range("E37").Select
    ActiveCell.FormulaR1C1 = "Separation Factor"
    Range("E37:O37").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D39:E44").Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range("'Jan 6th 2021'!$D$39:$E$44")
   
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.FullSeriesCollection(1).Name = "=""B"""
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""Ba"""
    ActiveChart.FullSeriesCollection(2).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(2).Values = "='Jan 6th 2021'!$F$39:$F$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Mg"""
    ActiveChart.FullSeriesCollection(3).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(3).Values = "='Jan 6th 2021'!$I$39:$I$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).Name = "=""S"""
    ActiveChart.FullSeriesCollection(4).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(4).Values = "='Jan 6th 2021'!$K$39:$K$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(5).Name = "=""Si"""
    ActiveChart.FullSeriesCollection(5).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(5).Values = "='Jan 6th 2021'!$L$39:$L$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(6).Name = "=""Sr"""
    ActiveChart.FullSeriesCollection(6).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(6).Values = "='Jan 6th 2021'!$M$39:$M$44"
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    
    ActiveChart.ChartTitle.Text = "Light Metals Separation Factors"
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (mins)"
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Separation Factor (SF)"

    Range("D39:D44,G39:G44").Select
    Range("G39").Activate
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range( _
        "'Jan 6th 2021'!$D$39:$D$44,'Jan 6th 2021'!$G$39:$G$44")
    ActiveSheet.Shapes("Chart 2").IncrementLeft 78
    ActiveSheet.Shapes("Chart 2").IncrementTop 171.75
    ActiveChart.FullSeriesCollection(1).Name = "=""Cu"""
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""Fe"""
    ActiveChart.FullSeriesCollection(2).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(2).Values = "='Jan 6th 2021'!$H$39:$H$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Mn"""
    ActiveChart.FullSeriesCollection(3).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(3).Values = "='Jan 6th 2021'!$J$39:$J$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).Name = "=""Zn"""
    ActiveChart.FullSeriesCollection(4).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(4).Values = "='Jan 6th 2021'!$N$39:$N$44"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(5).Name = "=""Zr"""
    ActiveChart.FullSeriesCollection(5).XValues = "='Jan 6th 2021'!$D$39:$D$44"
    ActiveChart.FullSeriesCollection(5).Values = "='Jan 6th 2021'!$O$39:$O$44"
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    
    ActiveChart.ChartTitle.Text = "Heavy Metals Separation Factors"
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time (mins)"
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Separation Factor (SF)"
    
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 2").IncrementLeft -66
    ActiveSheet.Shapes("Chart 2").IncrementTop 2.25
End Sub

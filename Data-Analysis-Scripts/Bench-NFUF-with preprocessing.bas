Attribute VB_Name = "Module11"
 
Sub Main()

preprocessData
createSheets
summaryTable
PermeateFluxVsTime
PermFluxAndAvgMemPressureVsTime
saveAs

End Sub
Sub preprocessData()

' A small note: we may have to run this section twice to eliminate points where there are > 2 consecutive points of outlier fluxes

    ' Let's normalize the flux first
    Dim lastRow As Long
    Dim normPressure As Variant
    Dim normViscosity As Long
    
    lastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
                    
    'Get the normal pressure from the user and store it in a variable
    normPressure = InputBox("Please enter the experiment's normal pressure in psi: ")
    
    'Calculate the viscosity at 22 degrees and store it in a column
    Range("Q1") = "Calculation Intermediate"
    Range("Q2").Formula = "=EXP((-52.843)+(3703.6/(273.15+RC[-13])+5.866*LN(273.15+RC[-13])-(5.879*10^(-29))*(273.15+RC[-13])^10))"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 17), Cells(lastRow, 17)), Type:=xlFillDefault
    'Round(Exp((-52.843) + (3703.6 / 295.15) + 5.866 * 5.68748370169 - (5.879 * 10 ^ (-29)) * 295.15 ^ 10), 7)
    'Debug.Print normViscosity
    Range("N1") = "Normalized Flux"
    Range("N2").Formula = "= K2 * Q2 / 0.000975735*" & normPressure & "/((E2+F2)/2)" 'Normalized Flux
    Range("N2").AutoFill Destination:=Range(Cells(2, 14), Cells(lastRow, 14)), Type:=xlFillDefault
    
' ---------------------------------------------------------------------------------------------------------------
    'This section will delete rows where the flux is =< 15
    
    Dim j As Long
    j = 2
    
    Do While j <= ThisWorkbook.ActiveSheet.Range("K1").CurrentRegion.Rows.Count

        If Cells([j], [14]) <= 15 Then
            ThisWorkbook.ActiveSheet.Cells(j, 14).EntireRow.Delete
        Else
            j = j + 1
        End If

    Loop
    
' -----------------------------------------------------------------------------------------------------------------
    
    'Now we will delete the rows where the flux sporadically increased outside the backflush period
    
    Dim backflushFreq As Variant
    backflushFreq = InputBox("Please enter the backflush frequency in minutes below: ")
    
    'we will delete rows where the permeate flux deviates > 10% from the previous timepoint's flux on the interval of
    '[0.1, backflushx] where x is the backflush number
    ' first we must convert the time in minutes to hours
    
    Dim min2hr  As Variant
    Dim totaltime As Integer
    Dim x As Integer
    Dim k As Long
    'Dim del As Integer
    Dim rows2del As Object
    
    Set rows2del = CreateObject("System.Collections.ArrayList")
    'Dim n As Long
    
    del = 0
    min2hr = backflushFreq / 60
    totaltime = Round(((Cells([lastRow], [3]) / min2hr)), 1)
    
    'Uncomment for debugging purposes
    'Debug.Print lastRow
    'Debug.Print Round((Cells([lastRow], [3]).Value / min2hr))
    
    For x = 1 To totaltime 'Number of times backflush is occuring
        Debug.Print x
        For k = (37 + (x - 1) * backflushFreq * 6) To ((6 * backflushFreq * x) - 1)
            'Debug.Print k
            'If (Abs(Cells(k, 11).Value - Cells(k - 1, 11).Value) / Cells(k - 1, 11).Value) = 1 Then
            If k > ThisWorkbook.ActiveSheet.Range("K1").CurrentRegion.Rows.Count Then
                Exit For
            ElseIf (Abs(Cells(k, 14).Value - Cells(k - 1, 14).Value) / Cells(k - 1, 14).Value) > 0.115 Then
                'Debug.Print ("Deleted row " & k & " % Difference: " & Abs(Cells(k, 11).Value - Cells(k - 1, 11).Value) / Cells(k - 1, 11).Value)
                'ThisWorkbook.ActiveSheet.Cells(k, 11).EntireRow.Delete
                'del = del + 1
                ' Add the row index to our list to be deleted after reversing the order of the list
                rows2del.Add k
            End If
        Next k
    Next x
        
    rows2del.Reverse
    
    For Each Item In rows2del
    
        'Print the row number and difference for debugging reasons
        Debug.Print ("Deleted row " & Item & _
        " % Difference: " & Abs(Cells(Item, 14).Value - Cells(Item - 1, 14).Value) _
        / Cells(Item - 1, 14).Value)
        'Now we can delete the row going bottom up so the rows don't shift and we don't have to calculate the offset
        ThisWorkbook.ActiveSheet.Cells(Item, 14).EntireRow.Delete
        
    Next Item
    
    Debug.Print rows2del.Count
        
    
End Sub
Sub createSheets()


    ' Let us make a copy of the sheet first in case something goes wrong
    ' Then create the sheet for the summary table and our two graphs on separate sheets
     
    ActiveSheet.Copy After:=Worksheets(1)

    ActiveSheet.Name = "Summary Table"

    ActiveSheet.Copy After:=Worksheets(2)

    ActiveSheet.Name = "Permeate Flux Vs. Time"
 

    'Let the last graph simply be duplicated from the first graph to conserve code

    'ActiveSheet.Copy After:=Worksheets(3)  'moved to fourth sub

    'ActiveSheet.Name = "PF and Avg MP Vs. Time" ' moved to fourth sub
     

End Sub



Sub summaryTable()


    ' Now we will create the summary table on sheet 2

    Worksheets(2).Select
     
     
    'Create an "array" (list) of all the headings to then be looped through

    Dim headings As Object
    Dim i As Long
    Set headings = CreateObject("System.Collections.ArrayList")
   

    ' Find out if there is a way to write the list more compactly!

    headings.Add "Summary Table"
    headings.Add "Experiment Date:"
    headings.Add "Membrane Material:"
    headings.Add "Membrane Pore Size (nm):"
    headings.Add "Membrane Surface Area (m2):"
    headings.Add "Backflush Frequency (min):"
    headings.Add "Backflush Duration (sec):"
    headings.Add "Average Operating Pressure (psi):"
    headings.Add "Standard Deviation for Operating Pressure (psi):"
    headings.Add "Average Permeate Flux (LMH):"
    headings.Add "Standard Deviation for Permeate Flux (LMH):"
    headings.Add "Average Normalized Flux (LMH)"
    headings.Add "Standard Deviation for Normalized Flux (LMH)"
    headings.Add "Average Differential Pressure Loss (psi):"
    headings.Add "Average Operating Temperature (°C):"
    headings.Add "Minumum Operating Temperature (°C):"
    headings.Add "Maximum Operating Temperature (°C):"
     
    ' Loop through all the items on the list headings and write them in our desired cells

    For i = 0 To headings.Count - 1
    
        Cells(1 + i, 15).Value = headings(i)

    Next i

     
    'Reformat the headings now and autofit the column

    Columns("O:O").Select

    Columns("O:O").EntireColumn.AutoFit

    With Selection

        .Font.Bold = True

        .Font.Size = 15

    End With
     
    ' Get unknown parameters from the user and set them as variables

    Dim memPoreSize As Variant
    Dim backflushFreq As Variant
    Dim backflushDuration As Variant
    Dim normPressure As Variant
    Dim normViscosity As Long
    Dim tRng As Range
    Dim lastRow As Long
    Dim appWkFuncn As WorksheetFunction 'may not be needed as this function can only have 30 arguements
    Dim ws As Worksheet
    
    ' tRng is the range for the temperature to make it easier to refer to,
    ' and appWkFuncn is the method that is called for the operating temperature parameters
    ' Note the appWkFuncn application worksheet function average can only have up to 30 arguements
    Set tRng = Range(Range("D2"), Range("D2").End(xlDown))
    Set appWkFuncn = Application.WorksheetFunction
    Set ws = Worksheets("Summary Table")
    
    'We need to find out how many rows of data are present in order to tell excel how many rows to autofill
    'and when to stop the caluclation of the average differential pressure loss
    'lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    'Debug.Print lastRow  'uncomment this if you want to ensure that excel is accounting for the correct number of rows
    
    memPoreSize = InputBox("Please enter the membrane pore size in nm below: ")
    
    'we have to ask for the backflush frequency again because the variable backflush Freq was local to the previous Sub
    backflushFreq = InputBox("Please enter the backflush frequency in minutes below: ")
    
    backflushDuration = InputBox("Please enter the backflush duration in seconds below: ")
    MsgBox ("Choose a membrane material by clicking on cell B3 and choosing an option from the dropdown menu.")
    'normPressure = InputBox("Please enter the experiment's normal pressure in psi: ")
    
    'Start filling in the data values with values from the original data
    [P2].Value = Left(Range("A2"), [10]) ' date of experiment

    'insert a dropdown menu for membrane material? Or maybe some options on the pop-up message box?
    Range("P3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Al2O3, ZrO2, TiO2"   'membrane material
    [P4].Value = memPoreSize    ' membrane pore size'
    [P5].Value = [G2].Value ' Membrane Surface area


    ' Get these from user
    [P6].Value = backflushFreq ' Backflush Frequency
    [P7].Value = backflushDuration  ' Backflush Duration
    [P8].Value = Application.Average(Range(Cells(2, 5), Cells(lastRow, 6))) 'Average operating pressure
    [P9].Value = Application.StDev(Range(Cells(2, 5), Cells(lastRow, 6))) 'Standard Deviation for Average operating pressure
    [P10].Value = Application.Average(Range(Cells(2, 11), Cells(lastRow, 11))) 'Average permeate flux
    [P11].Value = Application.StDev(Range(Cells(2, 11), Cells(lastRow, 11))) 'Standard Deviation for Average permeate flux
    'Calculate the viscosity at 22 degrees and store it in a variable
    Range("Q2").Formula = "=EXP((-52.843)+(3703.6/(273.15+RC[-13])+5.866*LN(273.15+RC[-13])-(5.879*10^(-29))*(273.15+RC[-13])^10))"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 17), Cells(lastRow, 17)), Type:=xlFillDefault
    'Round(Exp((-52.843) + (3703.6 / 295.15) + 5.866 * 5.68748370169 - (5.879 * 10 ^ (-29)) * 295.15 ^ 10), 7)
    'Debug.Print normViscosity
    'Range("N2").Formula = "= K2 * Q2 / 0.000975735*" & normPressure & "/((E2+F2)/2)" 'Normalized Flux
    'Range("N2").AutoFill Destination:=Range(Cells(2, 14), Cells(lastRow, 14)), Type:=xlFillDefault
    [P12].Value = Application.Average(Range(Cells(2, 14), Cells(lastRow, 14))) 'Average normalized permeate flux
    [P13].Value = Application.StDev(Range(Cells(2, 14), Cells(lastRow, 14))) 'Standard Deviation for normalized permeate flux
    ' Make a new column for differential pressure loss. This column can be deleted later if you wish
    
    Range("M1") = "Differential Pressure Loss (psi)"
    Range("M2").Formula = "=F2-E2"
    
    Range("M2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 13), Cells(lastRow, 13)), Type:=xlFillDefault

    [P14].Value = Application.Average(Range(Cells(2, 13), Cells(lastRow, 13))) 'Average differential pressure loss along membrane
    [P15].Value = appWkFuncn.Average(tRng)  'Average operating temperature
    [P16].Value = appWkFuncn.Min(tRng) 'Min operating temperature
    [P17].Value = appWkFuncn.Max(tRng) 'Max operating temperature
    
    ' Fix formatting of column now
    Columns("O:P").EntireColumn.AutoFit
    Columns("P:P").Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        With Selection.Font
            .Size = 12
            .Bold = False
        End With

    End With
    
    Range("O1:P1").Select
    
    ' Reformat the header/top row of table

    With Selection

        '.MergeCells = True
        .Font.Size = 20
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlCenter

    End With
    
    ' Change Table theme

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("O1:P17"), xlYes).Name = "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium6"
    Range("Table1[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "Value"

    'copy the normalized flux column and paste it to the next sheet for plotting
    Sheets("Permeate FLux Vs. Time").Columns(14).Value = ActiveSheet.Columns(14).Value
    
    'delete raw data and move table
    Worksheets(2).Select
    Rows(Cells(18, 16).Row & ":" & Rows.Count).Delete
    Range("A:N").Delete Shift:=xlToLeft
    Range("C:Q").Delete Shift:=xlToLeft

    'un-comment this if the table didn't shift to A1 automatically upon deletion of columns A-N
    '(sort of like a failsafe to copy and paste the table)
    'Range("Table1[#All]").Copy Range("A1")

End Sub

 

Sub PermeateFluxVsTime()

    ' this sub will create the first graph on sheet 2

    Worksheets(3).Select
    

    Application.Union(Range("c2", Range("c2").End(xlDown)), Range("n2", Range("n2").End(xlDown))).Select

    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select

    ' x axis naming

    With ActiveChart.Axes(xlCategory)

        .HasTitle = True
        .AxisTitle.Text = "Time (h)"

    End With
    

    ' y axis naming

    With ActiveChart.Axes(xlValue)

        .MinimumScale = 0
        .HasTitle = True
        .AxisTitle.Text = "Permeate Flux (LMH)"

    End With


    ' Rename title and format text

    With ActiveChart.ChartTitle

        .Text = "Normalized Permeate Flux Vs. Time (h)"

        With ActiveChart.ChartTitle.Font

            .Name = "Times New Roman"
            .Size = 16
            .Bold = True

        End With

    End With


    ' Create an array with the 2 axis titles in it so you can easily loop through to reformat all the font

    Dim vAxis As Variant
   

    ' reformat the axis titles

    For Each vAxis In Array(xlCategory, xlValue)

        With ActiveChart.Axes(vAxis).AxisTitle.Format.TextFrame2.TextRange.Font

            .Name = "Times New Roman"
            .Size = 12

        End With

    Next vAxis


    ' resize and move graph 1

    With ActiveChart.Parent

         .Height = 350 ' resize
         .Width = 530  ' resize
         .Top = 25    ' reposition
         .Left = 50   ' reposition

     End With


End Sub

 

Sub PermFluxAndAvgMemPressureVsTime()


    ' Since the second graph is essentially the first graph with another data set, we can simply copy the graph/sheet as a whole to simplify things

    ActiveSheet.Copy After:=Worksheets(3)

    ActiveSheet.Name = "PF and Avg MP Vs. Time"

    'Find out how many rows there are so we can reference this for the autofill function
    Dim lastRow As Long
    Dim ws2 As Worksheet
    
    Set ws2 = Worksheets("PF and Avg MP Vs. Time")
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).Row
    lastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    'Debug.Print lastRow

    'Make another column with the average Membrane Pressure

    Range("M1") = "Average Membrane Pressure (psi)"

    Range("M2").Formula = "=AVERAGE(E2:F2)"

    Range("M2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 13), Cells(lastRow, 13)), Type:=xlFillDefault
    

    ' Now adding a second axis for the second data set (Average Membrane Pressure)

    Worksheets("PF and Avg MP Vs. Time").ChartObjects(1).Activate
   
    Dim graph2 As Chart
    Dim avgMemPressure As Series
     
    Set graph2 = ActiveSheet.ChartObjects(1).Chart


    graph2.SeriesCollection.Add Source:=Range(Cells(2, 13), Cells(lastRow, 13))
    
    Set avgMemPressure = graph2.SeriesCollection(2)
   

    With graph2.SeriesCollection(1)

        .XValues = Range(Cells(2, 3), Cells(lastRow, 3))
        .Values = Range(Cells(2, 14), Cells(lastRow, 14))
        .Name = "Normalized Permeate Flux"

    End With

    
    With graph2.SeriesCollection(2)

        .XValues = Range(Cells(2, 3), Cells(lastRow, 3))
        .Values = Range(Cells(2, 13), Cells(lastRow, 13))
        .Name = "Average Membrane Pressure (psi)"

    End With


    With graph2

        .SeriesCollection(2).AxisGroup = xlSecondary
        .HasAxis(xlValue, xlSecondary) = True
        .Axes(xlCategory, xlSecondary).CategoryType = xlAutomatic

    End With
 

    'Change graph name

    graph2.ChartTitle.Text = "Permeate Flux Vs. Average Membrane Pressure Vs. Time (h)"


    'format average membrane pressure curve

    With graph2.SeriesCollection(2)

        .Border.LineStyle = xlContinuous
        .Border.Color = RGB(255, 204, 0)
        .MarkerBackgroundColor = RGB(255, 204, 0)
        .MarkerForegroundColor = RGB(255, 204, 0)

    End With


    ' Format secondary axis

    With graph2.Axes(xlValue, xlSecondary)

        .MinimumScale = 0
        .HasTitle = True
        .HasTitle = True

        With .AxisTitle

            .Text = "Average Membrane Pressure (psi)"
            .Font.Name = "Times New Roman"
            .Font.Size = 12

        End With

    End With
    

    'Format legend

    graph2.HasLegend = True

    graph2.Legend.Position = xlLegendPositionRight
     

    'resize and reposition graph

    With ActiveChart.Parent

         .Height = 350 ' resize
         .Width = 545  ' resize
         .Top = 25    ' reposition
         .Left = 50   ' reposition

     End With
     

End Sub


Sub saveAs()

    Dim expDate As Variant
    expDate = Left(ActiveSheet.Range("A2"), 10)

    'Change the following lines as necessary
    'ChDrive "C"
    'ChDir "C:\Users\KCHONG\Documents\E2Metrix"
    'ChDir "
    'Debug.Print CurDir ' uncomment if you want to check the current directory

    'Note: The extension needs to be changed below to save to your desired folder
    'ActiveWorkbook.saveAs Filename:="C:\Users\KCHONG\Documents\E2Metrix\" & "Compiled data - " & Format(expDate, "mm/dd/yyyy") & ".xlsx" ', FileFormat:=52 'xlOpenXMLWorkbookMacroEnabled"
    
    'For some reason, you can't open macro-enabled workbooks so as long as you don't want to run the macro again on the new file, we can
    'save it as a regular excel file
    ActiveWorkbook.saveAs Filename:="C:\Users\KCHONG\Documents\E2Metrix\Compiled Data\" & "Compiled data - " & Format(expDate, "mm/dd/yyyy") & ".xlsm", FileFormat:=52
    ActiveSheet.Name = "PF and Avg MP Vs. Time"
    ActiveWorkbook.Save
    Debug.Print ("Saved as: Compiled data - " & Format(expDate, "mm/dd/yyyy") & ".xlsm")
    
End Sub

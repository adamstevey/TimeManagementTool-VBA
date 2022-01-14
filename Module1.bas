Attribute VB_Name = "Module1"
Sub Button1_Click()
    TimeSpendingForm.Show
End Sub

Sub createClusteredBarChart()
'create barchart from Time Spending Input Table
 
    'declare object variables to hold references to worksheet, source data cell range, created bar chart, and destination cell range
    Dim myWorksheet As Worksheet
    Dim mySourceData As Range
    Dim myChart As Chart
    Dim myChartDestination As Range
    Dim lastRow As Integer
    
    'identify worksheet where you want to create bar chart
    Set myWorksheet = ThisWorkbook.Worksheets("Output")
    
    'Find the last non-blank cell in column A(1) of Time Spending Input sheet
    lastRow = Worksheets("Time Spending Input").Cells(Rows.Count, 1).End(xlUp).Row
 
    With myWorksheet
 
        'identify source data (does not include the description column) by appending lastRow to column C
        Set mySourceData = Worksheets("Time Spending Input").Range("A2:C" & lastRow)
 
        'identify chart location
        Set myChartDestination = .Range("B2:H22")
 
        'create bar chart
        Set myChart = .Shapes.AddChart2(Style:=-1, XlChartType:=xlBarClustered, Left:=myChartDestination.Cells(1).Left, Top:=myChartDestination.Cells(1).Top, Width:=myChartDestination.Width, Height:=myChartDestination.Height, NewLayout:=False).Chart

    End With
 
    'set source data for created bar chart
    myChart.SetSourceData source:=mySourceData
 
End Sub

Sub ClearTimeSpendingInputTable()
'Clear table (column header not included) of Time Spending Input sheet

    'Find the last non-blank cell in column A(1) of Time Spending Input sheet
    Dim lastRow As Integer
    lastRow = Worksheets("Time Spending Input").Cells(Rows.Count, 1).End(xlUp).Row
    
    If lastRow <> 1 Then
    
        Range("A2:E" & lastRow).ClearContents
        
    End If
    
End Sub

Sub add_class_form()
    AddClassForm.Show
End Sub

Sub clear_classes()
    
    Range("E4:J4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub

Sub new_project()
    NewProjectForm.Show
End Sub

Sub clear_projects()
    Range("B3:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub

Sub add_time_spending()
    TimeSpendingForm.Show
End Sub

Sub CreatePivotTable()
'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet
'Source: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

'Determine the data range you want to pivot
  SrcData = Sheets("Tim Spending Input").Range("A1:F10").Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")

End Sub
Function totalTimeCat(my_category) As Double
    totalTimeCat = 0
    Dim category As String
    Dim lRow As Long
    Dim ws As Worksheet
    
    Set ws = Worksheets("Time Spending Input")
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    For i = 2 To lRow
        category = Worksheets("Time Spending Input").Range("A" & i)
        If category = my_category Then
            totalTimeCat = totalTimeCat + Worksheets("Time Spending Input").Range("B" & i)
        End If
    Next
    
End Function

Function totalRecTimeCat(my_category) As Double
    totalRecTimeCat = 0
    Dim lRow As Long
    Dim ws As Worksheet
    Dim lookupbool As Boolean
    lookupbool = False
    
    Set ws = Worksheets("Time Spending Input")
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    Dim recommendedTime As Double
    'A1:B7 is TimeList -> is a named range in the LookupList sheet that contains recommended time for each category
    'It is CRUCIAL to have 'false' as the 4th parameter of vlookup. True will treat 1 and 5 as an approximate match. Hence making it not useful in our context.
    recommendedTime = Application.WorksheetFunction.VLookup(my_category, Worksheets("LookupList").Range("TimeList"), 2, lookupbool)
    totalRecTimeCat = recommendedTime * 7
    
End Function

Sub suggestion()
    
    Dim categories(5) As String
        categories(0) = "Self - Study"
        categories(1) = "Class"
        categories(2) = "Commute"
        categories(3) = "Sleep"
        categories(4) = "Entertainment"
        categories(5) = "Others"
        
    Dim netTimeCat(5) As Double
    Dim lackedCats As String
    Dim overCats As String
    lackedCats = ""
    overCats = ""
    
    For i = 0 To 5
        netTimeCat(i) = totalTimeCat(categories(i)) - totalRecTimeCat(categories(i))
        If netTimeCat(i) < 0 Then
            lackedCats = lackedCats & ", " & categories(i)
        ElseIf netTimeCat(i) > 0 Then
            overCats = overCats & ", " & categories(i)
        End If
    Next
    
    MsgBox ("You should spend more time in " & lackedCats & ". You can do this by reducing your time in " & overCats & ".")
    
End Sub

Sub Button4_Click()
    Call suggestion
End Sub



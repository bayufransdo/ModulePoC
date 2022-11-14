Attribute VB_Name = "Mdl_Chart"
Option Explicit

Sub chartExport(chartName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim myChart As Chart
    Dim fname As String
    Application.ScreenUpdating = True
    Set ws = ThisWorkbook.Sheets("Charts")
    ws.Activate
    ActiveWindow.Zoom = 80
    Application.ScreenUpdating = False
    fname = ThisWorkbook.Path & "\Chart\" & chartName & ".jpg"
    Set myChart = ws.ChartObjects(chartName).Chart
    
    'biar si chart bisa di export guys
    myChart.Activate
    myChart.Export Filename:=fname, Filtername:="JPG"
    
End Sub

Sub createExportChart()
    On Error Resume Next
    Dim ws As Worksheet
    Dim sheetcht As Worksheet
    Dim Rng As Range
    Dim myChart As ChartObject
    Dim chartName As String
'    Left:=ActiveCell.Left, _
'    Width:=800, _
'    Top:=ActiveCell.Top, _
'    Height:=250)
    
    myChart.Activate
    myChart.Chart.SetSourceData Source:=Rng
    myChart.Chart.ChartType = xlColumnClustered
    myChart.Chart.ChartStyle = 47
    myChart.Chart.SeriesCollection(1).Delete
    myChart.Chart.Axes(xlValue).MaximumScale = 20
    
    Set ws = ThisWorkbook.Sheets("Charts")
'    fname = ThisWorkbook.Path & "\Chart\" & chartName & ".jpg"
    Set myChart = ws.ChartObjects(chartName).Chart
    
'    myChart.Export Filename:=fname, Filtername:="JPG"
End Sub
'Sub asdsada()
'On Error Resume Next
'    Dim ws As Worksheet
'    Dim myChart As Chart
'    Set ws = ThisWorkbook.Sheets("Charts")
'    Set myChart = ws.ChartObjects("total").Chart
'    myChart.ExportAsFixedFormat(
'End Sub

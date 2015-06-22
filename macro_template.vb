Public Sub Workbook_Open() 
' This sets the macro to automatically run upon opening the xlsm file 

' Get data from Excel Add-on PowerQuery
' it's the equivalent of "Refresh All" under "Data" tab
Workbooks(1).RefreshAll 

' (Optional) In case you need to use the TextToColumns funtion:
ThisWorkbook.Sheets("Sheet1").Columns("B:B").TextToColumns 

' (Optional) In case you need to use the filldown function: 
With ThisWorkbook.Sheets("Sheet1")
    .Range("A3:A50").FillDown
End With


' Define your dashboard's name (1st half before the date)
Dim dt As String
Dim wbNam As String

wbNam = "Your1stDash_"
' Assuming the Dashboard you send only contains data upto the end of day yesterday
dt = Format(DateAdd("d", -1, Now), "yyyymmdd") 
savePath = "C:\Users\Dude\Daily_Report\" & dt ' This is equivalent to : C:\Users\Dude\Daily_Report\20150620
MkDir savePath ' Make a new directory with the above path

' Save as new xlxs file
Application.DisplayAlerts = False ' Disable the alert of saving macro-enabled workbook (xlsm) into macro-free xlxs file
' Save to the newly created folder 
ActiveWorkbook.SaveAs Filename:=savePath & "\" & wbNam & dt, FileFormat:=xlOpenXMLWorkbook 

' (Optional) Sometimes it's helpful to include a screenshot of your dashboard so people can quickly see the dashboard's information without downloading it
' The following section takes a screenshot of Summary worksheet and export the screenshot to folder
' Set Range you want to take a screenshot
    Dim rgExp As Range: Set rgExp = Range("A1:P35")
    ''' Copy range as picture onto Clipboard
    rgExp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ''' Create an empty chart with exact size of range copied
    With ActiveSheet.ChartObjects.Add(Left:=rgExp.Left, Top:=rgExp.Top, _
    Width:=rgExp.Width, Height:=rgExp.Height)
    .Name = "ChartVolumeMetricsDevEXPORT"
    .Activate
    End With
    ''' Paste into chart area, export to file, delete chart.
    ActiveChart.Paste
    ' Every time you run this macro, this jpg file gets overwritten
    ActiveSheet.ChartObjects("ChartVolumeMetricsDevEXPORT").Chart.Export "C:\Users\Dude\Daily_Report\Your1stDash.jpg"
    ActiveSheet.ChartObjects("ChartVolumeMetricsDevEXPORT").Delete

' (Optional) Close Excel after done updating
Application.Quit()

End Sub

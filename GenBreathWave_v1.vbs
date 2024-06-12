Set Doc = GetObject ("\\chss.datastore.ed.ac.uk\chss\education\users\jkelly4\IndividualSupport\Heini\Macro InhaleExhale - JK_v0.adicht", "ADIChart.Document")
Set App = Doc.Application
Set Services = Doc.Services

Sub GenBreathWave_v1 ()
	inhalePeriod = 3 'seconds
	exhalePeriod = 4 'seconds
	numvals = 1000 ' Number of data points to generate
	pi = 3.14159265358979 ' Really, there must be a better way of doing this, surely this is available as a constant
	inhaleInc = (pi / 2)/ (numvals*(inhalePeriod/(inhalePeriod+exhalePeriod)))	' increment = pi/2 (time to peak) / number of data points scaled to inhalation as proportion of total breath cycle time
	exhaleInc = (pi / 2)/ (numvals*(exhalePeriod/(inhalePeriod+exhalePeriod)))	' this gives the nth - 1 (numvals -1) value as pi with the spread of points before and after peak scaled to inhalation/exhalation
	x = 0 'Set inital value to write

' Fill Data Pad with values for Breath Wave   
	Call Doc.DataPadSetEditMode (True)
	Call Doc.RenameDataPadColumn (1, 1, "x")
	Call Doc.RenameDataPadColumn (1, 2, "Breathwave")
	For i = 1 to numvals step+1
		Call Doc.SetDataPadValue(1, i, 1, x)	' Set column 1 to increments (0 to pi/2)
		Call Doc.GetDataPadValue (1, 1, 1) ' Retrieves the value from the top-left cell in the Data Pad sheet.
		Call Doc.GetDataPadValue (1, 1, 1) ' Retrieves the value from the top-left cell in the Data Pad sheet.
		Call Doc.GetDataPadCurrentValue (1) ' Retrieves the current value in the first column.
		Call Doc.SetDataPadValue(1, i, 2, Sin(Doc.GetDataPadValue(1,i,1))) 'Set column to sin of column 1
		If (X <= pi/2) Then 
			x = x + inhaleInc
		Else 
			x = x + exhaleInc
		End If
	next
	Call Doc.DataPadSetEditMode (False)
' Graph Breath Wave
Call Doc.CreateEmptyPlot ("", False)
Call Doc.SetViewState ("Data Pad", 1, 61728)
Call Doc.OpenView ("Data Pad")
Call Doc.SetViewState ("Plots View", 65537, 61728)

' Begin SetDataPadCellSelection
WorkSheet = 1
SelectionLeft = 2
SelectionTop = 2
SelectionWidth = 1
SelectionHeight = -1
Call Doc.SetDataPadCellSelection (WorkSheet, SelectionLeft, SelectionTop, SelectionWidth, SelectionHeight)
' End SetDataPadCellSelection

' Begin SetDataPadCellSelection
WorkSheet = 1
SelectionLeft = 2
SelectionTop = 2
SelectionWidth = 1
SelectionHeight = -1
Call Doc.SetDataPadCellSelection (WorkSheet, SelectionLeft, SelectionTop, SelectionWidth, SelectionHeight)
' End SetDataPadCellSelection

Call Doc.ShowPlotSeriesProperties ("Plot 2", "Series", True)
Call Doc.SetViewState ("Data Pad", 1, 61728)

' Begin SetEtchScaleRangeYEx
Name = "Plot 2"
Top = 1.05542
Bottom = -0.0530684
IsAutoScale = True
Call Doc.SetEtchScaleRangeYEx (Name, Top, Bottom, IsAutoScale)
' End SetEtchScaleRangeYEx

' Begin SetEtchScaleRangeXEx
Name = "Plot 2"
Top = 1055.44
Bottom = -53.4444
IsAutoScale = True
Call Doc.SetEtchScaleRangeXEx (Name, Top, Bottom, IsAutoScale)
' End SetEtchScaleRangeXEx

' Begin SetPlotProperty
PlotName = "Plot 2"
SeriesName = "General Properties"
SectionName = "Display"
Property = "Show title"
Value = "No"
Call Doc.SetPlotProperty (PlotName, SeriesName, SectionName, Property, Value)
' End SetPlotProperty

' Begin SetPlotProperty
PlotName = "Plot 2"
SeriesName = "General Properties"
SectionName = "Display"
Property = "Include new data"
Value = "No"
Call Doc.SetPlotProperty (PlotName, SeriesName, SectionName, Property, Value)
' End SetPlotProperty

' Begin SetPlotProperty
PlotName = "Plot 2"
SeriesName = "General Properties"
SectionName = "Display"
Property = "Show graticule"
Value = "No"
Call Doc.SetPlotProperty (PlotName, SeriesName, SectionName, Property, Value)
' End SetPlotProperty

' Begin SetPlotProperty
PlotName = "Plot 2"
SeriesName = "General Properties"
SectionName = "Display"
Property = "Show legend"
Value = "No"
Call Doc.SetPlotProperty (PlotName, SeriesName, SectionName, Property, Value)
' End SetPlotProperty



End Sub


Call GenBreathWave_v1 ()
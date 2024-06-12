Set Doc = CreateObject ("ADIChart.Document")
Set App = Doc.Application
Set Services = Doc.Services

Sub inhahleExhale ()

    inhalePeriod = 3 'seconds
    exhalePeriod = 4 'seconds
    
' we use the Data Pad to count up and down

' set up the Data Pad mini window
    Call Doc.DataPadSetEditMode (True)
    Call Doc.SetDataPadValue(1, 1, 30, "")
    Call Doc.RenameDataPadColumn (1, 30, "")
    Call Doc.ShowDataPadMiniwindow (30, -1, -1)
    Call Doc.FormatDataPadCells ("0.0")
    

' Begin PositionWindow
    ViewTypeId = "2148335621"
    ViewInstance = 29
    Dim Pos(3)
    Pos(0) = 10 'left edge from left
    Pos(1) = 120 'top edge from top
    Pos(2) = 310 'right edge from left - default: (0) + 206, increase to enlarge window
    Pos(3) = 220 'bottom edge from top - default: (1) + 74, increase to enlarge window
    Call Doc.PositionWindow (ViewTypeId, ViewInstance, Pos)
' End PositionWindow

Do while true
    For a = 1 to 10 step+1
        Call Doc.RenameDataPadColumn (1, 30, "Inhale")
        'Call Doc.SetDataPadValue(1, 1, 30, a)
' or
        str = "I "
        result = replace(space(a), " ", str)
        Call Doc.SetDataPadValue(1, 1, 30, result)

        Call Services.Sleep(inhalePeriod*1000/10)
    next

    For a = 10 to 1 step-1
        Call Doc.RenameDataPadColumn (1, 30, "Exhale")
        'Call Doc.SetDataPadValue(1, 1, 30, a)
' or
        str = "I "
        result = replace(space(a), " ", str)
        Call Doc.SetDataPadValue(1, 1, 30, result)

        Call Services.Sleep(exhalePeriod*1000/10)
    next
loop

End Sub

Call inhahleExhale ()
Attribute VB_Name = "modBandList"

    '***********************************************************************************
    '*
    '*  Module to write and save a list of Band List(s) and Mode
    '*  File used app.path & "\BandMode.lst"
    '*
    '*  Mark Mokoski
    '*  18-MAY-2004
    '*
    '************************************************************************************
    Private BlankStr            As String
    Dim LowLimit(256)
    Dim UpperLimit(256)
    Dim Mode(256)


Public Sub BandListWrite()

    'Write Headers to MSHFlexGrid
    Headers$ = ">      Lower Limit, KHz            |>     Upper Limit, KHz            |<  Mode        "
    frmProperties.BandGrid.FormatString = Headers$
    frmProperties.BandGrid.FixedRows = 1
    'Write black records to set up screen and get scroll bar active
    '    BlankStr = ""
    '    For x = 1 To 6
    '         frmProperties.BandGrid.AddItem BlankStr & vbTab & BlankStr, 1
    '   Next x
    'Get data from BandMode file

    On Error GoTo NoBandList
    BandFile = FreeFile
    Open (App.Path & "\BandMode.lst") For Input As BandFile
    i = 0

        Do Until EOF(BandFile) = True
            i = i + 1
            Input #BandFile, LowLimit(i), UpperLimit(i), Mode(i)
            DoEvents
        Loop

    Close BandFile
        
    'Write to MSHFlexGrid

        For c = 1 To i
            frmProperties.BandGrid.AddItem LowLimit(c) & vbTab & UpperLimit(c) & vbTab & Mode(c), c
            DoEvents
        Next c
    
    'Remove blank row from MSFlexGrod
    frmProperties.BandGrid.RemoveItem (c)

NoBandList:

End Sub

Public Sub BandListSave()

    'Save Telnet Host MSHFlexGrid to file app.path & "\BandMode.lst"

    'Read current MSHFlexGrid contents

        For i = 1 To (frmProperties.BandGrid.Rows - 1)
            frmProperties.BandGrid.Row = i
            frmProperties.BandGrid.Col = 0
            LowLimit(i) = frmProperties.BandGrid
            frmProperties.BandGrid.Col = 1
            UpperLimit(i) = frmProperties.BandGrid
            frmProperties.BandGrid.Col = 2
            Mode(i) = frmProperties.BandGrid

                If LowLimit(i) = "" Then Exit For
            DoEvents
        Next i
    
    'Write out file to disk

    On Error GoTo EndWrite
    BandFile = FreeFile
    Open (App.Path & "\BandMode.lst") For Output As BandFile

        For c = 1 To (i - 1)
        
            Write #BandFile, LowLimit(c), UpperLimit(c), Mode(c)
            DoEvents
        Next c

    Close BandFile
    
EndWrite:

End Sub

Public Function BandMode(Freq)

    'Get the operating mode per defined bands from app.path/bandmode.lst file
    'Compair all entries, and return mode type if found.
    'If no match, return "NONE" as value to calling sub.

    'Freruqency arrives as Hz, convert to KHz used in Band Mode List
    FreqKHz = (Freq / 1000)

    'Set Default BandMode return to "NONE"
    BandMode = "NONE"

    'Load in the Bad/Mode List from file App.Path\BandMode.lst
    On Error GoTo NoBandModeList
    BandFile = FreeFile
    Open (App.Path & "\BandMode.lst") For Input As BandFile
    i = 0

        Do Until EOF(BandFile) = True
            i = i + 1
            Input #BandFile, LowLimit(i), UpperLimit(i), Mode(i)
            DoEvents
        Loop

    Close BandFile

    'Compare BAnd/Mode list with DX Spot Frequency

        For c = 1 To i

                If FreqKHz >= Val(LowLimit(c)) And FreqKHz <= Val(UpperLimit(c)) Then BandMode = Mode(c)
            DoEvents
        Next c
        
    'Error fall thru point
NoBandModeList:

End Function

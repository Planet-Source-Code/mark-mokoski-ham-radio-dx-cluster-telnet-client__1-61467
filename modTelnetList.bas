Attribute VB_Name = "modTelnetList"
    Option Explicit

    '***********************************************************************************
    '*
    '*  Module to write and save a list of Telnet Hosts and Ports
    '*  File used app.path & "\TelnetHosts.lst"
    '*
    '*  Mark Mokoski
    '*  18-MAY-2004
    '*
    '************************************************************************************
    Private BlankStr              As String
    Private TelnetFile            As String
    Dim HostTelnet(256)
    Dim PortTelnet(256)
    Dim LogInTelnet(256)
    Dim PasswTelnet(256)


Public Sub TelnetListWrite()

    Dim Headers            As String
    Dim i                  As Integer
    Dim c                  As Integer

    'Set up Telnet Host MSHFlexGrid with stored file contents

    'Write Headers to MSHFlexGrid
    Headers = "<        Host Name - IP Number                          |^  Port Number|^  Login Name   |^   Password   "
    frmProperties.TelnetGrid.FormatString = Headers$
    frmProperties.TelnetGrid.FixedRows = 1
        
    'Write black records to set up screen and get scroll bar active
    '    BlankStr = ""
    '    For x = 1 To 5
    '         frmProperties.TelnetGrid.AddItem BlankStr & vbTab & BlankStr, 1
    '    Next x

    'Get data from TelnetHosts file

    On Error GoTo NoTelnetHostList
    TelnetFile = FreeFile
    Open (App.Path & "\TelnetHosts.lst") For Input As TelnetFile
    i = 0

        Do Until EOF(TelnetFile) = True
            i = i + 1
            Input #TelnetFile, HostTelnet(i), PortTelnet(i), LogInTelnet(i), PasswTelnet(i)
            DoEvents
        Loop

    Close TelnetFile
              
    'Write to MSHFlexGrid

        For c = 1 To i
            frmProperties.TelnetGrid.AddItem HostTelnet(c) & vbTab & PortTelnet(c) & vbTab & LogInTelnet(c) & vbTab & PasswTelnet(c), c
            DoEvents
        Next c
    
NoTelnetHostList:

End Sub

Public Sub TelnetListSave()

    Dim i                  As Integer
    Dim c                  As Integer

    'Save Telnet Host MSHFlexGrid to file app.path & "\TelnetHosts.lst"

    'Read current MSHFlexGrid contents

        For i = 1 To (frmProperties.TelnetGrid.Rows - 2)
            frmProperties.TelnetGrid.Row = i
            frmProperties.TelnetGrid.Col = 0
            HostTelnet(i) = frmProperties.TelnetGrid
            frmProperties.TelnetGrid.Col = 1
            PortTelnet(i) = frmProperties.TelnetGrid
            frmProperties.TelnetGrid.Col = 2
            LogInTelnet(i) = frmProperties.TelnetGrid
            frmProperties.TelnetGrid.Col = 3
            PasswTelnet(i) = frmProperties.TelnetGrid

                If HostTelnet(i) = "" Then Exit For
            DoEvents
        Next i
    
    'Write out file to disk

    On Error GoTo EndWrite
    TelnetFile = FreeFile
    Open (App.Path & "\TelnetHosts.lst") For Output As TelnetFile

        For c = 1 To (i - 1)
        
            Write #TelnetFile, HostTelnet(c), PortTelnet(c), LogInTelnet(c), PasswTelnet(c)
            DoEvents
        Next c

    Close TelnetFile
    
EndWrite:

End Sub

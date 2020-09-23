Attribute VB_Name = "DXspots"
    Dim FmCall(32767)
    Dim DXCall(32767)
    Dim DXfreq(32767)
    Dim DXcomment(32767)
    Dim DXtime(32767)
    Dim DXhour(32767)
    Dim DXspoted(32767)
    Dim SpotedDX(32767)



Public Sub DXHeardList(szInfo)

    Dim BalloonTile               As String
    Dim BalloonMessage            As String
    Dim hWndTelnet                As Long
    Dim hWndDX                    As Long
    Dim hWndActive                As Long
    Dim DXspotCaption             As String

    'Test to see if szInfo is DX spot


        If Mid(szInfo, 1, 5) = "DX de" Then

            'Get message elements
            Callfrm = ""
            RawChr = ""
            r = 6

                Do Until RawChr = ":" 'extract from call
                    r = r + 1
                    Callfm = Callfm & RawChr
                    RawChr = Mid(szInfo, r, 1)
                    DoEvents
                Loop
    
            CallDX = Trim(Mid(szInfo, 27, 10)) 'Get Dx Call
            freqDX = Trim(Mid(szInfo, (r + 1), 11)) 'Get Frequency
            'Test "Use_Local_Clock", the get spot arrival time from selected source

                If Use_Local_Clock = True Then
                    'Use local CPU clock
                    timedx = Format(UTCtime, "hhmm")
                Else
                    'Use CLuster spot time as arrival time
                    timedx = Trim(Mid(szInfo, 71, 4)) 'Get DXspot UTC time
                End If
            
            commentdx = Trim(Mid(szInfo, 40, 30)) 'Get comment info
            commentdx = Replace(commentdx, Chr(34), Space(1)) 'Get rid of quotes
        Else
            Exit Sub 'Not a DX spot
        End If

    'Build full DXHeard List in file DXHeard.lst
    i = 0
    'On (Err = 53) GoTo newfile 'File not found error
    On Error GoTo newfile 'File not found error
    DXfile = FreeFile
    Open (App.Path + "\" + "DXHeard.lst") For Input As DXfile
     
        Do Until EOF(DXfile) = True
            i = i + 1
            Input #DXfile, FmCall(i), DXCall(i), DXfreq(i), DXcomment(i), DXtime(i), DXhour(i), SpotedDX(i)
            DoEvents
        Loop

    Close DXfile
        
newfile:

        If i > 0 Then   'If records exist, do below, else create records
            DXFound = False
        
            'If using this module in packet monitor mode (nonconnected), use loop below to
            'filter out duplicate spots.  Comment out for telnet use

            '            For c = 1 To I
            '                'Search for existing entry
            '                If DXCall(c) = CallDX And DXtime(c) = timedx Then
            '                   DXFound = True   'DX Spot already in table
            '                End If
            '            Next c
            
        End If
        
        If DXFound = False Then   'Add new station DX spot heard
            i = i + 1
            FmCall(i) = Callfm
            DXCall(i) = CallDX
            DXfreq(i) = freqDX
            DXcomment(i) = commentdx
            DXtime(i) = timedx
            
            'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
            DXspotCaption = "Latest DX - " + DXCall(i) + " at " + Format((Str(Val(DXfreq(i)) / 1000)), "##,##0.000") + " MHz <" + DXtime(i) + " UTC>"
            'Set frmTelnet menu item mnuShowSpots caption to DX spot info
            frmTelnet.mnuShowSpots.Caption = DXspotCaption
            'Set Systray Incon Caption
            Call ChangeSystrayToolTip(frmSystray, DXspotCaption)
            
            'Get hWnd handles for Telnet and DX windows
            hWndTelnet = frmTelnet.HWND
            hWndDX = DXwindow.HWND
            hWndActive = GetForegroundWindow
             
            'Set Balloon Tip if main windows are not visible

                If DXwindow.Visible = False And frmTelnet.Visible = False And BalloonToolTip = True Then
                    'Close any popup menus befor balloon tip visible
                    Call frmTelnet.mnuRclose_Click
                    Call frmTimeSync.mnuTSclose_Click
                    Call PopupBalloon(frmSystray, DXspotCaption, "New DX Spot")
                End If
                
            'Set Balloon Tip if one of the main windows is not the foreground window

                If hWndTelnet <> hWndActive And hWndDX <> hWndActive And BalloonToolTip = True Then
                    'Close any popup menus befor balloon tip visible
                    Call frmTelnet.mnuRclose_Click
                    Call frmTimeSync.mnuTSclose_Click
                    Call PopupBalloon(frmSystray, DXspotCaption, "New DX Spot")
                End If
            
                If DXwindow.Visible = True Then

                    'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
                    DXwindow.Caption = "DX Cluster Telnet Client - " + DXspotCaption
            
                        With DXwindow
                            'Set DX Spot on DX window for NEW DX spot
                            .DXspotText.Caption = CallDX & Space(12 - Len(CallDX)) & freqDX & Space(16 - Len(freqDX)) & commentdx & Space(40 - Len(commentdx)) & timedx & "Z"
                            .DXspotText.Visible = True
                            .DXSpotLabel.Caption = "DX"
                            .DXSpotLabel.Visible = True
                            .DXFrame.Caption = "DX de " & Callfm
                            .DXFrame.Visible = True
                            'reset timer if there is a current spot on screen
                            .SpotTimer.Enabled = False
                            .SpotTimer.Enabled = True
                        End With

                    DoEvents
                End If
            
            'Look for DX alert set for the spot

            Dim alertmatch                As Boolean

            alertmatch = False
            DXfile = FreeFile
            On Error GoTo PlaySound
            Open App.Path & "\" & "DXwatch.lst" For Input As DXfile

                Do Until EOF(DXfile) = True
                    Input #DXfile, watchCall

                        If Mid(CallDX, 1, Len(watchCall)) = watchCall Then alertmatch = True
                    DoEvents
                Loop

            Close DXfile
            'Play Sounds for DX Spot
PlaySound:                           If DXwindow.Visible = True Or _
            frmTelnet.Visible = True Then

                If DXsound = "ON" Then
                    'See if WAVE sounds enabled

                        Select Case Sound_OK
                            '0 = no wave sounds selected
                            Case 0
                                'Sound beep on new DX spot
                                Beep
                                '1 = wave sounds selected
                            Case 1
                                'Play wave file on new DX spot
                            
                                If WAVdxspot = "" Then
                                    '"NULL" file name
                                    Beep
                                Else
                                    PlaySound (WAVdxspot)
                                    DoEvents
                                End If

                        End Select

                End If

        End If
            
    'Play Sound for DX Watch Alert

        If alertmatch = True And DXwatchEnabled = True Then
            'Set Spoted Dx Alert flag
            SpotedDX(i) = 1 '"True"
            
                If DXwatch = "ON" Then

                        If WAVdxwatch = "" Then
                            'NULL file name

                                If frmDxAlert.Visible = False Then
                                    Load frmDxAlert
                                    frmDxAlert.Visible = True
                                End If

                            Call frmDxAlert.AlertDX(CallDX, freqDX)
                            Beep
                        Else

                                If frmDxAlert.Visible = False Then
                                    Load frmDxAlert
                                    frmDxAlert.Visible = True
                                End If

                            Call frmDxAlert.AlertDX(CallDX, freqDX)
                            PlaySound (WAVdxwatch)
                            DoEvents
                        End If

                End If

        Else
            SpotedDX(i) = 0
        End If
            
    'Write out the new DX list to file
    DXfile = FreeFile
    On Error Resume Next
    Open (App.Path + "\" + "DXHeard.lst") For Append As DXfile
        
    'Append the new DX spot to end of DXheard file
    Write #DXfile, FmCall(i), DXCall(i), DXfreq(i), DXcomment(i), DXtime(i), "0"; SpotedDX(i)
           
    Close DXfile
            
    'If Archive Enabled, add Spot to DXlist.csv file
    On Error Resume Next

        If DXarchiveEnabled = True Then
            DXarcFile = FreeFile
            'For <TAB> delimited file, use below
            'Open (App.Path + "\" + "DXlist.txt") For Append As DXarcFile
            'Print #DXarcFile, UTCdate; Tab; DXtime(i); Tab; "DX de"; Tab; FmCall(i); Tab; "DX"; Tab; DXCall(i); Tab; DXfreq(i); Tab; DXcomment(i)
            'For comma delimited file (.CSV), use below
            Open (App.Path + "\" + "DXlist.csv") For Append As DXarcFile
            'Write #DXarcFile, Format(UTCdate, mmddyyyy), DXtime(i), "UTC", "DX de", FmCall(i), "DX", DXCall(i), DXfreq(i), DXcomment(i)
            Write #DXarcFile, Format(UTCdate, mmddyyyy), Format(UTCtime, "hh:mm"), "UTC", "DX de", FmCall(i), "DX", DXCall(i), DXfreq(i), DXcomment(i)
            Close DXarcFile
        End If
            
    'Add DX Spot to top of DXheard list

        If DXwindow.Visible = True Then
            DXfreq(i) = Format(DXfreq(i), "#####0.00")
            DXtime(i) = Mid(DXtime(i), 1, 2) & ":" & Mid(DXtime(i), 3, 2)
            DXwindow.DXGrid.AddItem FmCall(i) & vbTab & DXCall(i) & vbTab & DXfreq(i) & vbTab & DXtime(i) & vbTab & DXcomment(i), 1
            'See if there was a DX Alert Match, if yes (true) then
            'change the color of the row in the grid

                If SpotedDX(i) = 1 Then
                    DXwindow.DXGrid.Row = 1
                    'Run thru the colums and set the color

                        For g = 0 To 4
                            DXwindow.DXGrid.Col = g
                            DXwindow.DXGrid.CellBackColor = vbRed
                            DXwindow.DXGrid.CellForeColor = vbWhite
                            DXwindow.DXGrid.CellFontBold = True
                        Next g

                End If
                           
            'Set label for number of DX Spots currently in the DXheard list

                If i = "" Or i = 0 Then i = 0

                If i = 1 Then
                    DXwindow.DXcount.Caption = "( " & Format(i, "####0") & " DX spot in list )"
                Else
                    DXwindow.DXcount.Caption = "( " & Format(i, "####0") & " DX spots in list )"
                End If
                 
        End If

End If
        
End Sub

Public Sub ViewDXHeardList()

    Dim GridColor
    Dim FontColor
    Dim Bold                      As Boolean

    'Get data from MHeard file

    On Error GoTo DXHeardList
    DXfile = FreeFile
    Open (App.Path + "\" + "DXHeard.lst") For Input As DXfile
    i = 0

        Do Until EOF(DXfile) = True
            i = i + 1
            Input #DXfile, FmCall(i), DXCall(i), DXfreq(i), DXcomment(i), DXtime(i), DXhour(i), DXspoted(i)
            DoEvents
        Loop

    Close DXfile
        
DXHeardList:

    Load DXwindow
    DXwindow.Visible = True
    'Write Headers to MSHFlexGrid
    Headers$ = "^        DX de         |^     DX Station       |>Frequency   |^Time (UTC)|<Comments                                                                                               "
    DXwindow.DXGrid.FormatString = Headers$
    DXwindow.DXGrid.FixedRows = 1
        
    'Write to MSHFlexGrid
    b = 1
    c = i

        Do Until c = 0
            DXfreq(c) = Format(DXfreq(c), "#####0.00")
            DXtime(c) = Mid(DXtime(c), 1, 2) & ":" & Mid(DXtime(c), 3, 2)
            DXwindow.DXGrid.AddItem FmCall(c) & vbTab & DXCall(c) & vbTab & DXfreq(c) & vbTab & DXtime(c) & vbTab & DXcomment(c), b
            
                Select Case DXspoted(c)
                    Case 0
                        GridColor = vbWhite
                        FontColor = vbBlack
                        Bold = False
                    Case 1
                        GridColor = vbRed
                        FontColor = vbWhite
                        Bold = True
                    Case 2
                        GridColor = vbYellow
                        FontColor = vbBlack
                        Bold = False
                    Case 3
                        GridColor = vbMagenta
                        FontColor = vbYellow
                        Bold = True
                End Select
                
            DXwindow.DXGrid.Row = b
            'Run thru the colums and set the color

                For g = 0 To 4
                    DXwindow.DXGrid.Col = g
                    DXwindow.DXGrid.CellBackColor = GridColor
                    DXwindow.DXGrid.CellForeColor = FontColor
                    DXwindow.DXGrid.CellFontBold = Bold
                Next g
            
            c = c - 1
            b = b + 1
            DoEvents
        Loop

    'Set label for number of DX Spots currently in the DXheard list

        If i = "" Or i = 0 Then i = 0

        If i = 1 Then
            DXwindow.DXcount.Caption = "( " & Format(i, "####0") & " DX spot in list )"
        Else
            DXwindow.DXcount.Caption = "( " & Format(i, "####0") & " DX spots in list )"
        End If
        
End Sub

Public Sub DXGridRefresh()

    'Test to see if window open

        If DXwindow.Visible = False Then Exit Sub
    'Clear the DXHeard table

        For X = 1 To (DXwindow.DXGrid.Rows - 2)
            DXwindow.DXGrid.RemoveItem 1
            DoEvents
        Next X

    'Reload DXHeard table
    Call ViewDXHeardList

End Sub

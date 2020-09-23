Attribute VB_Name = "WWVspots"
   
    

    Dim Callfm(1000)
    Dim WWVtime(1000)
    Dim WWVnumbers(1000)
    Dim WWVhour(1000)

Public Sub WWVinfo(szInfo)

    Dim WWVCaption                As String
    Dim BalloonTile               As String
    Dim BalloonMessage            As String
    Dim hWndTelnet                As Long
    Dim hWndDX                    As Long
    Dim hWndActive                As Long

    'Test to see if szInfo is WWV spot





        If Mid(szInfo, 1, 6) = "WWV de" Or Mid(szInfo, 1, 6) = "WCY de" Or Mid(szInfo, 1, 6) = "JJY de" Then
    
            'Get message elements
            Callfrm = ""
            RawChr = ""
            r = 6

                Do Until RawChr = ":"  'extract from call
                    r = r + 1
                    WWVCall = WWVCall & RawChr
                    RawChr = Mid(szInfo, r, 1)
                    DoEvents
                Loop

            WWVCall = Trim(WWVCall)   'Trim Spaces
            
            t = InStr(1, WWVCall, "<")
            'See id UTC hour is part of call ( ex. W1XYZ<12>: )

                If t > 0 Then '"<" in call string, remove all after
                    WWVCall = Mid(WWVCall, 1, (t - 1))
                End If
            
            'Get rid of <CTRL-G> and <CRLF>
            infoWWV = Mid(szInfo, (r + 1), ((Len(szInfo) - 1) - (r + 1)))
            infoWWV = Trim(infoWWV) 'Trim extra spaces
        
            'Use current system UTC for WWV report
            timeWWV = Mid(UTCtime, 1, 5)
            
        Else
            Exit Sub 'Not a WWV spot
        End If
    
    'Build full WWVHeard List in file DXHeard.lst
    i = 0
    'On (Err = 53) GoTo newfile 'File not found error
    On Error GoTo newfile 'File not found error
    WWVfile = FreeFile
    Open (App.Path + "\" + "WWVHeard.lst") For Input As WWVfile
     
        Do Until EOF(WWVfile) = True
            i = i + 1
            Input #WWVfile, Callfm(i), WWVtime(i), WWVnumbers(i), WWVhour(i)
            DoEvents
        Loop

    Close WWVfile
        
newfile:

        If i > 0 Then   'If records exist, do below, else create records
            WWVFound = False
        
            'If using this module in packet monitor mode (nonconnected), use loop below to
            'filter out duplicate spots.  Comment out for telnet use

            '            For c = 1 To I
            '                'Search for existing entry
            '                If Callfm(c) = WWVCall And WWVnumbers(c) = infowwv Then
            '                   WWVFound = True   'WWV Spot already in table
            '                End If
            '                DoEvents
            '            Next c
            
        End If
        
        If WWVFound = False Then   'Add new station WWV spot heard
            i = i + 1
            Callfm(i) = WWVCall
            WWVnumbers(i) = infoWWV
            WWVtime(i) = timeWWV
            
                If DXwindow.Visible = True Then

                        With DXwindow
                            'Set WWV report on DX window for NEW WWV report
                            .DXspotText.Caption = infoWWV
                            .DXspotText.Visible = True
                            .DXSpotLabel.Caption = "WWV"
                            .DXSpotLabel.Visible = True
                            .DXFrame.Caption = "WWV de " & WWVCall
                            .DXFrame.Visible = True
                            'reset timer if there is a current spot on screen
                            .SpotTimer.Enabled = False
                            .SpotTimer.Enabled = True
                        End With
                
                    DoEvents
                Else
                    DoEvents
                
                End If
            
            'Play WWV Sounds

                If DXwindow.Visible = True Or _
                    frmTelnet.Visible = True Then
                    
                        If WWVsound = "ON" Then
                            'See if WAVE sounds enabled

                                Select Case Sound_OK
                                    '0 = no wave sounds selected
                                    Case 0
                                        'Sound beep on new WWV spot
                                        Beep
                                        '1 = wave sounds selected
                                    Case 1
                                        'Play wave file on new WWV spot

                                        If WAVwwv = "" Then
                                            '"NULL" file name
                                            Beep
                                        Else
                                            PlaySound (WAVwwv)
                                        End If

                                End Select

                        End If

                End If
            
            'Write out the new WWV list to file
            WWVfile = FreeFile
            'Append the new WWV report to end of WWVheard file
            On Error Resume Next
            Open (App.Path + "\" + "WWVHeard.lst") For Append As WWVfile
        
            Write #WWVfile, Callfm(i), WWVtime(i), WWVnumbers(i), "0"
               
            Close WWVfile
            'Get hWnd handles for Telnet and DX windows
            hWndTelnet = frmTelnet.HWND
            hWndDX = DXwindow.HWND
            hWndActive = GetForegroundWindow
            
            'Set Form caption to "DX Cluster Telnet Client - " and the last DX spot
            WWVCaption = "WWV - DE " + WWVCall + "  " + infoWWV + " <" + timeWWV + " UTC>"
            
            'Set Balloon Tip if main windows are not visible

                If DXwindow.Visible = False And frmTelnet.Visible = False And BalloonToolTip = True Then
                    'Close any popup menus befor balloon tip visible
                    Call frmTelnet.mnuRclose_Click
                    Call frmTimeSync.mnuTSclose_Click
                    Call PopupBalloon(frmSystray, WWVCaption, "New WWV Report")
                End If
            
            'Set Balloon Tip if one of the main windows is not the foreground window

                If hWndTelnet <> hWndActive And hWndDX <> hWndActive And BalloonToolTip = True Then
                    'Close any popup menus befor balloon tip visible
                    Call frmTelnet.mnuRclose_Click
                    Call frmTimeSync.mnuTSclose_Click
                    Call PopupBalloon(frmSystray, WWVCaption, "New WWV Report")
                End If
            
            'Add DX Spot to top of DXheard list

                If DXwindow.Visible = True Then
                    DXwindow.WWVGrid.AddItem Callfm(i) & vbTab & WWVtime(i) & vbTab & WWVnumbers(i), 1
                    'Set label for number of WWV reports currently in the WWVheard list

                        If i = "" Or i = 0 Then i = 0

                        If i = 1 Then
                            DXwindow.WWVcount.Caption = "( " & Format(i, "####0") & " WWV report in list )"
                        Else
                            DXwindow.WWVcount.Caption = "( " & Format(i, "####0") & " WWV reports in list )"
                        End If

                End If

        End If

End Sub

Public Sub ViewWWVHeardList()

    'Get data from WWVHeard file

    On Error GoTo WWVHeardList
    WWVfile = FreeFile
    Open (App.Path + "\" + "WWVHeard.lst") For Input As WWVfile
    i = 0

        Do Until EOF(WWVfile) = True
            i = i + 1
            Input #WWVfile, Callfm(i), WWVtime(i), WWVnumbers(i), WWVhour(i)
            DoEvents
        Loop

    Close WWVfile
        
WWVHeardList:

    Load DXwindow
    DXwindow.Visible = True
    'Write Headers to MSHFlexGrid
    Headers$ = "^        WWV de         |^Time (UTC)|<WWV Information                                                                                                                                     "
    DXwindow.WWVGrid.FormatString = Headers$
    DXwindow.WWVGrid.FixedRows = 1
        
    'Write to MSHFlexGrid
    b = 1
    c = i

        Do Until c = 0
            
            DXwindow.WWVGrid.AddItem Callfm(c) & vbTab & WWVtime(c) & vbTab & WWVnumbers(c), b
            c = c - 1
            b = b + 1
            DoEvents
        Loop
        
    'Set label for number of WWV reports currently in the WWVheard list

        If i = "" Or i = 0 Then i = 0

        If i = 1 Then
            DXwindow.WWVcount.Caption = "( " & Format(i, "####0") & " WWV report in list )"
        Else
            DXwindow.WWVcount.Caption = "( " & Format(i, "####0") & " WWV reports in list )"
        End If

End Sub

Public Sub WWVGridRefresh()

    'Test to see if window open

        If DXwindow.Visible = False Then Exit Sub
    'Clear the WWVHeard table

        For X = 1 To (DXwindow.WWVGrid.Rows - 2)
            DXwindow.WWVGrid.RemoveItem 1
            DoEvents
        Next X

    'Reload WWVHeard table
    Call ViewWWVHeardList

End Sub

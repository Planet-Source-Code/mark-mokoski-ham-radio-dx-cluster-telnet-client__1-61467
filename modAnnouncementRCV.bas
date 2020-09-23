Attribute VB_Name = "modAnnouncementRCV"
    Option Explicit

    Dim RcvComment
    Dim RawChr
    Dim Callfrm
    Dim r
    Dim YourCall



Public Sub RcvAnnoncement(szInfo)

    'See if ShowAnnouncements flag is set (true), if true do code below, else exit

        If ShowAnnouncements = True Then

            'Test to see if szInfo is DX spot

                If UCase(Mid(szInfo, 1, 2)) = "TO" Then
        
                    'Get enterd call from regisrty, exclude any of your announcments from triggering the form
                    YourCall = UCase(GetSetting("WA1ZEK", "DXtelnet", "UserCall", "None"))
                    'Get the "de" position in the string, add 3 to start call string
                    r = InStr(1, szInfo, " de ")
                    r = r + 3
                    'Loop thru string after "de" and extract call

                        Do Until RawChr = ":" Or RawChr = "-" Or RawChr = " " 'extract from call
                            r = r + 1
                            Callfrm = Callfrm & RawChr
                            RawChr = Mid$(szInfo, r, 1)
                            DoEvents
                        Loop

                    'Force uppercase
                    Callfrm = UCase(Callfrm)
                    'compare your call with announcment call, exit sub if same (you sent it)

                        If YourCall = Callfrm Then Exit Sub
            
                    'Get message elements
                    Callfrm = ""
                    RawChr = ""
                    r = 0

                        Do Until RawChr = ":" 'extract from call
                            r = r + 1
                            Callfrm = Callfrm & RawChr
                            RawChr = Mid$(szInfo, r, 1)
                            DoEvents
                        Loop
            
                        With frmRCVannouncement
                            .Visible = True
                            .AnncTimer.Enabled = False  'Stop and reset timer if window already open
                            .AnncTimer.Enabled = True   'Restart timer for this announcement
                            .RCVframe.Caption = Callfrm
                            .AccounceText.Caption = Mid$(szInfo, (r + 1), (Len(szInfo) - 1))
                            .SetFocus
                        End With
               
                        If DXwindow.Visible = True Or _
                            frmTelnet.Visible = True Then
        
                                If ANNsound = "ON" Then
                                    'See if WAVE sounds enabled

                                        Select Case Sound_OK
                                            '0 = no sound card detected
                                            Case 0
                                                'Sound beep on new DX spot
                                                Beep
                                                '1 = SOund card detected, play .WAV file
                                            Case 1
                                                'Play wave file on new DX spot
                                    
                                                If WAVannounce = "" Then
                                                    '"NULL" file name
                                                    Beep
                                                Else
                                                    PlaySound (WAVannounce)
                                                    DoEvents
                                                End If

                                        End Select

                                End If

                        End If

                Else
                    Exit Sub 'Not a DX spot
                End If

        End If

End Sub

Attribute VB_Name = "modExpireDX"
    '***********************************************************************************
    '*
    '*  Module to purge DX and WWV spots based on Exipre_DX and Expire_WWV times (Hours)
    '*
    '*  Mark Mokoski
    '*  11-APR-2003
    '*
    '************************************************************************************

    Dim FmCall(32767)
    Dim DXCall(32767)
    Dim DXfreq(32767)
    Dim DXcomment(32767)
    Dim DXtime(32767)
    Dim DXhour                 As Integer
    Dim DXlastHour(32767)
    Dim DXspoted(32767)
    Dim WWVhour                As Integer
    Dim WWVtime(32767)
    Dim WWVnumbers(32767)
    Dim WWVlastHour(32767)
    Dim StartRecord            As Long
    Dim MaxRecord              As Long


Public Sub ExpireDX()

    Call PurgeDX
    Call PurgeWWV

End Sub

Private Sub PurgeDX()

    'Get time offset
    DXhour = Val(Expire_DX)
        
    'Load DX list into array
    i = 0
    'On (Err = 53) GoTo NoDXFile 'File not found error
    On Error GoTo NoDXFile 'File not found error
    DXfile = FreeFile
    Open (App.Path + "\" + "DXHeard.lst") For Input As DXfile
     
        Do Until EOF(DXfile) = True
            i = i + 1
            Input #DXfile, FmCall(i), DXCall(i), DXfreq(i), DXcomment(i), DXtime(i), DXlastHour(i), DXspoted(i)
            DoEvents
            MaxRecord = i
            DoEvents
        Loop

    Close DXfile
        
    'Get current font and cell colors

        If DXwindow.Visible = True Then
            g = 1

                Do Until g = MaxRecord + 1
                    DXwindow.DXGrid.Row = i
                    DXwindow.DXGrid.Col = 0
            
                        Select Case DXwindow.DXGrid.CellBackColor
                            Case vbWhite
                                DXspoted(g) = 0
                            Case vbRed
                                DXspoted(g) = 1
                            Case vbYellow
                                DXspoted(g) = 2
                            Case vbMagenta
                                DXspoted(g) = 3
                        End Select

                    i = i - 1
                    g = g + 1
                Loop

        End If
   
    'Now find oldset record that at or new than time offset
    i = 1
    StartRecord = i
    
        Do Until Val(DXlastHour(i)) < DXhour
            i = i + 1
            StartRecord = i
            DoEvents
        Loop
    
    'Write new DXlist records
    
    Kill (App.Path + "\" + "DXHeard.lst")
    DXfile = FreeFile
    Open (App.Path + "\" + "DXHeard.lst") For Append As DXfile
    
        For i = StartRecord To MaxRecord
            Write #DXfile, FmCall(i), DXCall(i), DXfreq(i), DXcomment(i), DXtime(i), Str$(Val(DXlastHour(i)) + 1), DXspoted(i)
            DoEvents
        Next i
    
    Close DXfile
    
    'Refresh the DXlist grid
    Call DXspots.DXGridRefresh
    
NoDXFile:

End Sub

Private Sub PurgeWWV()

    'Get time offset
    WWVhour = Val(Expire_WWV)
        
    'Load WWV list into array
    i = 0
    'On (Err = 53) GoTo NoWWVFile 'File not found error
    On Error GoTo NoWWVFile 'File not found error
    WWVfile = FreeFile
    Open (App.Path + "\" + "WWVHeard.lst") For Input As WWVfile
     
        Do Until EOF(WWVfile) = True
            i = i + 1
            Input #WWVfile, FmCall(i), WWVtime(i), WWVnumbers(i), WWVlastHour(i)
            MaxRecord = i
            DoEvents
        Loop

    Close WWVfile
        
    'Now find oldset record that at or new than time offset
    i = 1
    StartRecord = i
    
        Do Until Val(WWVlastHour(i)) < WWVhour
            i = i + 1
            StartRecord = i
            DoEvents
        Loop
    
    'Write new WWVlist records
    
    Kill (App.Path + "\" + "WWVHeard.lst")
    WWVfile = FreeFile
    Open (App.Path + "\" + "WWVHeard.lst") For Append As WWVfile
    
        For i = StartRecord To MaxRecord
            Write #WWVfile, FmCall(i), WWVtime(i), WWVnumbers(i), Str$(Val(WWVlastHour(i)) + 1)
            DoEvents
        Next i
    
    Close WWVfile
    
    'Refresh the WWVlist grid
    Call WWVspots.WWVGridRefresh
    
NoWWVFile:

End Sub

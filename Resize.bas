Attribute VB_Name = "Resize"
    Option Explicit

    ' ==================================================================================
    '  Here is some code I received from a friend that works.  Over the years at least a
    '  dozen times it has been modified and we have no idea who originated this code.
    '  So I apologize I don't have any programmer information, but who ever it was
    '  they did a good job.  I have fixed up the code to the point where it works great
    '  now.
    '
    '  Just place the following in your form that you wish to resize:
    '      Private Sub Form_Resize()
    '            ResizeFormFor Me
    '      End Sub
    '
    '  For questions and comments, or if you find out who the originator of this code
    '  is, let me know: billsecond@clear100.com
    '
    ' ==================================================================================



    Private FormRecord()            As FObject
    Private CMobj()                 As FObject
    Private BArx                    As Boolean
    Private FormMax                 As Long
    Private ControlMax              As Long

        Private Type FObject
            Name                    As String
            Index                   As Long
            Parent                  As String
            Top                     As Long
            Left                    As Long
            Height                  As Long
            Width                   As Long
            Rows                    As Long
            ScaleHeight             As Long
            ScaleWidth              As Long
            FontSize                As Long
            CellFontSize            As Long
            FontBold                As Boolean
            RTBtext                 As String
        End Type

Public Sub ResizeFormFor(ByRef FormToResize As Form)

    Dim FormControl               As Control
    Dim MDIStartHeight            As Long
    Dim MDIStartWidth             As Long
    Dim bIsANewForm               As Boolean






        If Not BArx Then
            BArx = True

                If FindForm(FormToResize) < 0 Then
                    bIsANewForm = True
                Else
                    bIsANewForm = False
                End If

                For Each FormControl In FormToResize
                    ResizeControl FormControl, FormToResize
                Next FormControl

            BArx = False
        End If





    Set FormControl = Nothing
End Sub

Private Sub ResizeControl(ByRef FResize As Control, ByRef FormToResize As Form)








    On Error Resume Next

    Dim i                         As Long
    Dim x                         As Integer
    Dim lTop                      As Long
    Dim lLeft                     As Long
    Dim lWidth                    As Long
    Dim lHeight                   As Long
    Dim lFontSize                 As Long
    Dim yRatio                    As Long
    Dim xRatio                    As Long
    Dim OldSize                   As Integer
    
    xRatio = GetWidthRatio(FormToResize)
    yRatio = GetHeightRatio(FormToResize)
    i = FindControl(FResize, FormToResize.Name)

        If FResize.Left < 0 Then
            lLeft = CLng(((CMobj(i).Left * xRatio) \ 100) - 75000)
        Else
            lLeft = CLng((CMobj(i).Left * xRatio) \ 100)
        End If

    lTop = CLng((CMobj(i).Top * yRatio) \ 100)

        If TypeOf FResize Is CommandButton Or _
            TypeOf FResize Is TextBox Then
        
            lWidth = CLng((CMobj(i).Width) * xRatio) \ 100
            lHeight = CLng((CMobj(i).Height) * yRatio) \ 100
        Else
            lWidth = CLng((CMobj(i).Width * xRatio) \ 100)
            lHeight = CLng((CMobj(i).Height * yRatio) \ 100)
        End If
        
    lFontSize = CLng((CMobj(i).FontSize * yRatio) \ 100)
    
        If TypeOf FResize Is Line Then

                If FResize.X1 < 0 Then
                    FResize.X1 = CLng(((CMobj(i).Left * xRatio) \ 100) - 75000)
                Else
                    FResize.X1 = CLng((CMobj(i).Left * xRatio) \ 100)
                End If

            FResize.Y1 = CLng((CMobj(i).Top * yRatio) \ 100)

                If FResize.X2 < 0 Then
                    FResize.X2 = CLng(((CMobj(i).Width * xRatio) \ 100) - 75000)
                Else
                    FResize.X2 = CLng((CMobj(i).Width * xRatio) \ 100)
                End If

            FResize.Y2 = CLng((CMobj(i).Height * yRatio) \ 100)

        Else
    
            'For this APP, don't fix font for list boxes
            'Add "TypeOf" line for listBox

                If TypeOf FResize Is Label Or _
            TypeOf FResize Is CommandButton Or _
            TypeOf FResize Is OptionButton Or _
            TypeOf FResize Is CheckBox Or _
            TypeOf FResize Is Frame Or _
            TypeOf FResize Is TextBox Or _
            TypeOf FResize Is ComboBox Or _
            TypeOf FResize Is ListBox Or _
            TypeOf FResize Is DriveListBox Then FResize.Font.Size = lFontSize

                If TypeOf FResize Is MSHFlexGrid Then
                    lFontSize = CLng((CMobj(i).FontSize * xRatio) \ 100)
                    FResize.Font.Size = lFontSize
                End If

                If TypeOf FResize Is ComboBox Then
                    FResize.Move lLeft, lTop, lWidth
                Else
                    On Error Resume Next
                    FResize.Move lLeft, lTop, lWidth, lHeight
                End If

            '********************************************************************
            '
            '   Resize text in RichTextBox
            '
            '   Preserve colors and formating (Bold, Underline, Italic)
            '
            '   Mark Mokoski
            '   09-JAN-2003
            '
            '*******************************************************************
            '        If TypeOf FResize Is RichTextBox Then
            '               'Temp file to store current RTB text
            '                FResize.SaveFile App.Path & "\" & "RTBtext.rtf"
            '                'Set the font size of the new window size
            '                FResize.Font.Size = lFontSize
            '                'Restore RTB text to new size RTB
            '                FResize.LoadFile App.Path & "\" & "RTBtext.rtf"
            '                'Delete the temp file
            '                Kill (App.Path & "\" & "RTBtext.rtf")
            '                'Select whole text of RTB
            '                FResize.SelStart = 0
            '                FResize.SelLength = Len(FResize.Text)
            '                'Find "normal" screen stored font size
            '               OldSize = GetSetting("WA1ZEK", "AX25Monitor", "SysFont_size", 8)
            '                'Upon comming "down" in screen size, restore font size to stored one
            '                If OldSize > lFontSize Then lFontSize = OldSize
            '                'Change the font sizt of current RTB
            '                FResize.SelFontSize = lFontSize
            '                'For this APP, set font size constants
            '                InfoFont_Size = lFontSize
            '                MyFont_Size = lFontSize
            '                DataFont_size = lFontSize
            '                SysFont_Size = lFontSize
            '                'For this APP, Set cursor and scroll bars to end of RTB
            '                FResize.SelStart = 0
            '                FResize.SelStart = Len(FResize.Text)
            '        End If
            '*** End Added code by Mark Mokoski
        End If

End Sub

Private Function FindForm(ByRef FormToResize As Form) As Long

    Dim i                         As Long

    FindForm = -1

        If FormMax > 0 Then

                For i = 0 To (FormMax - 1)

                        If FormRecord(i).Name = FormToResize.Name Then
                            FindForm = i
                            Exit Function
                        End If

                Next i

        End If

End Function

Private Function AddForm(ByRef FormToResize As Form) As Long

    Dim FormControl               As Control
    Dim i                         As Long

    ReDim Preserve FormRecord(FormMax + 1)
    FormRecord(FormMax).Name = FormToResize.Name
    FormRecord(FormMax).Top = FormToResize.Top
    FormRecord(FormMax).Left = FormToResize.Left
    FormRecord(FormMax).Height = FormToResize.Height
    FormRecord(FormMax).Width = FormToResize.Width
    FormRecord(FormMax).ScaleHeight = FormToResize.ScaleHeight
    FormRecord(FormMax).ScaleWidth = FormToResize.ScaleWidth
    AddForm = FormMax
    FormMax = FormMax + 1

        For Each FormControl In FormToResize
            i = FindControl(FormControl, FormToResize.Name)

                If i < 0 Then i = AddControl(FormControl, FormToResize.Name)
        Next FormControl

    Set FormControl = Nothing
End Function

Private Function FindControl(ByVal FResize As Control, ByVal sName As String) As Long

    Dim i                         As Long

    FindControl = -1

        For i = 0 To (ControlMax - 1)

                If CMobj(i).Parent = sName Then

                        If CMobj(i).Name = FResize.Name Then
                            On Error Resume Next

                                If CMobj(i).Index = FResize.Index Then
                                    FindControl = i
                                    Exit Function
                                End If

                            On Error GoTo 0
                        End If

                End If

        Next i

End Function

Private Function AddControl(ByRef FResize As Control, ByVal sName As String) As Long

    ReDim Preserve CMobj(ControlMax + 1)
    On Error Resume Next
    CMobj(ControlMax).Name = FResize.Name
    CMobj(ControlMax).Index = FResize.Index
    CMobj(ControlMax).Parent = sName

        If TypeOf FResize Is Line Then
            CMobj(ControlMax).Top = FResize.Y1
            CMobj(ControlMax).Left = LeftPos(FResize.X1)
            CMobj(ControlMax).Height = FResize.Y2
            CMobj(ControlMax).Width = LeftPos(FResize.X2)
        Else
            CMobj(ControlMax).Top = FResize.Top
            CMobj(ControlMax).Left = LeftPos(FResize.Left)
            CMobj(ControlMax).Height = FResize.Height
            CMobj(ControlMax).Width = FResize.Width
            CMobj(ControlMax).FontSize = FResize.Font.Size
        End If

    FResize.IntegralHeight = False
    On Error GoTo 0
    AddControl = ControlMax
    ControlMax = ControlMax + 1

End Function

Private Function GetWidthRatio(ByRef FormToResize As Form) As Long

    Dim i                         As Long

    i = FindForm(FormToResize)

        If i < 0 Then i = AddForm(FormToResize)
    GetWidthRatio = (FormToResize.ScaleWidth * 100) \ FormRecord(i).ScaleWidth

End Function

Private Function GetHeightRatio(ByRef FormToResize As Form) As Single

    Dim i                         As Long

    i = FindForm(FormToResize)

        If i < 0 Then i = AddForm(FormToResize)
    GetHeightRatio = (FormToResize.ScaleHeight * 100) \ FormRecord(i).ScaleHeight

End Function

Private Function LeftPos(ByVal lLeftPosition As Long) As Long

        If lLeftPosition < 0 Then
            LeftPos = lLeftPosition + 75000
        Else
            LeftPos = lLeftPosition
        End If

End Function

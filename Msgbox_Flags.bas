Attribute VB_Name = "MsgBoxFlags"
Option Explicit

Public Enum MsgBox_Flags 'Define for custom skinned message box form
    '########### Button Combinations ###############
    vbOKOnly = 0 '          &H0&  OK button only (default)
    vbOKCancel = 1 '        &H1&  OK and Cancel buttons
    vbAbortRetryIgnore = 2 '&H2&  Abort, Retry, and Ignore buttons
    vbYesNoCancel = 3 '     &H3&  Yes, No, and Cancel buttons
    vbYesNo = 4 '           &H4&  Yes and No buttons
    vbRetryCancel = 5 '     &H5&  Retry and Cancel buttons
    vbCustomButtons = &H6& 'THIS IS YOUR OWN CHOICE OF BUTTON CAPTIONS
    '########### Icon Types available '##########
    vbCritical = 16 '&H10&        Critical message
    vbQuestion = 32 '&H20&        Warning query
    vbExclamation = 48 '&H30&     Warning message
    vbInformation = 64 '&H40&     Information message
    vbUserIcon = &H50&
    vbSecurityIcon = &H60&
    vbFindIcon = &H70&
    '###### Default Selected button #####
    vbDefaultButton1 = 0 'First button is default (default)
    vbDefaultButton2 = 256 ' Second button is default
    vbDefaultButton3 = 512 'Third button is default
End Enum


















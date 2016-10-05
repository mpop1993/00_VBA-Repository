Attribute VB_Name = "genericFct"
''' ---------------------------------------------------------------------------------------------
' checkTemplate     ' Check wether this document is the template or nor
' checkAge          ' Check if this is the first initialization of the document
' logEvent          ' Send data logs to the master LOG file
'
'
''' ---------------------------------------------------------------------------------------------

''' ---------------------------------------------------------------------------------------------
''' NAME: checkTemplate
''' DESCRIPTION: This functions checks if the current path is the same as the template path.
''' EXAMPLE:
'''
''' GLOBALS:    CHECK_TEMPLATE - boolean
'''             TEMPLATE_PATH - string
''' ---------------------------------------------------------------------------------------------
Public Sub checkTemplate()

    Dim sCurrentPath As String
    
On Error GoTo Handler

    If CHECK_TEMPLATE Then
        ' get current path
        sCurrentPath = Application.ActiveWorkbook.Path
        
        ' check if current path is template path
        If sCurrentPath = TEMPLATE_PATH Then
            'MsgBox "The file you are trying to open is a template." & vbNewLine & vbNewLine & "Please make a local copy on your computer in order to use it or log as administrator!"
            'Application.ActiveWorkbook.Save
            'Application.ActiveWorkbook.Close
            genericFormAdmin.Show
        End If
    End If
    
    Exit Sub
    
Handler:

    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Error"
    
End Sub

''' ---------------------------------------------------------------------------------------------
''' NAME: checkAge
''' DESCRIPTION: This functions checks if this is the first initialization or not.
''' EXAMPLE:
'''
''' GLOBALS:    CHECK_AGE - boolean
'''             SHEET_NAME_CONFIG - string
''' ---------------------------------------------------------------------------------------------
Public Function checkAge() As Boolean

On Error GoTo Handler
    
    If CHECK_AGE Then
        If Sheets(SHEET_NAME_CONFIG).Range(CONFIG_AGE).Value Then
            ' Print first initailization message
            MsgBox "This is the first initialization of the template. Please check ReadMe for more information."
            ' Set value to false for future usage
            Sheets(SHEET_NAME_CONFIG).Range(CONFIG_AGE).Value = False
            ' Return true
            checkAge = True
            ' Exit function
            Exit Function
        End If
    End If
    
    checkAge = False
    
    Exit Function
    
Handler:

    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Error"
    
End Function

''' ---------------------------------------------------------------------------------------------
''' NAME:       logEvent(string)
''' DESCRIPT:   This function is used to log a specific event in the shared log file
''' EXAMPLE:    logEvent("New entry - do xxxx")
'''
''' GLOBALS:    LOG_FILE - string - the actual path + file
'''             LOG_SHEET - string - the sheet name
''' ---------------------------------------------------------------------------------------------
Public Sub logEvent(sEvent As String)

    Dim wbLog As Workbook
    Dim wsLog As Worksheet
    Dim iFirstEmpty As Integer
    
On Error GoTo Handler
    Application.ScreenUpdating = False
    
    Set wbLog = Workbooks.Open(LOG_FILE)
    Set wsLog = wbLog.Worksheets(LOG_SHEET)
    
    ' get the first empty cell
    iFirstEmpty = wsLog.Cells(Rows.Count, 1).End(xlUp).Row + 2
    
    wsLog.Cells(iFirstEmpty, 1).Value = Environ$("username")
    wsLog.Cells(iFirstEmpty, 2).Value = Environ$("computername")
    wsLog.Cells(iFirstEmpty, 3).Value = Date
    wsLog.Cells(iFirstEmpty, 4).Value = Time
    wsLog.Cells(iFirstEmpty, 5).Value = sEvent
    
    wbLog.Save
    wbLog.Close
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
Handler:

    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Error"
    
End Sub

''' ---------------------------------------------------------------------------------------------
''' NAME:
''' DESCRIPT:
''' EXAMPLE:
'''
''' GLOBALS:
'''
''' ---------------------------------------------------------------------------------------------
Public Sub insertRow(iReferenceRow As Integer, iHeight As Integer)

End Sub












Sub Mail_workbook_Outlook_1()
'Working in Excel 2000-2016
'This example send the last saved version of the Activeworkbook
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .to = "mihai.pop@eltrex.ro"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .Body = "Hi there"
        .Attachments.Add ActiveWorkbook.FullName
        'You can add other files also like this
        '.Attachments.Add ("C:\test.txt")
        .Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


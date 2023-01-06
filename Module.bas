Attribute VB_Name = "Module"
Option Explicit

Sub Main()

    ' Show the Wait Screen
    shWait.Visible = xlSheetVisible
    shWait.Activate

    ' Set Main variables and Objects
    Const url$ = "https://www.golf.org.au/login"
    Const scoresUrl$ = "https://www.golf.org.au/member/dashboard"
    
    Dim username As String, password As String
    username = shData.Range("username").Value
    password = shData.Range("password").Value

    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    'ie.Visible = True
    
    ' Login to Website and navigate to main dashboard / scores page
    Login ie, url, username, password

    ' Scrape table data
    ScrapeData ie, scoresUrl
    
    ' Hide the Wait Screen and move to the Output Screen
    shWait.Visible = xlSheetHidden
    shOut.Activate
    
    ' Clean up Objects
    Set ie = Nothing

End Sub

Sub ScrapeData(ByRef ie As Object, scoresUrl As String)

    ' Navigate to the login page
    ie.navigate scoresUrl
    IeBusy ie
    LocalSleep 20
    
    Dim html As String
    html = ie.document.DocumentElement.outerHTML
    
    ' Extract the scores data from the HTML using regular expressions
    Dim scores As Object
    Set scores = CreateObject("VBScript.RegExp")
    scores.Pattern = "kgXSLa" & Chr(34) & ">(.*?)<"
    scores.Global = True
    
    ' Execute the RegEx to extract the scores
    Dim matches As Object
    Set matches = scores.Execute(html)

    ' Clear existing sheet content and populate with new content
    shOut.Cells.ClearContents

    shOut.Cells(1, 1).Value = "Hcp Score"
    shOut.Cells(1, 2).Value = "Daily Difficulty"
    shOut.Cells(1, 3).Value = "Scratch Rating"
    shOut.Cells(1, 4).Value = "Slope Rating"
    shOut.Cells(1, 5).Value = "Par"
    shOut.Cells(1, 6).Value = "Daily Handicap"
    shOut.Cells(1, 7).Value = "Adjusted Gross"
    shOut.Cells(1, 8).Value = "Gross Diff"
    shOut.Cells(1, 9).Value = "New GA Handicap"
    
    Dim i, j, idx As Integer
    idx = 0
    
    For i = 2 To 21 ' Get the last 20 matches
        For j = 0 To 8 ' 9 columns in the table
            idx = (i - 2) * 9 + j
            shOut.Cells(i, j + 1).Value = Mid(matches(idx), 9, Len(matches(idx)) - 9)
        Next j
    Next i
    
    ' Clean up Objects
    Set scores = Nothing
    Set matches = Nothing

End Sub

Sub Login(ByRef ie As Object, url As String, username As String, password As String)

    Dim waitSeconds As Integer
    waitSeconds = 2
    
    With ie

        ' Navigate to the login page
        .navigate url
        IeBusy ie

        ' Set the object and variats to get the login page fields
        Dim oInputFields As Object, oButtons As Object
        Dim oField As Variant, oLogin As Object, oPassword As Object
        LocalSleep waitSeconds
        
        ' Get the login form fields and button
        Set oInputFields = .document.getElementsByClassName("MuiInputBase-input MuiInput-input")
        Set oButtons = .document.getElementsByClassName("MuiButtonBase-root MuiButton-root MuiButton-text Buttonsstyle__Button-yum10d-0 egMYaS SubmitButtonstyle__StyledButton-sc-92coke-1 ceaovC")
        LocalSleep waitSeconds
        
        ' Enter Username and Password into the form
        For Each oField In oInputFields
            If oField.ID = "username" Then
                InputToField oField, username
            End If
            If oField.ID = "password" Then
                InputToField oField, password
            End If
        Next oField
        LocalSleep waitSeconds
        
        ' Click the Login Button
        oButtons(0).Click

    End With
    
    ' Clean up Objects
    Set oInputFields = Nothing
    Set oButtons = Nothing
    

End Sub

Sub IeBusy(ie As Object)
    Do While ie.Busy Or ie.readyState < 4
        DoEvents
    Loop
End Sub

Function InputToField(field As Variant, text As String)
    field.Focus
    field.innertext = text
    'oField.Value = text
    LocalSleep 2
End Function

Function LocalSleep(waitSeconds As Integer)
    Application.Wait (Now + TimeValue("0:00:0" & waitSeconds))
End Function

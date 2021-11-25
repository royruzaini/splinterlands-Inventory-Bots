Sub InventoryTransfer()

    Dim driver As New WebDriver
    Dim rowc, cc, columnC As Integer
    Dim by As New Selenium.by
        
'Find CH Unfulfilled
    Dim CH As Workbook
    Dim s As Workbook
 
    For Each s In Workbooks
    If Left(s.Name, 13) = "InventoryBots" Then
        Set CH = Workbooks(s.Name)
    End If
    Next s
        
'login
    driver.Start "edge"
    driver.Timeouts.ImplicitWait = 10000 ' 10 seconds
        
    Dim ws As Worksheet
    Set ws = CH.Worksheets("User")
    
    Dim username As Variant
    Dim Password As Variant
    Dim Activekey As Variant
    
    Dim lRow As Long
    lRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 1 To lRow

    MainAccount = ws.Range("A1").Value
    declimit = ws.Range("G1").Value
    username = ws.Range("A" & i).Value
    Password = ws.Range("B" & i).Value
    Activekey = ws.Range("C" & i).Value
    
    Application.Wait Now + TimeValue("00:01:00")
    Call login(driver, by, username, Password)
    
'claim SPS
    Call claim_sps(driver, by)
'ssend sps
    If i <> 1 Then
    Call send_sps(driver, by, MainAccount, username, Activekey)
    End If
'send dec
    If i <> 1 Then
    Call send_dec(driver, by, MainAccount, username, Activekey, declimit)
    End If
'transferCard
    If i <> 1 Then
    Call transfercards(driver, by, MainAccount, username, Activekey)
    End If
    
    Next i

    driver.Quit
         
End Sub

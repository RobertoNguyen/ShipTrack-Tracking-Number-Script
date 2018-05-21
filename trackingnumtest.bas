Attribute VB_Name = "Module1"
Public ieTrack As InternetExplorer
Public aShipTrackCache() As Variant
Option Explicit
Option Compare Text
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub testShipTrack()
    Dim vTracking As Variant
    Dim sCarrier As String
    Dim sReturn As String
    Dim a() As Variant
    Dim i As Long
   
    vTracking = "1872148784"
    sCarrier = "DHL"
    sReturn = "Status"
    'a = ShipTrackScraper(vTracking, sCarrier)
   'For i = LBound(a) To UBound(a)
   '    Debug.Print "-" & a(i) & "-"
   'Next i
    Debug.Print ""
    Debug.Print "Tracking#- " & ShipTrack(vTracking, sCarrier, "Tracking") & ""
    Debug.Print "Status- " & ShipTrack(vTracking, sCarrier, "Status") & ""
    Debug.Print "Delivered- " & ShipTrack(vTracking, sCarrier, "Delivered") & ""
    Debug.Print "RecBy- " & ShipTrack(vTracking, sCarrier, "RecBy") & ""
    Debug.Print "ShipTo- " & ShipTrack(vTracking, sCarrier, "ShipTo") & ""
    Debug.Print "ServiceLvl- " & ShipTrack(vTracking, sCarrier, "ServiceLvl") & ""
    Debug.Print "Origin- " & ShipTrack(vTracking, sCarrier, "Origin") & ""
    Debug.Print "Manifest- " & ShipTrack(vTracking, sCarrier, "Manifest") & ""
    Debug.Print "Scheduled- " & ShipTrack(vTracking, sCarrier, "Scheduled") & ""
    Debug.Print "TimeStamp- " & ShipTrack(vTracking, sCarrier, "TimeStamp") & ""
   
 
End Sub
 
 
Function ShipTrack(vTracking As Variant, sCarrier As String, sReturn As String, Optional bRefresh As Boolean = False) As Variant
    Dim i As Long
    Dim iRow As Long
    Dim bArchived As Boolean
    Dim sStatus As String
    Dim vDelDate As Variant
    Dim sRecBy As String
    Dim sShipTo As String
    Dim sSerLvl As String
    Dim vOrgDate As Variant
    Dim vManDate As Variant
    Dim vSchDate As Variant
    Dim dblTimeStamp As Double
    Dim aShipTrackScrape() As Variant
    Dim dblStart As Double
    Dim dblEnd As Double
   
    'Check if array already exists
   On Error Resume Next
    i = UBound(aShipTrackCache, 2)
    If Err.Number <> 0 Then
        ReDim aShipTrackCache(0 To 9, 0 To 0)
    End If
    On Error GoTo 0
   
    'aShipTrackCache(0, 0) = vTracking
   
    'Check if this ESN is already archived
   For i = LBound(aShipTrackCache, 2) To UBound(aShipTrackCache, 2)
        If vTracking = aShipTrackCache(0, i) Then
            iRow = i
            bArchived = True
            Exit For
        End If
    Next i
   
    'Archived and no refresh? Print old results
   If bRefresh = False And bArchived = True Then
        sStatus = aShipTrackCache(1, iRow)
        vDelDate = aShipTrackCache(2, iRow)
        sRecBy = aShipTrackCache(3, iRow)
        sShipTo = aShipTrackCache(4, iRow)
        sSerLvl = aShipTrackCache(5, iRow)
        vOrgDate = aShipTrackCache(6, iRow)
        vManDate = aShipTrackCache(7, iRow)
        vSchDate = aShipTrackCache(8, iRow)
        dblTimeStamp = aShipTrackCache(9, iRow)
        GoTo PrintInfo
    End If
   
    'If it's not archived, find out what row we're on
   If bArchived = False Then
        If UBound(aShipTrackCache, 2) = 0 And aShipTrackCache(0, 0) = "" Then
            iRow = 0
        Else
            iRow = UBound(aShipTrackCache, 2) + 1
            ReDim Preserve aShipTrackCache(9, iRow)
        End If
    End If
   
    'Scrape data
   aShipTrackScrape = ShipTrackScraper(vTracking, sCarrier)
   
    'Exit if bad data
   If aShipTrackScrape(0) = "Page Not Found" Or aShipTrackScrape(0) = "Bad Tracking #" Then
        ShipTrack = aShipTrackScrape(0)
        Exit Function
    End If
   
    'Update Array
   dblTimeStamp = Now()
    aShipTrackCache(0, iRow) = vTracking
    aShipTrackCache(1, iRow) = aShipTrackScrape(0)
    aShipTrackCache(2, iRow) = aShipTrackScrape(1)
    aShipTrackCache(3, iRow) = aShipTrackScrape(2)
    aShipTrackCache(4, iRow) = aShipTrackScrape(3)
    aShipTrackCache(5, iRow) = aShipTrackScrape(4)
    aShipTrackCache(6, iRow) = aShipTrackScrape(5)
    aShipTrackCache(7, iRow) = aShipTrackScrape(6)
    aShipTrackCache(8, iRow) = aShipTrackScrape(7)
    aShipTrackCache(9, iRow) = dblTimeStamp
   
PrintInfo::
    Select Case sReturn
        Case Is = "Tracking"
            ShipTrack = aShipTrackCache(0, iRow)
        Case Is = "Status"
            ShipTrack = aShipTrackCache(1, iRow)
        Case Is = "Delivered"
            ShipTrack = aShipTrackCache(2, iRow)
        Case Is = "RecBy"
            ShipTrack = aShipTrackCache(3, iRow)
        Case Is = "ShipTo"
            ShipTrack = aShipTrackCache(4, iRow)
        Case Is = "ServiceLvl"
            ShipTrack = aShipTrackCache(5, iRow)
        Case Is = "Origin"
            ShipTrack = aShipTrackCache(6, iRow)
        Case Is = "Manifest"
            ShipTrack = aShipTrackCache(7, iRow)
        Case Is = "Scheduled"
            ShipTrack = aShipTrackCache(8, iRow)
        Case Is = "TimeStamp"
            ShipTrack = aShipTrackCache(9, iRow)
    End Select
   
    Exit Function
   
ErrHandler:
    Debug.Print Err.Description
    ShipTrack = "Error"
    ieTrack.Quit
    Set ieTrack = Nothing
    Exit Function
End Function
 
Private Function ShipTrackScraper(vTracking As Variant, sCarrier As String) As Variant
    Dim ieTagP As IHTMLElementCollection
    Dim ieTagTD As IHTMLElementCollection
    Dim ieTagDL As IHTMLElementCollection
    Dim ieTagFieldSet As IHTMLElementCollection
    Dim ieTagDD As IHTMLElementCollection
    Dim ieTagDiv As IHTMLElementCollection
    Dim ieTagClass As IHTMLElementCollection
    Dim ieTagTR As IHTMLElementCollection
    Dim ieTagH3 As IHTMLElementCollection
    
    Dim tableCell As IHTMLElementCollection
    
    Dim sURL As String
    Dim sCarrierConfirm As String
    Dim sCheck As String
    Dim i As Long
    Dim iTimeOut As Long
    Dim sTrck As String
    Dim bCheck As Boolean
    Dim bPageFound As Boolean
    Dim bBadTracking As Boolean
    Dim sStatus As String
    Dim sDelDate As String
    Dim sDelTime As String
    Dim vDelDate As Variant
    Dim sRecBy As String
    Dim sShipTo As String
    Dim sSerLvl As String
    Dim sOrgDate As String
    Dim sOrgTime As String
    Dim vOrgDate As Variant
    Dim sManDate As String
    Dim sManTime As String
    Dim vManDate As Variant
    Dim dblWeight As Double
    Dim sSchDate As String
    Dim aSchDate() As String
    Dim vSchDate As Variant
    Dim aTrackData() As Variant
    Dim sError As String
    Dim j As Long
    Dim bTagFound As Boolean
 
    'Some constants
   iTimeOut = 10
    ReDim aTrackData(0 To 7)
 
    'What carrier are we using? Exit if unknown
   If sCarrier = "UPS" Then
        sURL = "https://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=" & vTracking
        sCarrierConfirm = "UPS"
    ElseIf sCarrier = "FedEx" Then
        sURL = "https://www.fedex.com/apps/fedextrack/?tracknumbers=" & vTracking
        sCarrierConfirm = "FedEx"
    ElseIf sCarrier = "DHL" Then
        sURL = "http://www.dhl.com/en/express/tracking.html?AWB=" & vTracking
        sCarrierConfirm = "DHL"
    Else
        sError = "Unknown Carrier"
        GoTo ErrExit
        Exit Function
    End If
   
    'Is IE already open? Awesome. Otherwise open it
   On Error Resume Next
    sCheck = ieTrack.LocationURL
    On Error GoTo 0
    If sCheck = "" Then
        Set ieTrack = CreateObject("InternetExplorer.Application")
    End If
   
    'Open URL
   With ieTrack
    .navigate sURL
    ieTrack.Visible = True
    Do While (.Busy Or .READYSTATE <> READYSTATE_COMPLETE): DoEvents: Loop
    bPageFound = False
    bBadTracking = False
       
        '===========================================
       'UPS
       '===========================================
       If sCarrierConfirm = "UPS" Then
            'Make sure we're on the right page
           bPageFound = False
            For i = 1 To iTimeOut
                Set ieTagH3 = .document.getElementsByTagName("H3")
                For j = 0 To ieTagH3.Length - 1
                    If ieTagH3(j).innerText Like "*" & Trim(vTracking) & "*" Then
                        bPageFound = True
                        Exit For
                    End If
                Next j
                If bPageFound = True Then
                    Exit For
                End If
                Sleep (100)
            Next i
           
            'Page not found?
           If bPageFound = False Then
                'Bad Tracking Number?
               Set ieTagP = .document.getElementsByTagName("p")
                For i = 0 To ieTagP.Length - 1
                    If ieTagP(i).innerText Like "*The number you entered is not a valid tracking number*" Then
                        bBadTracking = True
                        Exit For
                    End If
                Next i
                If bBadTracking = True Then
                    sError = "Bad Tracking #"
                    GoTo ErrExit
                    Exit Function
                End If
               
                'Catch all for page just not working for some reason...
               sError = "UPS - Page Not Found"
                GoTo ErrExit
                Exit Function
            End If
           
            'We're on the right page, lets get the dataz...
           On Error Resume Next
           
            'Current Status
           sStatus = Trim(.document.getElementById("tt_spStatus").innerText)
           
            'Delivery Date
           Set ieTagTD = .document.getElementsByTagName("TD")
            For i = 0 To ieTagTD.Length - 1
                If Trim(ieTagTD(i).innerText) = "Delivered" Then
                    sDelDate = Trim(ieTagTD(i - 2).innerText)
                    sDelTime = Replace(Trim(ieTagTD(i - 1).innerText), ".", "")
                    vDelDate = CDate(sDelDate & " " & sDelTime)
                    Exit For
                End If
            Next i
           
            'Received By
           Set ieTagP = .document.getElementsByTagName("P")
            For i = 0 To ieTagP.Length - 1
                If Trim(ieTagP(i).innerText) Like "Received By:*" Then
                    sRecBy = ieTagP(i + 1).innerText
                    Exit For
                End If
            Next i
           
            'Shipped To
           Set ieTagDL = .document.getElementsByTagName("Address")
            sShipTo = Trim(WorksheetFunction.Clean(ieTagDL(0).innerText))
           
            'Service Level
           Set ieTagFieldSet = .document.getElementsByTagName("div")
            For i = 0 To ieTagFieldSet.Length - 1
                If Trim(ieTagFieldSet(i).innerText) Like "Service*" Then
                    sSerLvl = Replace(Trim(ieTagFieldSet(i).innerText), "Service", "")
                    sSerLvl = WorksheetFunction.Clean(sSerLvl)
                    Exit For
                End If
            Next i
           
            'Pickup Date
           Set ieTagTD = .document.getElementsByTagName("TD")
            For i = 0 To ieTagTD.Length - 1
                If Trim(WorksheetFunction.Clean(ieTagTD(i).innerText)) = "Origin Scan" Then
                    sOrgDate = Trim(ieTagTD(i - 2).innerText)
                    sOrgTime = Replace(Trim(ieTagTD(i - 1).innerText), ".", "")
                    vOrgDate = CDate(sOrgDate & " " & sOrgTime)
                    Exit For
                End If
            Next i
           
            'Manifest Date
           Set ieTagTD = .document.getElementsByTagName("TD")
            For i = 0 To ieTagTD.Length - 1
                If Trim(WorksheetFunction.Clean(ieTagTD(i).innerText)) = "Order Processed: Ready for UPS" Then
                    sManDate = Trim(ieTagTD(i - 2).innerText)
                    sManTime = Replace(Trim(ieTagTD(i - 1).innerText), ".", "")
                    vManDate = CDate(sManDate & " " & sManTime)
                    Exit For
                End If
            Next i
           
            'Weight
           Set ieTagDL = .document.getElementsByTagName("DL")
            For i = 0 To ieTagDL.Length - 1
                If Trim(ieTagDL(i).innerText) Like "Scheduled Delivery:*" Then
                    sSchDate = Replace(Trim(ieTagDL(i).innerText), "Scheduled Delivery:", "")
                    sSchDate = WorksheetFunction.Clean(sSchDate)
                    aSchDate = Split(sSchDate, ",")
                    vSchDate = CDate(aSchDate(1))
                    Exit For
                End If
            Next i
           
            On Error GoTo 0
           
        '===========================================
       'FedEx
       '===========================================
       ElseIf sCarrierConfirm = "FedEx" Then
         'Make sure we're on the right page
        For i = 1 To iTimeOut
             Set ieTagDiv = .document.getElementsByTagName("Div")
             For j = 0 To ieTagDiv.Length - 1
                 If ieTagDiv(j).innerText = vTracking Then
                     bCheck = True
                     Exit For
                 End If
             Next j
             On Error GoTo 0
             If bCheck = True Then
                 bPageFound = True
                 Exit For
             End If
             Sleep (250)
         Next i
        
         'Page not found?
        If bPageFound = False Then
             'Bad Tracking Number?
            Set ieTagDiv = .document.getElementsByTagName("Div")
             For i = 0 To ieTagDiv.Length - 1
                 If Trim(ieTagDiv(i).innerText) = "Not Found" Then
                     bBadTracking = True
                     Exit For
                 End If
             Next i
             If bBadTracking = True Then
                 sError = "Bad Tracking #"
                 GoTo ErrExit
                 Exit Function
             End If
            
             'Catch all for page just not working for some reason...
             sError = "FedEx - Page Not Found"
            GoTo ErrExit
            Exit Function
         End If
        
         'We're on the right page, lets get the dataz...
        On Error Resume Next
    
         'Current Status
        sStatus = .document.getElementsByClassName("statusChevron_key_status bogus")(0).innerText
        
         'Delivery Date
        sDelDate = .document.getElementsByClassName("snapshotController_date dest")(0).innerText
         sDelDate = Right(sDelDate, Len(sDelDate) - 4)
         vDelDate = CDate(sDelDate)
        
         'Received By
        sRecBy = .document.getElementsByClassName("statusChevron_sub_status bogus")(0).innerText
         If sRecBy Like "Signed for by:*" Then
             sRecBy = Replace(sRecBy, "Signed for by: ", "")
         Else
             sRecBy = ""
         End If
        
         'Shipped To
        sShipTo = .document.getElementsByClassName("address_cscp")(1).innerText
        
         'Service Level
        Set ieTagTD = .document.getElementsByTagName("TD")
         For i = 0 To ieTagTD.Length - 1
             If Trim(WorksheetFunction.Clean(ieTagTD(i).innerText)) = "Service" Then
                 sSerLvl = Trim(ieTagTD(i + 1).innerText)
                 Exit For
             End If
         Next i
        
         'Pickup Date
        sOrgDate = .document.getElementsByClassName("snapshotController_date orig")(0).innerText
         sOrgDate = Right(sOrgDate, Len(sOrgDate) - 4)
         vOrgDate = CDate(sOrgDate)
        
         'Pickup Date
        'Set ieTagTR = .document.getElementsByTagName("TR")
        'For i = 0 To ieTagTR.Length - 1
        '    If ieTagTR(i).innerHTML Like "*travel_history_header_date*" Then
        '        sOrgDate = WorksheetFunction.Clean(Trim(ieTagTR(i).innerText))
        '        vOrgDate = CDate(Left(sOrgDate, 10))
        '    End If
        '    If ieTagTR(i).innerText Like "*Picked up*" Then
        '        bTagFound = True
        '        Exit For
        '    End If
        'Next i
        'If bTagFound = False Then
        '    vOrgDate = ""
        'End If
        
         'Manifest Date
        bTagFound = False
         Set ieTagTR = .document.getElementsByTagName("TR")
         For i = 0 To ieTagTR.Length - 1
             If ieTagTR(i).innerHTML Like "*travel_history_header_date*" Then
                 sManDate = WorksheetFunction.Clean(Trim(ieTagTR(i).innerText))
                 vManDate = CDate(Left(sManDate, 10))
             End If
             If ieTagTR(i).innerText Like "*Shipment information sent to FedEx*" Then
                 bTagFound = True
                 Exit For
             End If
         Next i
         If bTagFound = False Then
             vManDate = ""
         End If
        
         'Delivery Date
        vSchDate = CDate(Left(sDelDate, 10))
        
        
        '===========================================
       'DHL
       '===========================================
       Else
         'Make sure we're on the right page
        For i = 1 To iTimeOut
             Set ieTagDiv = .document.getElementsByTagName("strong")
             For j = 0 To ieTagDiv.Length - 1
                 If ieTagDiv(j).innerText = "Waybill: " + vTracking Then
                     bCheck = True
                     Exit For
                 End If
             Next j
             On Error GoTo 0
             If bCheck = True Then
                 bPageFound = True
                 Exit For
             End If
             Sleep (250)
         Next i
        
         'Page not found?
        If bPageFound = False Then
             'Bad Tracking Number?
            Set ieTagDiv = .document.getElementsByTagName("Div")
             For i = 0 To ieTagDiv.Length - 1
                 If Trim(ieTagDiv(i).innerText) = "Not Found" Then
                     bBadTracking = True
                     Exit For
                 End If
             Next i
             If bBadTracking = True Then
                 sError = "Bad Tracking #"
                 GoTo ErrExit
                 Exit Function
             End If
            
             'Catch all for page just not working for some reason...
             sError = "DHL - Page Not Found"
            GoTo ErrExit
            Exit Function
         End If
        
         'We're on the right page, lets get the dataz...
        On Error Resume Next
    
         'Current Status
        Set tableCell = .document.getElementsByClassName("result-summary result-has-pieces").getElementsByTagName("tbody")(0)
        'Debug.Print tableCell
        sStatus = tableCell.getElementsByTagName("title")
        
        'sStatus = .document.getElementsByClassName("result-summary result-has-pieces").innerText
        
         'Delivery Date
        sDelDate = .document.getElementsByClassName("snapshotController_date dest")(0).innerText
         sDelDate = Right(sDelDate, Len(sDelDate) - 4)
         vDelDate = CDate(sDelDate)
        
         'Received By
        sRecBy = .document.getElementsByClassName("statusChevron_sub_status bogus")(0).innerText
         If sRecBy Like "Signed for by:*" Then
             sRecBy = Replace(sRecBy, "Signed for by: ", "")
         Else
             sRecBy = ""
         End If
        
         'Shipped To
        sShipTo = .document.getElementsByClassName("address_cscp")(1).innerText
        
         'Service Level
        Set ieTagTD = .document.getElementsByTagName("TD")
         For i = 0 To ieTagTD.Length - 1
             If Trim(WorksheetFunction.Clean(ieTagTD(i).innerText)) = "Service" Then
                 sSerLvl = Trim(ieTagTD(i + 1).innerText)
                 Exit For
             End If
         Next i
        
         'Pickup Date
        sOrgDate = .document.getElementsByClassName("snapshotController_date orig")(0).innerText
         sOrgDate = Right(sOrgDate, Len(sOrgDate) - 4)
         vOrgDate = CDate(sOrgDate)
        
         'Pickup Date
        'Set ieTagTR = .document.getElementsByTagName("TR")
        'For i = 0 To ieTagTR.Length - 1
        '    If ieTagTR(i).innerHTML Like "*travel_history_header_date*" Then
        '        sOrgDate = WorksheetFunction.Clean(Trim(ieTagTR(i).innerText))
        '        vOrgDate = CDate(Left(sOrgDate, 10))
        '    End If
        '    If ieTagTR(i).innerText Like "*Picked up*" Then
        '        bTagFound = True
        '        Exit For
        '    End If
        'Next i
        'If bTagFound = False Then
        '    vOrgDate = ""
        'End If
        
         'Manifest Date
        bTagFound = False
         Set ieTagTR = .document.getElementsByTagName("TR")
         For i = 0 To ieTagTR.Length - 1
             If ieTagTR(i).innerHTML Like "*travel_history_header_date*" Then
                 sManDate = WorksheetFunction.Clean(Trim(ieTagTR(i).innerText))
                 vManDate = CDate(Left(sManDate, 10))
             End If
             If ieTagTR(i).innerText Like "*Shipment information sent to FedEx*" Then
                 bTagFound = True
                 Exit For
             End If
         Next i
         If bTagFound = False Then
             vManDate = ""
         End If
        
         'Delivery Date
        vSchDate = CDate(Left(sDelDate, 10))
        
        
        End If
    End With
 
    aTrackData(0) = sStatus
    aTrackData(1) = vDelDate
    aTrackData(2) = sRecBy
    aTrackData(3) = sShipTo
    aTrackData(4) = sSerLvl
    aTrackData(5) = vOrgDate
    aTrackData(6) = vManDate
    aTrackData(7) = vSchDate
 
    ShipTrackScraper = aTrackData
    
    'ieTrack.Quit
    'Set ieTrack = Nothing
    On Error GoTo 0
    Exit Function
   
ErrExit::
    aTrackData(0) = sError
    aTrackData(1) = sError
    aTrackData(2) = sError
    aTrackData(3) = sError
    aTrackData(4) = sError
    aTrackData(5) = sError
    aTrackData(6) = sError
    aTrackData(7) = sError
   
    ShipTrackScraper = aTrackData
    'ieTrack.Quit
    'Set ieTrack = Nothing
    On Error GoTo 0
End Function



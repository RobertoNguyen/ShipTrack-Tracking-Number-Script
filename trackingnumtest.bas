Attribute VB_Name = "Module1"
Public ieTrack As InternetExplorer
Public aShipTrackCache() As Variant
Option Explicit
Option Compare Text
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub runShipTrack()
    Dim vTracking As Variant
    Dim sCarrier As String
    Dim sReturn As String
    Dim a() As Variant
    Dim i As Long
   
    'vTracking = "1872148784" 'Delivered
    'vTracking = "7183073402" 'Delivered
    'vTracking = "3740172951" 'In Transit
    'vTracking = "4887545585" 'In Transit
    'vTracking = "4887538574"
    'vTracking = "4887545585" 'OFD
    sCarrier = "DHL"
    sReturn = "Status"
    
    'Debug.Print ""
    'Debug.Print "Status for " & vTracking & ": " & ShipTrack(vTracking, sCarrier, "Status") & ""
    
    For i = 6 To 24
        If Not IsEmpty(Cells(i, "D").Value) Then
            vTracking = Trim(Cells(i, "D").Value)
            Cells(i, "F") = ShipTrack(vTracking, sCarrier, sReturn, True)
            Debug.Print ""
            Debug.Print "Status for " & vTracking & ": " & ShipTrack(vTracking, sCarrier, "Status") & ""
        End If
    Next i
    
    ieTrack.Quit
    Set ieTrack = Nothing
 
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
    Dim ieTag As IHTMLElementCollection
    Dim sStatusElement As IHTMLElement
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
    'ieTrack.Visible = True
    ieTrack.Visible = False
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
                Set ieTag = .document.getElementsByTagName("H3")
                For j = 0 To ieTag.Length - 1
                    If ieTag(j).innerText Like "*" & Trim(vTracking) & "*" Then
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
               Set ieTag = .document.getElementsByTagName("p")
                For i = 0 To ieTag.Length - 1
                    If ieTag(i).innerText Like "*The number you entered is not a valid tracking number*" Then
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
           Set ieTag = .document.getElementsByTagName("TD")
            For i = 0 To ieTag.Length - 1
                If Trim(ieTag(i).innerText) = "Delivered" Then
                    sDelDate = Trim(ieTag(i - 2).innerText)
                    sDelTime = Replace(Trim(ieTag(i - 1).innerText), ".", "")
                    vDelDate = CDate(sDelDate & " " & sDelTime)
                    Exit For
                End If
            Next i
           
            'Received By
           Set ieTag = .document.getElementsByTagName("P")
            For i = 0 To ieTag.Length - 1
                If Trim(ieTag(i).innerText) Like "Received By:*" Then
                    sRecBy = ieTag(i + 1).innerText
                    Exit For
                End If
            Next i
           
            'Shipped To
           Set ieTag = .document.getElementsByTagName("Address")
            sShipTo = Trim(WorksheetFunction.Clean(ieTag(0).innerText))
           
            'Service Level
           Set ieTag = .document.getElementsByTagName("div")
            For i = 0 To ieTag.Length - 1
                If Trim(ieTag(i).innerText) Like "Service*" Then
                    sSerLvl = Replace(Trim(ieTag(i).innerText), "Service", "")
                    sSerLvl = WorksheetFunction.Clean(sSerLvl)
                    Exit For
                End If
            Next i
           
            'Pickup Date
           Set ieTag = .document.getElementsByTagName("TD")
            For i = 0 To ieTag.Length - 1
                If Trim(WorksheetFunction.Clean(ieTag(i).innerText)) = "Origin Scan" Then
                    sOrgDate = Trim(ieTag(i - 2).innerText)
                    sOrgTime = Replace(Trim(ieTag(i - 1).innerText), ".", "")
                    vOrgDate = CDate(sOrgDate & " " & sOrgTime)
                    Exit For
                End If
            Next i
           
            'Manifest Date
           Set ieTag = .document.getElementsByTagName("TD")
            For i = 0 To ieTag.Length - 1
                If Trim(WorksheetFunction.Clean(ieTag(i).innerText)) = "Order Processed: Ready for UPS" Then
                    sManDate = Trim(ieTag(i - 2).innerText)
                    sManTime = Replace(Trim(ieTag(i - 1).innerText), ".", "")
                    vManDate = CDate(sManDate & " " & sManTime)
                    Exit For
                End If
            Next i
           
            'Weight
           Set ieTag = .document.getElementsByTagName("DL")
            For i = 0 To ieTag.Length - 1
                If Trim(ieTag(i).innerText) Like "Scheduled Delivery:*" Then
                    sSchDate = Replace(Trim(ieTag(i).innerText), "Scheduled Delivery:", "")
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
             Set ieTag = .document.getElementsByTagName("Div")
             For j = 0 To ieTag.Length - 1
                 If ieTag(j).innerText = vTracking Then
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
            Set ieTag = .document.getElementsByTagName("Div")
             For i = 0 To ieTag.Length - 1
                 If Trim(ieTag(i).innerText) = "Not Found" Then
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
        Set ieTag = .document.getElementsByTagName("TD")
         For i = 0 To ieTag.Length - 1
             If Trim(WorksheetFunction.Clean(ieTag(i).innerText)) = "Service" Then
                 sSerLvl = Trim(ieTag(i + 1).innerText)
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
         Set ieTag = .document.getElementsByTagName("TR")
         For i = 0 To ieTag.Length - 1
             If ieTag(i).innerHTML Like "*travel_history_header_date*" Then
                 sManDate = WorksheetFunction.Clean(Trim(ieTag(i).innerText))
                 vManDate = CDate(Left(sManDate, 10))
             End If
             If ieTag(i).innerText Like "*Shipment information sent to FedEx*" Then
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
        Dim sTrackNum As String
         'Make sure we're on the right page
        For i = 1 To iTimeOut
             Set ieTag = .document.getElementsByTagName("strong")
             For j = 0 To ieTag.Length - 1
                sTrackNum = ieTag(j).innerText
                If sTrackNum = "Waybill: " + vTracking Then
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
            Set ieTag = .document.getElementsByTagName("Div")
             For i = 0 To ieTag.Length - 1
                 If Trim(ieTag(i).innerText) = "Not Found" Then
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
        Dim sTransKeyWords As String
        Dim sStringArray() As String
        Dim vWord As Variant
        Dim sOut4Del As String
        
        'Returns string results
        Set ieTag = .document.getElementsByTagName("td")
        sStatus = Trim(ieTag(0).Title)
        sOut4Del = .document.getElementsByTagName("tr")(0).innerHTML

        If sStatus Like "*delivered*" Then
            sStatus = "Delivered"
        ElseIf sOut4Del Like "*with delivery courier*" Then
            sStatus = "Out for Delivery"
        Else
            sTransKeyWords = Array("arrived", "departed", "processed", "clearance", "transferred")
            sStatus = .document.getElementsByTagName("tr")(0).innerText
            sStringArray = Split(Replace(Replace(sStatus, Chr(10), ""), Chr(13), ""))
            For Each vWord In sStringArray
                If IsInArray(vWord, sTransKeyWords) Then
                    sStatus = "In Transit"
                    Exit For
                End If
            Next vWord

        End If

        'Debug.Print sStatus
        
        '"In Transit": keywords to look for when results are parsed
        'sTransKeyWords = Array("with", "arrived", "departed", "processed", "clearance", "transferred")
        
        'Combines multilined result into 1 lined string and parses into array
        'sStringArray = Split(Replace(Replace(sStatus, Chr(10), ""), Chr(13), ""))
        
        'If sStatus Like "*Signed for by*" Then
        '    sStatus = "Delivered"
        'Else
        '    Checks if parsed string has keywords in ArrayList to check if in transit
        '    For Each vWord In sStringArray
        '        If IsInArray(vWord, sTransKeyWords) Then
        '            sStatus = "In Transit"
        '            Exit For
        '        End If
        '    Next vWord
        'End If
                
        
         'Delivery Date
        Dim iCommaPos As Integer
        Dim iAtPos As Integer
        Dim sDay As String
        Dim sTime As String
        Set ieTag = .document.getElementsByTagName("span")
        
        For i = 0 To ieTag.Length - 1
            'Debug.Print i & " - " & ieTag(i + 1).innerText
            If ieTag(i).innerText Like "*Proof of Delivery*" Or Not sStatus = "In Transit" And ieTag(i).innerText Like "Sign up for shipment notifications" Then
                sDelDate = ieTag(i + 1).innerText
                iCommaPos = InStr(sDelDate, ",")
                iAtPos = InStr(sDelDate, "at")
                'sDay = Left(sDelDate, Len(sDelDate) - iAtPos-1)
                sTime = Right(sDelDate, Len(sDelDate) - iAtPos - 2)
                sDay = Left(sDelDate, 3)
                sDelDate = Mid(Left(sDelDate, iAtPos - 2), iCommaPos + 2)
                
                If sStatus = "Delivered" Then
                    vDelDate = sDay & " " & CDate(sDelDate) & " " & Format(sTime, "h:mm AM/PM")
                Else
                    vDelDate = sDay & " " & CDate(sDelDate)
                End If
                    
                sStatus = sStatus & " - " & vDelDate
                Exit For
            ElseIf ieTag(i).innerText Like "Estimated Delivery:" And ieTag(i + 2).innerText Like "By End of Day" Then
                sDelDate = ieTag(i + 1).innerText
                iCommaPos = InStr(sDelDate, ",")
                'sDay = Left(sDelDate, iCommaPos - 1)
                sDay = Left(sDelDate, 3)
                sDelDate = Right(sDelDate, Len(sDelDate) - iCommaPos - 1)
                vDelDate = sDay & " " & CDate(sDelDate)
                sStatus = sStatus & " - " & vDelDate
                Exit For
            Else
                vDelDate = ""
            End If
        Next i
        
        
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
        Set ieTag = .document.getElementsByTagName("TD")
         For i = 0 To ieTag.Length - 1
             If Trim(WorksheetFunction.Clean(ieTag(i).innerText)) = "Service" Then
                 sSerLvl = Trim(ieTag(i + 1).innerText)
                 Exit For
             End If
         Next i
        
         'Pickup Date
        sOrgDate = .document.getElementsByClassName("snapshotController_date orig")(0).innerText
         sOrgDate = Right(sOrgDate, Len(sOrgDate) - 4)
         vOrgDate = CDate(sOrgDate)
        
         'Manifest Date
        bTagFound = False
         Set ieTag = .document.getElementsByTagName("TR")
         For i = 0 To ieTag.Length - 1
             If ieTag(i).innerHTML Like "*travel_history_header_date*" Then
                 sManDate = WorksheetFunction.Clean(Trim(ieTag(i).innerText))
                 vManDate = CDate(Left(sManDate, 10))
             End If
             If ieTag(i).innerText Like "*Shipment information sent to FedEx*" Then
                 bTagFound = True
                 Exit For
             End If
         Next i
         If bTagFound = False Then
             vManDate = ""
         End If

        
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

Function IsInArray(stringToBeFound As Variant, arr As String) As Boolean
  IsInArray = UBound(Filter(arr, stringToBeFound)) > -1
End Function



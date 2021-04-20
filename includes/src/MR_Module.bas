Attribute VB_Name = "MR_Module"
' Monthly Report Module

Sub FilterMonthlyTable()
    Dim filterStart As Long, filterEnd As Long
    filterStart = Worksheets("Monthly Report").Range("Monthly_StartDate").Value
    filterEnd = Worksheets("Monthly Report").Range("Monthly_EndDate").Value
    
    Worksheets("MR_Filter").ListObjects("Rfi__2").AutoFilter.ShowAllData
    
    Worksheets("MR_Filter").ListObjects("Rfi__2").Range.AutoFilter Field:=3, _
        Criteria1:="<=" & filterEnd
        
    Worksheets("MR_Filter").ListObjects("Rfi__2").Range.AutoFilter Field:=4, _
        Criteria1:=">=" & filterEnd, _
        Operator:=xlOr, _
        Criteria2:="=" & ""
        
    Worksheets("MR_Filter").ListObjects("Rfi__2").Range.AutoFilter Field:=5, _
        Criteria1:=">=" & filterEnd, _
        Operator:=xlOr, _
        Criteria2:="=" & ""
End Sub


Function RFI_Response_Time(dateSent, dateResponded, dateAnswered, Optional ByVal filterDateStart, Optional ByVal filterDateEnd)
    Application.Volatile
    'Dim dateSent: dateSent = Range("C20").Value
    'Dim dateResponded: dateResponded = Range("E20").Value
    'Dim dateAnswered: dateAnswered = Range("D20").Value
    'Dim filterDateStart: filterDateStart = Range("Monthly_StartDate").Value
    'Dim filterDateEnd: filterDateEnd = Range("Monthly_EndDate").Value
    
    'RFI_Response_Time([@Sent],[@[Responded On]],[@[Answer Marked On]],Monthly_StartDate,Monthly_EndDate)
    
    If IsMissing(filterDateEnd) Then filterDateEnd = Date
    
    If dateSent = "" Then
        RFI_Response_Time = ""
        Exit Function
    End If
    
    If dateResponded = "" And dateAnswered <> "" Then
        dateResponded = dateAnswered
    End If
    
    
    If dateResponded < dateSent Then
        If dateResponded = "" And dateAnswered = "" Then GoTo exitifCheck
        RFI_Response_Time = 1000000
        Exit Function
    End If
exitifCheck:
    
    If dateResponded - dateSent < 0 Then
        dateResponded = filterDateEnd
    End If
    
    RFI_Response_Time = dateResponded - dateSent

End Function

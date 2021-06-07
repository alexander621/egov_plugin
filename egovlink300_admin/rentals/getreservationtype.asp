<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getreservationtype.asp
' AUTHOR: Steve Loar
' CREATED: 10/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the reservation type selector. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/26/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, sResults, iReservationTypeId

iReservationTypeId = CLng(request("reservationtypeid"))

sSql = "SELECT reservationtypeselector FROM egov_rentalreservationtypes WHERE reservationtypeid = " & iReservationTypeId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRs.EOF Then
	sResults = oRs("reservationtypeselector")
Else
	sResults = "unknown"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults 

%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkselectedclassdates.asp
' AUTHOR: Steve Loar
' CREATED: 10/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed dates and times are OK, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/20/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iRentalId, iMaxRows, x, iReservationTempId, iEndDay, sStartDateTime, sEndDateTime
Dim iPosition, iEndHour, iEndMinute, sEndAmPm, sReturn, sOffSeasonFlag, iReservationTypeId, iRentalUserid

iRentalId = CLng(request("rentalid"))

iMaxRows = CLng(request("maxrows"))

iReservationTempId = CLng(request("rti"))

iPosition = 0
sReturn = "OK"

' Clear out the old date rows
sSql = "DELETE FROM egov_rentalreservationdatestemp WHERE reservationtempid = " & iReservationTempId
RunSQLStatement sSql

For x = 1 To iMaxRows
	If request("startdate" & x) <> "" Then 
		'response.write "includereservationtime" & x & " = " & request("includereservationtime" & x) & "\n"
		If LCase(CStr(request("includereservationtime" & x))) = "true" Then 
			iPosition = iPosition + 1

			sStartDateTime = request("startdate" & x) & " " & request("starthour" & x) & ":" & request("startminute" & x) & " " & request("startampm" & x)
			iEndDay = request("endday" & x)
			sEndDateTime = request("startdate" & x) & " " & request("endhour" & x) & ":" & request("endminute" & x) & " " & request("endampm" & x)

			' Save the row to the temp table
			SaveClassWantedDate iReservationTempId, sStartDateTime, sEndDateTime, iEndDay, iPosition, request("timedayid" & x)

			sCheckReturn = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, True )	' In rentalscommonfunctions.asp

			If sCheckReturn <> "OK" Then 
				sReturn = sCheckReturn
			End If 
		End If 
	End If 
Next

response.write sReturn

%>
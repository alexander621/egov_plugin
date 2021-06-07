<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: waiver_remove.asp
' AUTHOR: John Stullenberger
' CREATED: 2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Remove the waiver
'
' MODIFICATION HISTORY
' 1.0   2006	John Stullenberger - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

Dim iWaiverId, iFacilityId, sSql

iWaiverId = CLng(request("iWaiverId"))
iFacilityId = CLng(request("iFacilityId"))

	
sSql = "DELETE FROM egov_facilitywaivers WHERE facilityid = " & iFacilityId & " AND waiverid = " &  iWaiverId 

RunSQLStatement sSql

' REDIRECT TO facility waivers page
response.redirect( "facility_waivers.asp?facilityid=" & iFacilityId )


%>
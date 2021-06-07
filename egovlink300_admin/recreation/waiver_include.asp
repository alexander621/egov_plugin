<!-- #include file="../includes/common.asp" //-->
<%
Dim iWaiverId, iFacilityId, sSql

iWaiverId = CLng(request("iWaiverId"))
iFacilityId = CLng(request("iFacilityId"))

	
sSql = "INSERT INTO egov_facilitywaivers ( facilityid, waiverid ) VALUES (" & iFacilityId & ", " &  iWaiverId & " )"

RunSQLStatement sSql

' REDIRECT TO facility waivers page
response.redirect( "facility_waivers.asp?facilityid=" & iFacilityId )

%>
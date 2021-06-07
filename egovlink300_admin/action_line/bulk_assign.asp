<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
 sLevel     = "../"     'Override of value from common.asp

'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
end if

arrRequests = split(request.querystring("irequestid"),",")
%>
<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<meta content="width=device-width, minimum-scale=1, maximum-scale=1" name="viewport" />
	

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

</head>
<body>



	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">

	<div class="boxcontent">
			<!--BEGIN: PAGE TITLE-->
				<font size="+1"><strong>Bulk Processing</strong></font>
			<!--END: PAGE TITLE-->
			<br />
			<br />
<input type="button" name="sBack" id="sBack" value="Back" class="button" onclick="location.href='action_line_list.asp'" />

<%
NoRedirect = false
For Each requestform in arrRequests
	'response.write "<hr>" & requestform & "<hr><br />"

	sSQL = "SELECT * FROM egov_actionline_requests WHERE action_autoid = '" & requestform & "'"
	set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1

	if not oRs.EOF then
		paramList = "TrackID=" & requestform
		paramList = paramList & "&currentStatus=" & oRs("status")
		paramList = paramList & "&currentSubStatus=" & oRs("sub_status_id")
		paramList = paramList & "&prevAssignedemployeeid=" & oRs("assignedemployeeid")
		paramList = paramList & "&currentDepartmentID=" & oRs("groupid")
		paramList = paramList & "&currentDueDate=" & oRs("due_date")
		paramList = paramList & "&selSubStatus=" & oRs("sub_status_id")
		paramList = paramList & "&internal_comment="
		paramList = paramList & "&external_comment="
		paramList = paramList & "&due_date=" & oRs("due_date")
		paramList = paramList & "&autouid=" & request.cookies("user")("userid")
		paramList = paramList & "&autooid=" & request.cookies("user")("orgid")
		paramList = paramList & "&autofullname=" & request.cookies("user")("fullname")
		paramList = paramList & "&autolocid=" & request.cookies("user")("locationid")
		paramList = paramList & "&autosst=" & request.cookies("user")("showstockticker")
		paramList = paramList & "&autop=" & request.cookies("user")("permissions")
	
	
	
		paramList = paramList & "&assignedemployeeid=" & request.querystring("bulkemployeeid")
		paramList = paramList & "&deptid=" & request.querystring("bulkdeptid")
		paramList = paramList & "&selStatus=" & request.querystring("bulkstatus")
		'response.write paramList

      		TrackingNumber = requestform  & replace(FormatDateTime(cdate(oRs("submit_date")),4),":","")
	
	
		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
	
		' Set timeouts of resolve(0), connection(60000), send(30000), receive(30000) in milliseconds. 0 = infinite
		objWinHttp.SetTimeouts 0, 120000, 60000, 120000
	
		'URL = replace(request.servervariables("URL"),"bulk_assign","action_respond")
		URL = replace(session("egovclientwebsiteurl"),"http:","https:") & "/admin/action_line/action_respond.asp"
		'response.write URL & "<hr>"

		objWinHttp.Open "POST", URL, False
		objWinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	
		
	
		'Send Parameter List
		objWinHttp.Send paramList
	
		' Get the text of the response.
		transResponse = objWinHttp.ResponseText
	
		if instr(transResponse,"Tracking Number") > 0 then
			response.write "<p>" & TrackingNumber & " - SUCCESSFULLY UPDATED!</p>"
		else
			response.write "<p><a href=""action_respond.asp?control=" & TrackingNumber & """>" & TrackingNumber & "</a> - Sorry, we couldn't update this request.</p>"
			NoRedirect = true
			'response.write transResponse & "<hr>"
		end if

		Set objWinHttp = Nothing
	
	end if
	oRs.Close
	Set oRs = Nothing
next

if not NoRedirect then response.redirect "action_line_list.asp"



%>
	</div>

</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>

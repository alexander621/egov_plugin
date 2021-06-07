<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: emailinspectionreport.asp
' AUTHOR: Terry Foster
' CREATED: 12/16/2019
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	emails permit inspection reports
'
' MODIFICATION HISTORY
' 1.0   12/16/2019	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
intPermitInspectionReportID = CLng(request("permitinspectionreportid"))
intPermitID = 0
if request.servervariables("REQUEST_METHOD") = "POST" then
	EmailReport
end if

strContacts = ""

'GET FORM DATA FROM DATABASE BY INSPECTION REPORT ID
sSQL = "SELECT ir.*,p.permitstatusid,i.inspectionstatusid " _
	& " FROM egov_permitinspectionreports ir " _
	& " INNER JOIN egov_permits p ON p.permitid = ir.permitid " _
	& " INNER JOIN egov_permitinspections i on i.permitinspectionid = ir.permitinspectionid " _
	& " WHERE permitinspectionreportid = '" & intPermitInspectionReportID & "'"
set oRsIR = Server.CreateObject("ADODB.RecordSet")
oRsIR.Open sSQL, Application("DSN"), 3, 1
if not oRsIR.EOF then
	intPermitid = oRsIR("permitID")
	intPermitInspectionID = oRsIR("permitinspectionid")
	intPermitStatusID = oRsIR("permitstatusid")
	intInspectionStatusID = oRsIR("inspectionstatusid")
	sSQL = "SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, ISNULL(company,'') AS company ,email FROM egov_permitcontacts WHERE permitid = " & intPermitID & " and email is not null"
	set oRsC = Server.CreateObject("ADODB.RecordSet")
	oRsC.Open sSQL, Application("DSN"), 3, 1
	Do While Not oRsC.EOF
		if oRsC("firstname") <> "" then
			strContacts = strContacts & replace(oRsC("firstname"),",","|")
			if strContacts <> "" then strContacts = strContacts & " "
			strContacts = strContacts & replace(oRsC("lastname"),",","|")
			if strContacts <> "" then strContacts = strContacts & " "
		end if
		if oRsC("company") <> "" then
			strContacts = strContacts & replace(oRsC("company"),",","|")
			if strContacts <> "" then strContacts = strContacts & " "
		end if
		strContacts = trim(strContacts)

		strContacts = strContacts & "##" & oRsC("email")
		
		if strContacts <> "" then strContacts = strContacts & ","

		oRsC.MoveNext
	loop
	oRsC.Close
	Set oRsC = Nothing
end if
oRsIR.Close
Set oRsIR = Nothing
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/layers.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
  <style>
  	table th {font-weight:bold !important;}
	.tblinspectiontypes td {text-align:center;}
	.tblinspectiontypes tr.approvalchecks td 
	{
		text-align:left;
		font-weight:bold !important;
	}
  </style>

	<script language="Javascript">
	<!--
		
		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}






	//-->
	</script>


</head>

<body>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	

	<!--BEGIN: EDIT FORM-->
	<form name="frmEmailInspectionReport" method="post">
	<input type="hidden" name="permitid" value="<%=intPermitID%>" />
	<input type="hidden" name="permitinspectionid" value="<%=intPermitInspectionID%>" />
	<input type="hidden" name="permitinspectionreportid" value="<%=intPermitInspectionReportID%>" />
	<input type="hidden" name="permitstatusid" value="<%=intPermitStatusID%>" />
	<input type="hidden" name="inspectionstatusid" value="<%=intInspectionStatusID%>" />


	<% 
		arrContacts = split(strContacts, ",")
		For x = 0 to UBOUND(arrContacts) - 1
			arrContact = split(arrContacts(x),"##")
			'response.write arrContacts(x) & "<br />"
			strContact = arrContacts(x)
			strName = replace(arrContact(0),"|",",")
			strEmail = arrContact(1)
			'response.write strName & " has an email of " & strEmail & "<br />"
			%>
			<input type="checkbox" name="toaddresses" value="<%=strContact%>" /><%=strName%> - <%=strEmail%><br />
			<%
		next

	%>
	<br />

	(enter email addresses separated by a semicolon ";")<br />
	Additional TO: <input type="text" name="additionalto" style="width:100%" value="" />
	<br />
	<br />
	CC: <input type="text" name="cc" style="width:100%" value="" />

	</p>
	<input type="submit" class="button ui-button ui-widget ui-corner-all" value="SEND" />

	</form>
	<!--END: EDIT FORM-->

	</div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

sub EmailReport
	strFrom = getUserEmail(session("userid"))
	if strFrom = "" then strFrom = "noreplies@egovlink.com"
	strSubject = "Inspection Report for " & GetPermitNumber(request.form("PermitID")) & " " & GetPermitLocation( request.form("PermitID"), sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, False )
	strBody = GetReport(request.form("permitinspectionreportid"))
	'response.write strBody
	'response.end

	'strTo = replace(request.form("toaddresses"),",",";") & ";" & request.form("additionalto")
	strCC = request.form("cc")

	strTo = ""
	arrContacts = split(request.form("toaddresses"),",")
	for x = 0 to UBOUND(arrContacts)
		arrContact = split(arrContacts(x),"##")
		strTo = strTo & arrContact(1) & ";"
	next
	strTo = strTo & request.form("additionalto")

	sendEmail strFrom, strTo, strCC, strSubject, strBody, "", "N"

	'Enter a note
	MakeAPermitLogEntry request.form("permitid"), "'Permit Inspection Report Emailed'", "'Inspection Report emailed TO: " & strTo & " and CC: " & strCC & "'", "NULL", "NULL", request.form("PermitStatusId"), 1, 0, 0, "NULL", request.form("PermitInspectionId"), "NULL", request.form("InspectionStatusId")   ' in permitcommonfunctions.asp

	
	response.write "<script>parent.hideModal(window.frameElement.getAttribute(""data-close""));</script>"
	
end sub

Function GetReport(intIRID)
	Set obj = CreateObject("MSXML2.ServerXMLHTTP")
	obj.Open "GET", "https://www.egovlink.com/permitcity/admin/permits/inspectionreport.asp?print=true&permitid=" & request.form("PermitID") & "&permitinspectionreportid=" & intIRID, False
	obj.Send ""
	GetReport = obj.ResponseText
	Set obj = Nothing 
End Function


%>

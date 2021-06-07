<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectorroute.asp
' AUTHOR: Steve Loar
' CREATED: 08/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays the inspection route details for printing.
'
' MODIFICATION HISTORY
' 1.0   08/12/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iRowCount

iRowCount = 0

sSql = session("sSql")  ' This is filled in by the calling page
sSql = sSql & " ORDER BY routeorder, I.scheduleddate"
'response.write sSql & "<br /><br />"

%>

<html>
<head>
	<title>E-Gov Permit Inspection Route</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

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
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		
	<!--BEGIN: PAGE TITLE-->
	<!--END: PAGE TITLE-->

<%

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	response.write "<table id=""inspectorroute"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
	response.write "<tr><th>Route<br />Order</th><th>Permit #</th><th>Permit Type</th><th>Address/Location</th><th>Scheduled</th><th>Inspection</th><th>Inspection<br />Status</th><th>Reinspection</th><th>Final</th><th>Contact</th><th>Notes</th><th>Inspector</th></tr>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		iRowCount = iRowCount + 1
		' Route Order
		response.write "<td align=""center"">" & iRowCount & "</td>"
		
		' Permit No
		response.write "<td nowrap=""nowrap"">"
		response.write GetPermitNumber( oRs("permitid") )
		response.write "</td>"

		' Permit Type
		response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("permittype") & "</td>"
		
		' Address/Location
		response.write "<td nowrap=""nowrap"">"
		Select Case oRs("locationtype")

			Case "address"
				response.write oRs("residentstreetnumber")
				If oRs("residentstreetprefix") <> "" Then
					response.write " " & oRs("residentstreetprefix")
				End If 
				response.write " " & oRs("residentstreetname")
				If oRs("streetsuffix") <> "" Then
					response.write " " & oRs("streetsuffix")
				End If 
				If oRs("streetdirection") <> "" Then
					response.write " " & oRs("streetdirection")
				End If 
				response.write " " & oRs("residentunit")
				response.write "<br />" & oRs("residentcity")

			Case "location"
				response.write Replace(oRs("permitlocation"),Chr(10),"<br />")

			Case Else 
				response.write "&nbsp;"

		End Select
		response.write "</td>"
		
		' Scheduled date and time 
		response.write "<td align=""center"">" & FormatDateTime( oRs("scheduleddate"),2 )
		If oRs("scheduledtime") <> "" Then
			response.write "<br />" & oRs("scheduledtime") & " " & oRs("scheduledampm")
		End If 
		response.write "</td>"
		
		' Inspection Type
		response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("permitinspectiontype") & "</td>"
		
		' Status
		response.write "<td align=""center"">" & oRs("inspectionstatus") & "</td>"
		
		' Required
		response.write "<td align=""center"">" 
		If oRs("isreinspection") Then
			response.write "Yes" 
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		' Final Insp
		response.write "<td align=""center"">" 
		If oRs("isfinal") Then
			response.write "Yes" 
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		' Contact
		response.write "<td nowrap=""nowrap"" align=""center"">"
		If oRs("contact") = "" And oRs("contactphone") = "" Then
			response.write "&nbsp;"
		Else 
			response.write oRs("contact") & "<br />" & oRs("contactphone")
		End If 
		response.write "</td>"

		' Notes
		response.write "<td>"
		If oRs("schedulingnotes") <> "" Then 
			response.write oRs("schedulingnotes") 
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		' Inspector
		response.write "<td align=""center"">"
		response.write oRs("FirstName") & " " & oRs("LastName")
		response.write "</td>"

		response.write "</tr>"

		oRs.MoveNext 
	Loop 
	response.write vbcrlf & "</table>"
Else
	response.write vbcrlf & "<p>No Permits could be found that match your search criteria.</p>"
End If 

oRs.Close
Set oRs = Nothing 

%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



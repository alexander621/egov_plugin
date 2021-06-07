<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: LOCATION_MGMT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.1   04/17/06	TERRY FOSTER - MADE FUNCTIONAL
' 1.2	10/11/06	Steve Loar - Security, Header and nav changed
' 1.3	06/07/2010	Steve Loar - Changed display order to be in name order
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "locations" ) Then
	If Not UserHasPermission( Session("UserId"), "rentallocations" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If 
End If 

%>

<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
 	<link rel="stylesheet" type="text/css" href="classes.css" />
	
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		function deleteconfirm(ID, sName) 
		{
			var sStringName = new String( sName )
			if(confirm('Do you wish to delete ' + sStringName + '?')) 
			{
				window.location="location_delete.asp?iLocationid=" + ID;
			}
		}

	//-->
	</script>

</head>

<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: Location Management</strong></font><br />
	<!--<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS LIST-->
<%

	ListLocations 

%> 
<!--END: CLASS LIST-->

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>


</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ListLocations()
'--------------------------------------------------------------------------------------------------
Sub ListLocations()
	Dim sSql, oRs, iRowCount

	iRowCount = 0
	' GET ALL LOCATIONS FOR ORG
	sSql = "SELECT locationid, name, address1, address2, city, state, zip "
	sSql = sSql & "FROM egov_class_location WHERE orgid = " & SESSION("ORGID")
	sSql = sSql & " ORDER BY name" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	' DRAW LINK TO NEW location
'	response.write "<div id=""functionlinks""><a href=""location_edit.asp?locationid=0""><img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;New Location</a></div>"
	response.write "<div id=""functionlinks""><input type=""button"" class=""button"" onclick=""location.href='location_edit.asp?locationid=0';"" value=""New Location"" /></div>"

	If NOT oRs.EOF Then

		' DRAW TABLE 
		response.write vbcrlf & "<div class=""shadow"">" & vbcrlf & "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""locationlist style-alternate"" width=""60%"">"
		
		' HEADER ROW
		response.write vbcrlf & "<tr>"
		response.write "<th>Location Name</th><th>Address 1</th><th>Address 2</th><th>City</th><th>State</th><th>Zip</th>"
		response.write "</tr>"
		
		' LOOP THRU AND DISPLAY ROWS
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			response.write "<td title=""click to edit"" onClick=""location.href='location_edit.asp?locationid=" & oRs("locationid") & "';"">"
			response.write "&nbsp;" & Trim(oRs("name")) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='location_edit.asp?locationid=" & oRs("locationid") & "';"">" & Trim(oRs("address1")) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='location_edit.asp?locationid=" & oRs("locationid") & "';"">" & Trim(oRs("address2")) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='location_edit.asp?locationid=" & oRs("locationid") & "';"">" & Trim(oRs("city")) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='location_edit.asp?locationid=" & oRs("locationid") & "';"">" & Trim(oRs("state")) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='location_edit.asp?locationid=" & oRs("locationid") & "';"">" & Trim(oRs("zip")) & "</td>"

'			response.write "<td><a title=""click to delete"" href=""javascript:deleteconfirm(" & oRs("locationid") & ", '" & FormatForJavaScript(oRs("name")) & "')"">Delete</a></td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write vbcrlf & "</table></div>"
	
	Else
		' NO LOCATIONS WERE FOUND
		response.write "<font color=""red""><b>There are no locations to display.</b></font>"
	
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


%>



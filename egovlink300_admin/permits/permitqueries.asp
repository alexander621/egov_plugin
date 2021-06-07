<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
  Dim sError
	ReDim aStatuses(0)

 'Set Timezone information into session
  session("iUserOffset") = request.cookies("tz")

 'Override of value from common.asp
  sLevel = "../"

  intQueryID = request("queryid")

  if request("runquery") = "Run Query" then RunQuery()


%>
<html>
<head>
  <title><%=langBSHome%></title>

		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

<script language="javascript">
<!--

		function searchcontacts( sFieldId )
		{
			var w = (screen.width - 640)/2;
			var h = (screen.height - 480)/2;
			var winHandle = eval('window.open("contractorpicker.asp?fieldid=' + sFieldId + '", "_contact", "width=600,height=400,location=1,toolbar=1,statusbar=0,scrollbars=1,menubar=1,left=' + w + ',top=' + h + '")');
		}
		function doCalendar( sField ) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  //eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=permitquerycriteria", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
//-->
</script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content" style="width:auto;">
 	<div id="centercontent">
		<div class="gutters">

<table border="0" cellpadding="0" cellspacing="0" class="start">
  <tr valign="top">
    	<td>
	<h2>Permit Queries</h2>
	<fieldset>
		<legend>Select Your Query</legend>
	<form action="#" method="POST" name="querysel">
		<select name="queryid" onchange="document.querysel.submit()" style="min-width:600px; font-size:14px;">
			<option value="">Choose...</option>
			<%
				sSQL = "SELECT egov_permitqueryid,queryname FROM egov_permitquery WHERE orgid = 0 or orgid = " & session("orgid")
				Set oRs = Server.CreateObject("ADODB.RecordSet")
				oRs.Open sSQL, Application("DSN"), 3, 1
				Do While Not oRs.EOF
					selected = ""
					if oRs("egov_permitqueryid") & "" = intQueryID & "" then selected = " selected"
					%>
					<option value="<%=oRs("egov_permitqueryid")%>" <%=selected%>><%=oRs("queryname")%></option>
					<%
					oRs.MoveNext
				loop
				oRs.Close
				Set oRs = Nothing
			%>
		</select>
	</form>
	</fieldset>
	<br />
	<br />

	<% if intQueryID <> "" then %>
	<fieldset>
		<legend>Criteria</legend>
		<form name="permitquerycriteria" action="#" method="POST">
			<input type="hidden" name="queryid" value="<%=intQueryID%>" />
			<table cellpadding="2" cellspacing="0" border="0">
				<tr>
					<td>Category:</td><td><%  ShowPermitCategoryPicks iPermitCategoryId %></td>
				</tr>
				<tr>
					<td>Address:</td><td><%  DisplayLargeAddressList sStreetNumber, sStreetName %></td>
				</tr>
				<tr>
					<td>Location Like:</td><td><input type="text" name="LIKE:permitlocation" size="100" maxlength="100" value="<%=sPermitLocation%>" /></td>
				</tr>
				<tr>
					<td>Legal Desc:</td><td><input type="text" name="LIKE:legaldescription" size="100" maxlength="100" value="<%=sLegalDescription%>" /></td>
				</tr>
				<tr>
					<td>Parcel Id #:</td><td><input type="text" name="parcelidnumber" size="25" maxlength="25" value="<%=sParcelIdNumber%>" /></td>
				</tr>
				<tr>
					<td>Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
				</tr>
				<tr>
					<td>Listed Owner:</td><td><input type="text" name="LIKE:listedowner" size="100" maxlength="100" value="<%=sListedOwner%>" /></td>
				</tr>
				<!--tr>
					<td>Contact:</td><td nowrap="nowrap">
					<input type="text" id="contactname" name="contactname" size="65" maxlength="100" value="<%=sContactname%>" /> &nbsp; &nbsp;
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Contractor Search" onclick="searchcontacts('contactname');" />
					</td>
				</tr>
				<tr>
					<td>Contact Types:</td>
					<td>
						<input type="checkbox" name="contacttype" value="isapplicant"
<%						For Each item In request("contacttype")
							If item = "isapplicant" Then
								response.write " checked=""checked"" "
							End If 
						Next	%>
						/>Applicant &nbsp; 
						<input type="checkbox" name="contacttype" value="isprimarycontact" 
<%						For Each item In request("contacttype")
							If item = "isprimarycontact" Then
								response.write " checked=""checked"" "
							End If 
						Next	%>
						/>Primary Contact &nbsp; 
						<input type="checkbox" name="contacttype" value="isbillingcontact" 
<%						For Each item In request("contacttype")
							If item = "isbillingcontact" Then
								response.write " checked=""checked"" "
							End If 
						Next	%>
						/>Billing Contact &nbsp; 
						<input type="checkbox" name="contacttype" value="isprimarycontractor" 
<%						For Each item In request("contacttype")
							If item = "isprimarycontractor" Then
								response.write " checked=""checked"" "
							End If 
						Next	%>
						/>Primary Contractor &nbsp; 
						<input type="checkbox" name="contacttype" value="isarchitect" 
<%						For Each item In request("contacttype")
							If item = "isarchitect" Then
								response.write " checked=""checked"" "
							End If 
						Next	%>
						/>Architect/Engineer <br />
						<input type="checkbox" name="contacttype" value="iscontractor" 
<%						For Each item In request("contacttype")
							If item = "iscontractor" Then
								response.write " checked=""checked"" "
							End If 
						Next	%>
						/>Other Contractors
					</td>
				</tr-->
				<% bInitialLoad = True %>
				<tr>
					<td>Permit Status:</td><td><% ShowPermitStatuses aStatuses, bInitialLoad %></td>
				</tr>
				<tr>
					<td>Permit Type:</td><td><% ShowPermitTypes iPermitTypeId %></td>
				</tr>
				<tr>
					<td>Date Range:</td>
					<td nowrap="nowrap">
						<select name="permitdate">
							<option value="none"
<%							If request("permitdate") = "none" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Select a Date...</option>

							<option value="applieddate"
<%							If request("permitdate") = "applieddate" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Applied</option>
							<option value="releaseddate"
<%							If request("permitdate") = "releaseddate" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Released</option>
							<option value="approveddate"
<%							If request("permitdate") = "approveddate" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Approved</option>
							<option value="issueddate"
<%							If request("permitdate") = "issueddate" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Issued</option>
							<option value="completeddate"
<%							If request("permitdate") = "completeddate" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Completed</option>
							<option value="expirationdate"
<%							If request("permitdate") = "expirationdate" Then 
								response.write " selected=""selected"" "
							End If	%>
							>Expired</option>
						</select>
						&nbsp; From:
						<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" class="datepicker" />
						&nbsp; To:
						<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" class="datepicker" />
						&nbsp;
						<%DrawDateChoices "Date" %>
					</td>
				</tr>
				<tr>
					<td>Last Activity:</td>
					<td nowrap="nowrap">
						From:
						<input type="text" id="fromactivitydate" name="FROM:lastactivitydate" value="<%=sFromActivityDate%>" size="10" maxlength="10" class="datepicker" />
						&nbsp; To:
						<input type="text" id="toactivitydate" name="THRU:lastactivitydate" value="<%=sToActivityDate%>" size="10" maxlength="10" class="datepicker" />
						&nbsp;
						<%DrawDateChoices "activitydate" %>
					</td>
				</tr>
			</table>
			<br />
			<input type="submit" name="runquery" class="button ui-button ui-widget ui-corner-all" value="Run Query" style="font-size:16px;" />
		</form>
	</fieldset>
	<% end if %>

      </td>
  </tr>
</table>

  </div>
  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
Sub RunQuery()
	sSQL = "SELECT query FROM egov_permitquery WHERE egov_permitqueryid = " & intQueryID
	Set oQs = Server.CreateObject("ADODB.RecordSet")
	oQs.Open sSQL, Application("DSN"), 3, 1
	if not oQs.EOF then
		
		strWhere = ""
		for each item in request.form
			if item <> "queryid" and item <> "runquery" and request.form(item) <> "" and request.form(item) <> "0" _
				and request.form(item) <> "0000" and item <> "fromDate" and item <> "toDate" and item <> "Date" and item <> "activitydate" then

				'Date From
				if instr(item,"FROM:") > 0 then
					strWhere = strWhere & " AND " & replace(item,"FROM:","") & " >= '" & request.form(item) & " 12:00:00 AM' "

				'Date To
				elseif instr(item,"THRU:") > 0 then
					strWhere = strWhere & " AND " & replace(item,"THRU:","") & " <= '" & DateAdd("s",-1,DateAdd("d",1,request.form(item) & " 12:00:00 AM")) & "' "

				'LIKE
				elseif instr(item,"LIKE:") > 0 then
					strWhere = strWhere & " AND " & replace(item,"LIKE:","") & " LIKE '%" & dbsafe(request.form(item)) & "%' "

				'checkbox
				elseif instr(request.form(item),",") > 0 then
					strWhere = strWhere & " AND ("
					blnFirstCheck = true
					for each checkboxval in request.form(item)
						if not blnFirstCheck then strWhere = strWhere & " OR "

						strWhere = strWhere & item & " = '" & dbsafe(checkboxval) & "' "

						blnFirstCheck = false
					next
					strWhere = strWhere & ") "
					
				elseif item = "permitdate" then
					if request.form("fromDate") <> "" then
						strWhere = strWhere & " AND " & request.form(item) & " >= '" & request.form("fromDate") & " 12:00:00 AM'"
					end if

					if request.form("toDate") <> "" then
						strWhere = strWhere & " AND " & request.form(item) & " <= '" & DateAdd("s",-1,DateAdd("d",1,request.form("toDate") & " 12:00:00 AM")) & "' "
					end if

				elseif item = "permitno" then
					strWhere = strWhere & BuildPermitNoSearch( trim(request.form(item)) )

				elseif item = "streetname" then
					sStreetName = request("streetname")
					strWhere = strWhere & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
					strWhere = strWhere & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
					strWhere = strWhere & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
					strWhere = strWhere & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"

				'Everything else
				elseif request.form(item) <> "" then
					fieldName = item
					if item = "permitcategoryid" or item = "permittypeid" then fieldName = "p." & fieldName

					strWhere = strWhere & " AND " & fieldName & " = '" & dbsafe(request.form(item)) & "' "
				end if
			end if
		next

		sSQL = replace(replace(oQs("query"),"||WHERE||",strWhere),"||ORGID||",session("orgid"))

		'response.write sSQL
		Set oRs = Server.CreateObject("ADODB.RecordSet")
		oRs.Open sSQL, Application("DSN"), 3, 1
		if not oRs.EOF then 
			'response.write "<br />" & oRs.RecordCount
			session("displayquery") = sSQL
			response.redirect "../export/csv_export.asp"
		end if
		oRs.Close
		Set oRs = Nothing

		response.end

	end if
	oQs.Close
	Set oQs = Nothing
	
End Sub
'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sStreetNumber, sStreetName 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName, sOldCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"
		sOldCompareName = "qwerty"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = UCase(oRs("residentstreetprefix")) & " " 
			End If 

			sCompareName = sCompareName & UCase(oRs("residentstreetname"))

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & UCase(oRs("streetsuffix"))
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & UCase(oRs("streetdirection"))
			End If 

			If sOldCompareName <> sCompareName Then 
				' only write out unique values
				sOldCompareName = sCompareName
				response.write vbcrlf & "<option value=""" & sCompareName & """"

				If sStreetName = sCompareName Then 
					response.write " selected=""selected"" "
				End If 

				response.write " >"
				response.write sCompareName & "</option>" 
			End If 

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 
'--------------------------------------------------------------------------------------------------
' void ShowPermitStatuses aStatuses, bInitialLoad 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitStatuses( ByRef aStatuses, ByVal bInitialLoad )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid, permitstatus, iscompletedstatus FROM egov_permitstatuses "
	sSql = sSql & " WHERE isissuedback = 0 AND orgid = " & session("orgid") & " ORDER BY permitstatusorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While NOT oRs.EOF
			response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid"" value=""" & oRs("permitstatusid") & """"
			If bInitialLoad Then
				' commented out for Loveland 11/19/2009 - Steve Loar
				'If Not oRs("iscompletedstatus") Then 
					response.write " checked=""checked"" "
				'End If 
			Else
				For Each Item In aStatuses
					If CLng(Item) = CLng(oRs("permitstatusid")) Then
						response.write " checked=""checked"" "
					End If 
				Next 
			End If 
			response.write " />" & oRs("permitstatus") & " &nbsp; "
			oRs.MoveNext
		Loop
		response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid""  value=""-1"""
		For Each Item In aStatuses
			If CLng(Item) = CLng(-1) Then
				response.write " checked=""checked"" "
			End If 
		Next 
		response.write " />Hold &nbsp; "
		response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid""  value=""-2"""
		For Each Item In aStatuses
			If CLng(Item) = CLng(-2) Then
				response.write " checked=""checked"" "
			End If 
		Next 
		response.write " />Void &nbsp; "
		response.write vbcrlf & "<input type=""checkbox"" name=""permitstatusid""  value=""-3"""
		For Each Item In aStatuses
			If CLng(Item) = CLng(-3) Then
				response.write " checked=""checked"" "
			End If 
		Next 
		response.write " />Expired &nbsp; "

	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: instructor_roster_list.asp
' AUTHOR: Steve Loar
' CREATED: 05/10/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/10/07	Steve Loar - INITIAL VERSION
' 1.1	09/28/2012	Steve Loar - Adding roster export feature as on the roster list
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStatusid, iCategoryid, iClasstypeid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId
Dim sSearchName, iInstructorID

'Check to see if the feature is offline
If isFeatureOffline("activities") = "Y" Then 
	response.redirect "../admin/outage_feature_offline.asp"
End If 

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "instructor rosters" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

Session("RedirectPage") = GetCurrentURL()
Session("RedirectLang") = "Return To Instructor Rosters"

' GET INSTRUCTOR ID
iInstructorID = GetUserInstructorId( Session("UserId") )

Dim sShowDatesText, sShowDates, sShowDates2, bFilter

If request("showdates") = "" Or clng(request("showdates")) = clng(1) Then
	sShowDates = 1
	sShowDates2 = 2
	sShowDatesText = "Show All"
Else
	sShowDates = 2
	sShowDates2 = 1
	sShowDatesText = "Show Upcoming Only"
End If 
bFilter = False 

If request("classseasonid") = "" or clng(request("classseasonid")) = 0 Then 
	iClassSeasonId = 0
Else
	iClassSeasonId = clng(request("classseasonid"))
End If 

If request("categoryid") = "" or clng(request("categoryid")) = 0 Then 
	iCategoryid = 0
Else
	iCategoryid = request("categoryid")
End If 

If request("datefilter") = "" Then 
	iDatefilter = ""
Else
	iDatefilter = request("datefilter")
	bFilter = True 
End If 
If request("startdate") = "" Then 
	sStartdate = ""
Else
	sStartdate = request("startdate")
	bFilter = True
End If
If request("enddate") = "" Then 
	sEnddate = ""
Else
	sEnddate = request("enddate")
	bFilter = True
End If
If iDatefilter = "alldates" Then
	sStartdate = ""
	sEnddate = ""
	bFilter = False 
End If

sSearchName = request("searchname")

' if all date choices are blank, give them the current published classes and events
If sShowDates = 1 And bFilter = False Then
	sDefaultRange = " and (('' + convert(char(8),getdate(),112) + '' >= publishstartdate and '' + convert(char(8),getdate(),112) + '' <= publishenddate) Or publishstartdate > '' + convert(char(8),getdate(),112) + '' Or publishstartdate is null) "
Else 
	sDefaultRange = ""
End If 

%>


<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.7.2.min.js"></script>

	<script language="Javascript">
	<!--

		function GoToPrint( iClassid, iTimeid )
		{
			//var sendtoprint = document.getElementById("print_000000");
			//sendtoprint.value = iClassid + ',' + iTimeid;
			$("#print_000000").val( iClassid + ',' + iTimeid );
			document.frmprint.submit();
		}

		function ToggleDisplay( showDates )
		{
			window.location.href = "instructor_roster_list.asp?showdates=" + showDates;
		}

		function toggleChecks()
		{
			var checkAttr;

			if ( $("#toggleselects").prop('checked')  )
				checkAttr = true;
			else
				checkAttr = false;

			var classid;
			//alert( $("#maxclasscount").val() );
			if (parseInt($("#maxclasscount").val()) > parseInt('0') )
			{
				for (x=1; x <= parseInt($("#maxclasscount").val()); x++)
				{
					//alert( $("#classid" + x).prop('checked') );
					//classid = $("#classno" + x).val();
					if (checkAttr)
						$("#classid" + x).prop('checked', true);
					else
						$("#classid" + x).prop('checked', false);
				}
			}
		}

		function RosterDownload() 
		{
			var bPicked = false;

			if (document.classForm.classid.length) 
			{   // Several picked
				var checklength = document.classForm.classid.length;
				for (i = 0; i < checklength; i++) 
				{
					if (document.classForm.classid[i].checked) 
					{
						bPicked = true;
						break;
					}
				}
			}
			else
			{ // Just one picked
				if (document.classForm.classid.checked)	
				{
					bPicked = true;
				}
			}

			if( bPicked ) 
			{
				document.getElementById("classForm").action = "export_roster.asp"
				document.classForm.submit();
			}
			else
			{
				alert("Please select at least one class to Download from the list.");
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
		<font size="+1"><strong>Recreation: Instructor Rosters</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->


	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
		<fieldset class="filterselection">
		<legend class="filterselection">Filter Options</legend>
		<p>
			<form name="ClassForm" method="post" action="instructor_roster_list.asp">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td>Season: </td>
				<td>
					<% ShowSeasonFilterPicks iClassSeasonId ' In class_global_functions.asp %>
				</td>
			</tr>
			<tr>
				<td>Category:</td>
				<td>
					<% DisplayCategorySelect iCategoryid  ' in class_global_functions.asp %>
				</td>
			</tr>
			<tr>
				<td>Name Like:</td>
				<td colspan="3"><input type="text" name="searchname" value="<%=sSearchName%>" size="75" maxlength="255" /></td>
			<tr>
				<td>Date:</td>
				<td>
					<select name="datefilter">
						<option value="alldates"<%If iDatefilter = "alldates" Then
														response.write "selected=""selected"" "
													End If%>>No Date Filter</option>
						<option value="startdate"<%If iDatefilter = "startdate" Then
														response.write "selected=""selected"" "
													End If%>>Start Date</option>
						<option value="publishstartdate"<%If iDatefilter = "publishstartdate" Then
														response.write "selected=""selected"" "
													End If%>>Publish Start Date</option>
						<option value="registrationstartdate"<%If iDatefilter = "registrationstartdate" Then
														response.write "selected=""selected"" "
													End If%>>Registration Start Date</option>
					</select>
				</td>
			<td>
				From: <input type="text" name="startdate" value="<%=sStartdate%>" /> <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span> To: <input type="text" name="enddate" value="<%=sEnddate%>"/> <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
			</td>

			</tr>
			<tr>
				<td align="right" colspan="3"><input class="button" style="margin:5px;" type="submit" value="Refresh Results"></td>
			</tr>
			</table>
			</form>
		</p>
		</fieldset>
	</div>
	<!--END: FILTER SELECTION-->

	<p>
		<!-- <a href="instructor_roster_list.asp?showdates=<%=sShowDates2%>" ><%=sShowDatesText%></a>&nbsp; -->
		<input type="button" value="<%=sShowDatesText%>" class="button" onclick="ToggleDisplay( <%=sShowDates2%> )" />&nbsp;
		<input type="button" value="Download Selected Rosters" class="button" onclick="RosterDownload()" /> 
	</p>

	<!--BEGIN: CLASS LIST-->
	<form id="classForm" name="classForm" method="post" action="roster_list_print.asp">

		<% DisplayClassEvents session("orgid"), iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId, sSearchName, iInstructorID %>

	</form>
	<!--END: CLASS LIST-->
	</div>
</div>

<form name="frmprint" action="print_roster.asp" method="post">
	<input type ="hidden" value="000,000" name="print_000000" id="print_000000" />
</form>

<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
' DisplayClassEvents iorgid, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId, sSearchName, iInstructorID 
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEvents( ByVal iOrgid, ByVal iCategoryid, ByVal iDatefilter, ByVal sStartdate, ByVal sEnddate, ByVal sDefaultRange, ByVal iClassSeasonId, ByVal sSearchName, ByVal iInstructorId )
	Dim sSql, sWhere, oRs, sFrom

	sWhere = ""
	sFrom = ""

	If clng(iClassSeasonId) <> clng(0) Then 
		sWhere = sWhere & " and C.classseasonid = " & iClassSeasonId
	End If 

	If clng(iCategoryid) <> 0 Then
		sWhere = sWhere & " and CC.classid = C.classid and CC.categoryid = " & iCategoryid
		sFrom = sFrom & ", egov_class_category_to_class CC "
	End If 

	If iDatefilter <> "alldates" And (sStartdate <> "" Or sEnddate <> "") Then
		If sStartdate <> "" Then 
			sWhere = sWhere & " and C." & iDatefilter & " >= '" & sStartdate & "' " 
		End If
		If sEnddate <> "" Then 
			sWhere = sWhere & " and C." & iDatefilter & " <= '" & sEnddate & "' "
		End If 
		sWhere = sWhere & " and C." & iDatefilter & " is not NULL "
	Else 
		' add in the default range of classes and events to get
		sWhere = sWhere & sDefaultRange
	End If 

	If sSearchName <> "" Then 
		sSearchName = dbsafe(sSearchName)
		sWhere = sWhere & " and Lower(C.classname) like Lower('%" & sSearchName & "%') "
	End If 

	sWhere = sWhere & " and instructorid = " & iInstructorID


	' GET CLASS\EVENTS FOR ORG that are active and can be purchased
	sSql = "Select C.classid, C.classname, C.isparent, T.classtypename, C.startdate, C.registrationstartdate, C.publishstartdate,"
	sSql = sSql & " S.statusname, CT.activityno, CT.timeid, CT.min, CT.max, CT.enrollmentsize, CT.waitlistmax, CT.waitlistsize " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S, egov_registration_option RO, egov_class_time CT " & sFrom
	sSql = sSql & " Where C.classtypeid = T.classtypeid and C.statusid = S.statusid and S.statusname = 'ACTIVE' " 
	sSql = sSql & " and RO.optionid = C.optionid and RO.canpurchase = 1 and C.classid = CT.classid and CT.instructorid = " & iInstructorId & sWhere
	sSql = sSql & " and C.orgid = " & iOrgid & " ORDER BY C.classname, CT.activityno"

'	response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1


	If Not oRs.EOF Then

		' DRAW TABLE WITH CLASSES LISTED
		'response.write vbcrlf & "<div class=""shadow"">"
		response.write vbcrlf & "<table id=""rosterlist"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""instructortable style-alternate"" >" 
		
		' HEADER ROW
		response.write vbcrlf & "<tr>"
		response.write "<th><input type=""checkbox"" id=""toggleselects"" name=""toggleselects"" onclick=""toggleChecks()"" /></th>"
		response.write "<th>Class Name</th>"
		response.write "<th>Type</th><th>Start<br />Date</th>"
		response.write "<th>Min</th><th>Max</th><th>Total<br />Enrld</th><th>%<br />Full</th><th>Total<br />Waiting</th>"
		response.write "</tr>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			' DISPLAY INFORMATION
			response.write vbcrlf & "<tr"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
			response.write ">"

			' selection check box
			response.write "<td>"
			'response.write "<input type=""checkbox"" name=""classid"" id=""classid" & oRs("classid") & """ value=""" & oRs("classid") & """ />"
			response.write "<input type=""checkbox"" name=""classid"" id=""classid" & iRowCount & """ value=""" & oRs("classid") & """ />"
			response.write "<input type=""hidden"" name=""classno" & iRowCount & """ id=""classno" & iRowCount & """  value=""" & oRs("classid") & """ />"
			response.write "</td>"

			response.write "<td class=""classname""><a href=""#"" onclick=""GoToPrint( " &  oRs("classid") & ", " &  oRs("timeid") & " );"">" & oRs("classname") & " &nbsp; ( " & oRs("activityno") & " )"
			If oRs("isparent") AND UCase(oRs("classtypename")) = "SERIES" Then
				response.write " -Series Purchase" 
			End If 
			response.write "</a></td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("classtypename") & "</td>"
			response.write "<td align=""center"">" & oRs("startdate") & "</td>"

			' Min
			response.write "<td align=""center"">" 
			If IsNull(oRs("min")) Then 
				response.write "n/a" 
			Else
				response.write oRs("min")
			End If 
			response.write "</td>"
			
			' Max
			response.write "<td align=""center"">" 
			If IsNull(oRs("max")) Then 
				response.write "n/a" 
				iMax = 1
			Else
				response.write oRs("max")
				iMax = clng(oRs("max"))
			End If 
			response.write "</td>"

			response.write "<td align=""center"">" & oRs("enrollmentsize") & "</td>"

			response.write "<td align=""center"">" & FormatNumber((clng(oRs("enrollmentsize")) / iMax) * 100,0) & "</td>"

			response.write "<td align=""center"">" & oRs("waitlistsize") & "</td>"

			response.write "</tr>"

			oRs.MoveNext
		Loop 

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<input type=""hidden"" name=""maxclasscount"" id=""maxclasscount"" value=""" & iRowCount & """ />"
		'response.write vbcrlf & "</div>"

		oRs.Close
		Set oRs = Nothing 
	
	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write "<font color=""red""><b>There are no classes\events that have you as the assigned instructor.</b></font>"
	End If

End Sub

%>

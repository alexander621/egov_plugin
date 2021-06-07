<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_list.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
' 1.2	11/1/2006	Steve Loar - Added link to toggle View All and View Upcoming Only
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStatusid, iCategoryid, iClasstypeid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "manage classes" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

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

If request("statusid") = "" or clng(request("statusid")) = 0 Then 
	iStatusid = 0
Else
	iStatusid = request("statusid")
End If 

If request("classtypeid") = "" or clng(request("classtypeid")) = 0 Then 
	iClasstypeid = 0
Else
	iClasstypeid = request("classtypeid")
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

' if all date choices are blank, give them the current published classes and events
If sShowDates = 1 And bFilter = False Then
	sDefaultRange = " and (('' + convert(char(8),getdate(),112) + '' >= publishstartdate and '' + convert(char(8),getdate(),112) + '' <= publishenddate) Or publishstartdate > '' + convert(char(8),getdate(),112) + '' Or publishstartdate is null) "
Else 
	sDefaultRange = ""
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />


	<script language="Javascript" src="tablesort.js"></script>

	<script language="Javascript">
	<!--

		function deleteconfirm(ID, sName) 
		{
			if(confirm('Do you wish to delete ' + sName + '?')) 
			{
				window.location="class_delete.asp?classid=" + ID;
			}
		}

		function doCalendar(sField) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=ClassListForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function Validate()
		{
			// check the startdate
			if (document.ClassListForm.startdate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.ClassListForm.startdate.value);
				if (! Ok)
				{
					alert("From date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassListForm.startdate.focus();
					return;
				}
			}
			// check the enddate
			if (document.ClassListForm.enddate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.ClassListForm.enddate.value);
				if (! Ok)
				{
					alert("To date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassListForm.enddate.focus();
					return;
				}
			}
			document.ClassListForm.submit();
		}

	//-->
	</script>

</head>

<body>

 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Recreation: Class\Event Management</strong></font><br />
		<!--<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
	</p>
	<!--END: PAGE TITLE-->


	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
		<fieldset class="filterselection">
		<legend class="filterselection">Filter Options</legend>
		<p>
			<form name="ClassListForm" method="post" action="class_list.asp">
			<table>
			<tr>
				<td>Season: </td>
				<td>
					<% ShowSeasonFilterPicks iClassSeasonId ' In class_global_functions.asp %>
				</td>
			</tr>
			<tr>
				<td>Status:</td>
				<td>
					<% DisplayStatusSelect iStatusid %>
				</td>
			</tr>
			<tr>
				<td>Type:</td>
				<td>
					<% DisplayTypeSelect iClasstypeid %>
				</td>
			</tr>
			<tr>
				<td>Category:</td>
				<td>
					<% DisplayCategorySelect iCategoryid  %>
				</td>
			</tr>
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
						<!--<option value="enddate"<%If iDatefilter = "enddate" Then
														response.write "selected=""selected"" "
													End If%>>End Date</option>-->
						<option value="publishstartdate"<%If iDatefilter = "publishstartdate" Then
														response.write "selected=""selected"" "
													End If%>>Publish Start Date</option>
						<!--<option value="publishenddate"<%If iDatefilter = "publishenddate" Then
														response.write "selected=""selected"" "
													End If%>>Publish End Date</option>-->
						<option value="registrationstartdate"<%If iDatefilter = "registrationstartdate" Then
														response.write "selected=""selected"" "
													End If%>>Registration Start Date</option>
						<!--<option value="registrationenddate"<%If iDatefilter = "registrationenddate" Then
														response.write "selected=""selected"" "
													End If%>>Registration End Date</option>-->
						<!--<option value="evaluationdate"<%If iDatefilter = "evaluationdate" Then
														response.write "selected=""selected"" "
													End If%>>Evaluation Date</option>-->
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
		<a href="class_list.asp?showdates=<%=sShowDates2%>" ><%=sShowDatesText%></a>
	</p>

	<!--BEGIN: CLASS LIST-->

	<% DisplayClassEvents session("orgid"), iStatusid, iClassTypeId, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId %>

	<!--END: CLASS LIST-->
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

'--------------------------------------------------------------------------------------------------
' SUB DisplayClassEvents( iorgid, iStatusid, iClassTypeId, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEvents( iorgid, iStatusid, iClassTypeId, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId )
	Dim sSQL, sWhere, oClasslist, sFrom

	sWhere = ""
	sFrom = ""

	If clng(iClassSeasonId) <> clng(0) Then 
		sWhere = sWhere & " and C.classseasonid = " & iClassSeasonId
	End If 
	If clng(iStatusid) <> 0 Then 
		sWhere = sWhere & " and C.statusid = " & iStatusid
	End If 
	If clng(iClassTypeId) <> 0 Then
		sWhere = sWhere & " and C.classtypeid = " & iClassTypeId
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


	' GET CLASS\EVENTS FOR ORG that are either single events or are series parents
	sSQL = "Select C.classid, C.classname, C.isparent, T.classtypename, C.startdate, C.registrationstartdate, C.publishstartdate,"
	sSql = sSql & " S.statusname " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S " & sFrom
	sSql = sSql & " Where C.classtypeid = T.classtypeid and C.statusid = S.statusid " & sWhere
	sSql = sSql & " and C.orgid = " & iorgid & " AND C.parentclassid is null ORDER BY C.classname"

'	response.write sSql

	Set oClasslist = Server.CreateObject("ADODB.Recordset")
	oClasslist.Open sSQL, Application("DSN"), 0, 1


	If NOT oClasslist.EOF Then

		' DRAW TABLE WITH CLASSES LISTED
		response.write vbcrlf & "<div class=""shadow"">"
		response.write vbcrlf & "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""instructortable style-alternate sortable-onload-2"" >" ' -onload-2
		
		' HEADER ROW
		response.write vbcrlf & "<tr>"
		response.write "<th>&nbsp;</th><th class=""sortable"" >Class Name</th><th>Type</th><th class=""sortable"" >Status</th><th class=""sortable"">Publish Date</th><th class=""sortable"">Registration Date</th><th class=""sortable"">Start Date</th>"
		response.write "</tr>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oClasslist.EOF

			' DISPLAY INFORMATION
			response.write vbcrlf & "<tr><td nowrap><a href=""edit_class.asp?classid=" & oClasslist("classid") & """>Edit</a> | <a href=""javascript:deleteconfirm(" & oClasslist("classid") & ", '" & oClasslist("classname") & "')"">Delete</a></td>"
			response.write "<td class=""classname""><span class=""classname"">" & oClasslist("classname") & "</span></td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oClasslist("classtypename") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oClasslist("statusname") & "</td>"
			response.write "<td align=""center"">" & oClasslist("publishstartdate") & "</td>"
			response.write "<td align=""center"">" & oClasslist("registrationstartdate") & "</td>"
			response.write "<td align=""center"">" & oClasslist("startdate") & "</td>"
			response.write "</tr>"

			' DISPLAY CHILD CLASS\EVENTS
			If oClasslist("isparent") Then
				DisplayChildClassEvents iorgid, oClasslist("classid")
			End If

			oClasslist.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

		oClasslist.close
		Set oClasslist = Nothing 
	
	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write "<font color=red><b>There are no classes\events created.</b></font>"
	
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCHILDCLASSEVENTS(IPARENTID)
'--------------------------------------------------------------------------------------------------
Sub DisplayChildClassEvents( iorgid, iparentid )
	Dim sSql, oClasslist

	' GET ALL CLASS\EVENTS FOR ORG
	sSQL = "Select classid, classname, isparent, classtypename, startdate, registrationstartdate, publishstartdate, statusname " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S "
	sSql = sSql & " Where C.classtypeid = T.classtypeid and C.statusid = S.statusid and C.orgid = " & iorgid & " AND parentclassid = " & iparentid & " ORDER BY publishstartdate"

'	sSQL = "Select * FROM egov_class INNER JOIN egov_registration_option ON egov_class.optionid=egov_registration_option.optionid INNER JOIN egov_class_type ON egov_class.classtypeid = egov_class_type.classtypeid where orgid = '" & iorgid & "' AND parentclassid='" & iparentid & "' ORDER BY noenddate DESC,publishstartdate"
	Set oClasslist = Server.CreateObject("ADODB.Recordset")
	oClasslist.Open sSQL, Application("DSN"), 0, 1


	If NOT oClasslist.EOF Then

		' LOOP THRU AND DISPLAY CHILD CLASS\EVENTS
		Do While Not oClasslist.EOF

			' DISPLAY CHILD INFORMATION
			response.write vbcrlf & "<tr>"
			response.write "<td nowrap=""nowrap""><a href=""edit_class.asp?classid=" & oClasslist("classid") & """>Edit</a> | <a href=""javascript:deleteconfirm(" & oClasslist("classid") & ", '" & oClasslist("classname") & "')"">Delete</a></td>"
			response.write "<td class=""classname""><span class=""classname"">" & oClasslist("classname") & "</span></td>" 
			response.write "<td nowrap=""nowrap"" align=""center"">" & oClasslist("classtypename") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oClasslist("statusname") & "</td>"
			response.write "<td align=""center"">" & oClasslist("publishstartdate") & "</td>"
			response.write "<td align=""center"">" & oClasslist("registrationstartdate") & "</td>"
			response.write "<td align=""center"">" & oClasslist("startdate") & "</td>"
			response.write "</tr>"

			oClasslist.MoveNext
		Loop 

		' FREE OBJECTS
		oClasslist.close
		Set oClasslist = Nothing 
	
	End If

End Sub


%>

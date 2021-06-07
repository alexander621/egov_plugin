<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: roster_list.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0 01/17/06	John Stullenberger - Initial Version
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
' 1.2	11/01/06	Steve Loar - Added link to toggle View All and View Upcoming Only
' 1.3 03/30/08 David Boyer - Added ability to print more than one attendance sheet at a time.
' 1.4 06/09/08 David Boyer - Added Download Selected Reosters (excel download)
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iStatusid, iCategoryid, iClasstypeid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId
 Dim sSearchName, sSearchActivity, iInstructorId

'Check to see if the feature is offline
 if isFeatureOffline("activities") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"registration") then
    if not userhaspermission(session("userid"),"quick registration") then
       response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 Dim sShowDatesText, sShowDates, sShowDates2, bFilter, bQuickLoad

 If request("quick") <> "" Then 
    bQuickLoad = True
 Else
    bQuickLoad = False 
 End If 

 If request("showdates") = "" Or clng(request("showdates")) = clng(1) Then
    sShowDates     = 1
    sShowDates2    = 2
    sShowDatesText = "Show All"
 Else
    sShowDates     = 2
    sShowDates2    = 1
    sShowDatesText = "Show Upcoming Only"
 End If 

 bFilter = False 

 If request("classseasonid") = "" Then 
    iClassSeasonId = GetRosterSeasonId()
 Else
    iClassSeasonId = clng(request("classseasonid"))
    bFilter = True
 End If 

 If request("instructorid") = "" Then
    iInstructorId = CLng(0)
 Else
    iInstructorId = CLng(request("instructorid"))
 End If

 if request("supervisorid") = "" then
    iSupervisorId = CLng(0)
 else
    iSupervisorId = CLng(request("supervisorid"))
 end if

 If request("categoryid") = "" or clng(request("categoryid")) = 0 Then 
   	iCategoryid = 0
 Else
   	iCategoryid = request("categoryid")
 	  bFilter = True
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

 if iDatefilter = "alldates" then
	   sStartdate = ""
   	sEnddate = ""
   	'bFilter = False 
 end if

 if request("orderby") = "" then
 	  iOrderBy = ""
 else
   	iOrderBy = request("orderby")
 end if

 sSearchName     = request("searchname")
 sSearchActivity = request("searchactivity")

'if all date choices are blank, give them the current published classes and events
 if sShowDates = 1 AND bFilter = False then
   	sDefaultRange = " and (('' + convert(char(8),getdate(),112) + '' >= publishstartdate and '' + convert(char(8),getdate(),112) + '' <= publishenddate) Or publishstartdate > '' + convert(char(8),getdate(),112) + '' Or publishstartdate is null) "
 else
	   sDefaultRange = ""
 end if

 if bFilter = True then
	   sDefaultRange = ""
 end if
%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="javascript" src="tablesort.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.7.2.min.js"></script>

	<script language="javascript">
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
			document.ClasListsForm.submit();
		}

		function ViewCart()
		{
			location.href='class_cart.asp';
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
					classid = $("#classno" + x).val();
					if (checkAttr)
						$("#classid" + classid).prop('checked', true);
					else
						$("#classid" + classid).prop("checked", false);
				}
			}
		}

		function RosterPrint() 
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
				document.getElementById("classForm").action = "roster_list_print.asp";
				document.classForm.submit();
			}
			else  
			{
				alert("Please select at least one class to Print from the list.");
			}
		}

		function AttendancePrint() 
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
				document.getElementById("classForm").action = "attendance_sheet.asp";
				document.classForm.submit();
			}
			else
			{
				alert("Please select at least one class to Print from the list.");
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

		function rosterEmail()
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
				document.getElementById("classForm").action = "MultiClassEmail.asp"
				document.classForm.submit();
			}
			else
			{
				alert("Please select at least one class from the list.");
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
<% If CartHasItems() Then %>
	<div id="topbuttons">
		<input type="button" name="viewcart" class="button" value="View Cart" onclick="ViewCart();" />
	</div>
<%	End If %>

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Recreation: Class\Event Rosters and Registration</strong></font><br />
		<!--<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
	</p>
	<!--END: PAGE TITLE-->

	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
	 	<fieldset class="filterselection">
		   <legend class="filterselection">Search Options</legend>
     <p>
  	<form name="ClassListForm" method="post" action="roster_list.asp">
		<table border="0" cellspacing="0" cellpadding="0">
		  	<tr>
				<td>Season: </td>
         <td>
             <% ShowSeasonFilterPicks iClassSeasonId ' In class_global_functions.asp %>
         </td>
     		</tr>		
  			<tr>
     			 <td>Category:</td>
				     <td colspan="2">
        					<% DisplayCategorySelect iCategoryid  %>
     				</td>
  			</tr>
		  	<tr>
				     <td>Instructor:</td>
     				<td colspan="2">
					        <% ShowInstructorPicks iInstructorId %>
     				</td>
		  	</tr>
    	<tr>
				     <td>Supervisor:</td>
     				<td colspan="2">
					        <% ShowActivitySupervisorPicks iSupervisorId%>
     				</td>
		</tr>
		<tr>
				    <td>Name Like:</td>
    				<td colspan="2"><input type="text" name="searchname" value="<%=sSearchName%>" size="75" maxlength="255" /></td>
  		</tr>
  		<tr>
		    		<td>Activity No:</td>
    				<td colspan="2"><input type="text" name="searchactivity" value="<%=sSearchActivity%>" size="10" maxlength="10" /></td>
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
        						<option value="publishstartdate"<%If iDatefilter = "publishstartdate" Then
									      				response.write "selected=""selected"" "
      													End If%>>Publish Start Date</option>
        						<option value="registrationstartdate"<%If iDatefilter = "registrationstartdate" Then
									     					response.write "selected=""selected"" "
      													End If%>>Registration Start Date</option>
       					</select>
    				</td>
     			<td>
         			From: <input type="text" name="startdate" value="<%=sStartdate%>" />
				      	 <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
     	 					To: <input type="text" name="enddate" value="<%=sEnddate%>"/>
     	 					<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
      		</td>
   	</tr>
	<tr>
		<td>Order By:</td>
		<td>
			<% showOrderByFilterPicks iOrderBy 'In class_global_functions.asp %>
     	</td>
  	</tr>
  		<tr>
		     	<td>&nbsp;</td>
     			<td colspan="2"><input class="button" type="submit" value="Refresh Results" /></td>
  		</tr>
  </table>
	</form>
 	</p>
 	</fieldset>

	</div>
	<!--END: FILTER SELECTION-->

	<div id="rosterlistbuttons">
		<!-- <a href="roster_list.asp?showdates=<%=sShowDates2%>"><%=sShowDatesText%></a> -->
		<input type="button" value="<%=sShowDatesText%>" class="button" onclick="javascript:location.href='roster_list.asp?showdates=<%=sShowDates2%>'" />
		<% 
		If Not bQuickLoad Then 
			'response.write "&nbsp;" 
			response.write "<input type=""button"" value=""Print Selected Rosters"" class=""button"" onclick=""RosterPrint()"" />&nbsp;"
			response.write "<input type=""button"" value=""Print Selected Attendance Sheets"" class=""button"" onclick=""AttendancePrint()"" />&nbsp;"
			response.write "<input type=""button"" value=""Download Selected Rosters"" class=""button"" onclick=""RosterDownload()"" />&nbsp;"
			response.write "<input type=""button"" value=""Email Selected Rosters"" class=""button"" onclick=""rosterEmail()"" /> "
		End If 
		%>
	</div>

	<!--BEGIN: CLASS LIST-->

	<% 
		If Not bQuickLoad Then 
			response.write vbcrlf & "<form id=""classForm"" name=""classForm"" method=""post"" action=""roster_list_print.asp"">"
			DisplayClassEvents session("orgid"), iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId, sSearchName, sSearchActivity, iInstructorId, iSupervisorId
			response.write vbcrlf & "</form>"
		Else
			response.write "<strong>To view the class list, select from the filter options above then click the &quot;Refresh Results&quot; button.</strong>"
		End If 
	%>

	<!--END: CLASS LIST-->
	</div>
</div>

<form name="PrintForm" method="post" action="">
	<input type="hidden" name="selectedClassid" value="" />
</form>

<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void DisplayClassEvents iorgid, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId, sSearchName, sSearchActivity, iInstructorId, iSupervisorId
'------------------------------------------------------------------------------
Sub DisplayClassEvents( ByVal iorgid, ByVal iCategoryid, ByVal iDatefilter, ByVal sStartdate, ByVal sEnddate, ByVal sDefaultRange, ByVal iClassSeasonId, ByVal sSearchName, ByVal sSearchActivity, ByVal iInstructorId, ByVal iSupervisorId )
	Dim sSql, sWhere, oRs, sFrom

	sWhere = ""
	sFrom = ""

	If clng(iClassSeasonId) <> clng(0) Then 
		sWhere = sWhere & " and C.classseasonid = " & iClassSeasonId
	End If 

	If clng(iCategoryid) <> 0 Then
		sWhere = sWhere & " and CC.classid = C.classid and CC.categoryid = " & iCategoryid
		sFrom = sFrom & ", egov_class_category_to_class CC"
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

	If sSearchActivity <> "" Then 
		sSearchActivity = dbsafe(sSearchActivity)
		sFrom = sFrom & ", egov_class_time CT"
		'sWhere = sWhere & " and CT.classid = C.classid and Lower(CT.activityno) like Lower('%" & sSearchActivity & "%') "
		sWhere = sWhere & " and CT.classid = C.classid and Lower(CT.activityno) = Lower('" & sSearchActivity & "') "
	End If 

	If CLng(iInstructorId) > CLng(0) Then 
		If sSearchActivity = "" Then 
			sFrom  = sFrom & ", egov_class_time CT"
			sWhere = sWhere & " and CT.classid = C.classid "
		End If 
		sWhere = sWhere & " and CT.instructorid = " & iInstructorId
	End If 

	If CLng(iSupervisorId) > CLng(0) Then 
		sWhere = sWhere & " and C.supervisorid = " & iSupervisorId
	End If 

	'GET CLASS\EVENTS FOR ORG that are active and can be purchased
	sSql = "SELECT DISTINCT C.classid, C.classname, C.isparent, T.classtypename, C.startdate, C.registrationstartdate, C.publishstartdate,"
	sSql = sSql & " S.statusname, c.supervisorid " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S, egov_registration_option RO " & sFrom
	sSql = sSql & " WHERE C.classtypeid = T.classtypeid AND C.statusid = S.statusid "
	sSql = sSql & " AND S.statusname = 'ACTIVE' AND RO.optionid = C.optionid AND C.orgid = " & iorgid
	sSql = sSql & " AND C.isregatta = 0 AND RO.canpurchase = 1 " & sWhere

	'Setup the ORDER BY
	If UCase(iOrderBy) = "STARTDATE" Then 
		lcl_order_by = "C.startdate, C.classname "
	Else 
		lcl_order_by = "C.classname "
	End If 

	sSql = sSql & " ORDER BY " & lcl_order_by

'	If UserIsRootAdmin( Session("UserId") ) Then 
'		response.write sSql & "<br />"
'	End If 
'	response.end

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then

		'DRAW TABLE WITH CLASSES LISTED
		'response.write vbcrlf & "<div class=""shadow"">"
		'class=""instructortable style-alternate sortable-onload-2""
		response.write vbcrlf & "<table id=""rosterlist"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""instructortable style-alternate sortable-onload-2"">"
		
		'HEADER ROW
		response.write vbcrlf & "<tr>"
		response.write "<th>"
		response.write "<input type=""checkbox"" id=""toggleselects"" name=""toggleselects"" onclick=""toggleChecks()"" />"
		response.write "</th>"
		response.write "<th>Class Name</th>"
		response.write "<th>Type</th>"
		response.write "<th>Start<br />Date</th>"
		response.write "<th>Available<br />Activity<br />Count</th>"
		response.write "<th>Total<br />Max</th>"
		response.write "<th>Total<br />Enrld</th>"
		response.write "<th>%<br />Full</th>"
		response.write "<th>Total<br />Waiting</th>"
		response.write "</tr>"
		'response.write "<tbody>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oRs.EOF
  			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then 
				response.write " class=""altrow"" "
			End If 

			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			' selection check box
			response.write "<td>"
			response.write "<input type=""checkbox"" name=""classid"" id=""classid" & oRs("classid") & """ value=""" & oRs("classid") & """ />"
			response.write "<input type=""hidden"" name=""classno" & iRowCount & """ id=""classno" & iRowCount & """  value=""" & oRs("classid") & """ />"
			response.write "</td>"

			' class name
			response.write "<td class=""classname"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">"
			response.write oRs("classname")

			If oRs("isparent") And UCase(oRs("classtypename")) = "SERIES" Then 
				response.write " (Series Purchase)" 
			End If 

			response.write "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & oRs("classtypename") & "</td>"
			response.write "<td align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & oRs("startdate") & "</td>"
			response.write "<td align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & GetActivityCount( oRs("classid") ) & "</td>"
			response.write "<td align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & fnIsNull(GetClassTotalMax( oRs("classid") ),"n/a" ) & "</td>"
			response.write "<td align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & fnIsNull(GetClassTotalEnrld( oRs("classid") ),"n/a" ) & "</td>"
			response.write "<td align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & fnIsNull(GetClassPercentEnrld( GetClassTotalMax( oRs("classid") ), GetClassTotalEnrld( oRs("classid") ) ),"n/a" ) & "</td>"
			response.write "<td align=""center"" onClick=""location.href='class_offerings.asp?classid=" & oRs("classid") & "';"">" & fnIsNull(GetClassTotalWait( oRs("classid") ),"n/a" ) & "</td>"
			response.write "</tr>"

			oRs.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		'response.write "</tbody>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<input type=""hidden"" name=""maxclasscount"" id=""maxclasscount"" value=""" & iRowCount & """ />"
		'response.write vbcrlf & "</div>"

		oRs.Close
		Set oRs = Nothing 
	
	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write vbcrlf & "<font color=""red""><b>There are no classes\events created.</b></font>"
		response.write vbcrlf & "<input type=""hidden"" name=""maxclasscount"" id=""maxclasscount"" value=""0"" />"
	
	End If

End Sub


'------------------------------------------------------------------------------
' void DisplayChildClassEvents iorgid, iparentid
'------------------------------------------------------------------------------
Sub DisplayChildClassEvents( ByVal iorgid, ByVal iparentid )
	Dim sSql, oRs

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT classid, classname, isparent, classtypename, startdate, registrationstartdate, publishstartdate, statusname " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S "
	sSql = sSql & " Where C.classtypeid = T.classtypeid and C.statusid = S.statusid and C.orgid = " & iorgid & " AND parentclassid = " & iparentid & " ORDER BY publishstartdate"

'	sSql = "Select * FROM egov_class INNER JOIN egov_registration_option ON egov_class.optionid=egov_registration_option.optionid INNER JOIN egov_class_type ON egov_class.classtypeid = egov_class_type.classtypeid where orgid = '" & iorgid & "' AND parentclassid='" & iparentid & "' ORDER BY noenddate DESC,publishstartdate"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1


	If Not oRs.EOF Then

		' LOOP THRU AND DISPLAY CHILD CLASS\EVENTS
		Do While Not oRs.EOF
			' DISPLAY CHILD INFORMATION
			response.write vbcrlf & "<tr>"
			response.write "<td nowrap=""nowrap""><a href=""edit_class.asp?classid=" & oRs("classid") & """>Edit</a> | <a href=""javascript:deleteconfirm(" & oRs("classid") & ", '" & oRs("classname") & "')"">Delete</a></td>"
			response.write "<td class=""classname""><span class=""classname"">" & oRs("classname") & "</span></td>" 
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("classtypename") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("statusname") & "</td>"
			response.write "<td align=""center"">" & oRs("publishstartdate") & "</td>"
			response.write "<td align=""center"">" & oRs("registrationstartdate") & "</td>"
			response.write "<td align=""center"">" & oRs("startdate") & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 

		oRs.Close
		Set oRs = Nothing 
	
	End If

End Sub


'------------------------------------------------------------------------------
' integer GetClassTotalMax( iClassId )
'------------------------------------------------------------------------------
Function GetClassTotalMax( ByVal iClassId )
	Dim sSql, oMaxTotal

	sSql = "SELECT SUM(max) AS maxtotal FROM egov_class_time WHERE classid = " & iClassId & " GROUP BY classid"

	Set oMaxTotal = Server.CreateObject("ADODB.Recordset")
	oMaxTotal.Open sSql, Application("DSN"), 0, 1

	If Not oMaxTotal.EOF Then 
		If Not IsNull(oMaxTotal("maxtotal")) Then 
			GetClassTotalMax = clng(oMaxTotal("maxtotal")) 
		Else
			GetClassTotalMax = Null 
		End If 
	Else
		GetClassTotalMax = Null 
	End If 

	oMaxTotal.Close
	Set oMaxTotal = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetClassTotalEnrld( iClassId )
'------------------------------------------------------------------------------
Function GetClassTotalEnrld( ByVal iClassId )
	Dim sSql, oEnrldTotal

	sSql = "SELECT SUM(enrollmentsize) AS enrldtotal FROM egov_class_time WHERE classid = " & iClassId & " GROUP BY classid"

	Set oEnrldTotal = Server.CreateObject("ADODB.Recordset")
	oEnrldTotal.Open sSql, Application("DSN"), 0, 1

	If Not oEnrldTotal.EOF Then 
		If Not IsNull(oEnrldTotal("enrldtotal")) Then 
			GetClassTotalEnrld = clng(oEnrldTotal("enrldtotal")) 
		Else
			GetClassTotalEnrld = Null 
		End If 
	Else
		GetClassTotalEnrld = Null 
	End If 

	oEnrldTotal.Close
	Set oEnrldTotal = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetClassTotalEnrld( iClassId )
'------------------------------------------------------------------------------
Function GetClassTotalWait( ByVal iClassId )
	Dim sSql, oWaitTotal

	sSql = "SELECT SUM(waitlistsize) AS waittotal FROM egov_class_time WHERE classid = " & iClassId & " GROUP BY classid"

	Set oWaitTotal = Server.CreateObject("ADODB.Recordset")
	oWaitTotal.Open sSql, Application("DSN"), 0, 1

	If Not oWaitTotal.EOF Then 
		If Not IsNull(oWaitTotal("waittotal")) Then 
			GetClassTotalWait = clng(oWaitTotal("waittotal")) 
		Else
			GetClassTotalWait = null 
		End If 
	Else
		GetClassTotalWait = Null  
	End If 

	oWaitTotal.Close
	Set oWaitTotal = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetClassPercentEnrld( iTotalmax, iEnrolledTotal )
'------------------------------------------------------------------------------
Function GetClassPercentEnrld( ByVal iTotalmax, ByVal iEnrolledTotal )

	If Not IsNull(iTotalmax) Then 
		If clng(iTotalmax) > clng(0) Then
			GetClassPercentEnrld = Int((iEnrolledTotal / iTotalmax) * 100) & "%"
		Else
			GetClassPercentEnrld = Null 
		End If 
	Else
		GetClassPercentEnrld = Null
	End If 

End Function 


'------------------------------------------------------------------------------
' void ShowInstructorPicks( iInstructorId )
'------------------------------------------------------------------------------
Sub ShowInstructorPicks( ByVal iInstructorId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_class_instructor WHERE orgid = " & SESSION("ORGID") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""instructorid"">"
		response.write vbcrlf & "<option value=""0"" >All Instructors</option>"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("instructorid") & """ "  
			If clng(iInstructorId) = clng(oRs("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("lastname") & ", " & oRs("firstname")& "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub

'------------------------------------------------------------------------------
' void ShowActivitySupervisorPicks iSupervisorId
'--------------------------------------------------------------------
Sub ShowActivitySupervisorPicks( ByVal iSupervisorId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname + ' ' + lastname AS name "
	sSql = sSql & " FROM users WHERE isclasssupervisor = 1 AND orgid = " & session("orgid")
	sSql = sSql & " AND userid IN (SELECT DISTINCT supervisorid FROM egov_class "
	sSql = sSql & " WHERE orgid = " & session("orgid") & ") "
	sSql = sSql & " ORDER BY lastname, firstname "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""supervisorid"">"

		If CLng(iSupervisorId) = CLng(0) Then 
		   lcl_selected_all = " selected=""selected"""
		Else 
		   lcl_selected_all = ""
		End If 

		response.write vbcrlf & "<option value=""0""" & lcl_selected_all & ">All Supervisors</option>"

		Do While Not oRs.EOF
			If CLng(iSupervisorId) = CLng(oRs("userid")) Then 
				lcl_selected = " selected=""selected"""
			Else 
				lcl_selected = ""
			End If 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """" & lcl_selected & ">" & oRs("name") & "</option>"
			oRs.MoveNext
		loop
		response.write vbcrlf & "</select>"
	End If

	oRs.Close
	Set oRs = Nothing

End Sub




%>

<!DOCTYPE html>
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
' 1.0  01/17/06	 JOHN STULLENBERGER - INITIAL VERSION
' 1.1	 10/11/06	 Steve Loar - Security, Header and nav changed
' 1.2	 11/01/06	 Steve Loar - Added link to toggle View All and View Upcoming Only
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStatusid, iCategoryid, iClasstypeid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId
Dim sSearchName, sSearchActivity, sShowDatesText, sShowDates, sShowDates2, bFilter, bQuickLoad

'Check to see if the feature is offline
if isFeatureOffline("activities") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "manage classes" ) Then
   If Not UserHasPermission( Session("UserId"), "quick load class list" ) Then
      response.redirect sLevel & "permissiondenied.asp"
   End If 
End If 

If request("showdates") = "" Or clng(request("showdates")) = clng(1) Then
   sShowDates     = 1
   sShowDates2    = 2
   sShowDatesText = "Show All"
Else
   sShowDates     = 2
   sShowDates2    = 1
   sShowDatesText = "Show Upcoming Classes and Events Only"
End If 

bFilter = False 

If request("quick") <> "" Then 
   bQuickLoad = True
   sShowDates     = 2
   sShowDates2    = 1
   sShowDatesText = "Show Upcoming Classes and Events"
Else
   bQuickLoad     = False 
End If

If request("classseasonid") = "" Then 
   iClassSeasonId = GetRosterSeasonId()
Else
   iClassSeasonId = clng(request("classseasonid"))
   bFilter = True
End If 

If request("statusid") = "" or clng(request("statusid")) = 0 Then 
   iStatusid = 0
Else
   iStatusid = request("statusid")
   bFilter   = True
End If 

If request("classtypeid") = "" or clng(request("classtypeid")) = 0 Then 
	iClasstypeid = 0
Else
	iClasstypeid = request("classtypeid")
	bFilter = True
End If 

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
If iDatefilter = "alldates" Then
	sStartdate = ""
	sEnddate = ""
'	bFilter = False 
End If

' if all date choices are blank, give them the current published classes and events
If sShowDates = 1 And bFilter = False Then
	sDefaultRange = " and (('' + convert(char(8),getdate(),112) + '' >= publishstartdate and '' + convert(char(8),getdate(),112) + '' <= publishenddate) Or publishstartdate > '' + convert(char(8),getdate(),112) + '' Or publishstartdate is null) "
Else 
	sDefaultRange = ""
End If 

If bFilter Then
	sDefaultRange = ""
End If 

sSearchName = request("searchname")

sSearchActivity = request("searchactivity")

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../recreation/facility.css" />
	<link rel="stylesheet" href="classes.css" />

	<script src="../scripts/ajaxLib.js"></script>    
	<script src="../scripts/modules.js"></script>    

	<script>
	<!--

		function deleteconfirm( ID ) 
		{
			lcl_name = document.getElementById("classname" + ID).innerHTML;

//			if(confirm('Do you wish to delete ' + sName + '?')) 
			if(confirm('Do you wish to delete ' + lcl_name + '?')) 
			{
  				// Fire off AJAX check of registrants. Do not delete if there is anyone on the classlist, even dropped
		   		doAjax('check_class_for_deletion.asp', 'classid=' + ID, 'ClassCheckReturn', 'get', '0');
  				//window.location="class_delete.asp?classid=" + ID;
			}
		}

		function ClassCheckReturn( sResult )
		{
			//alert( sResult );
			if (sResult != "KEEPCLASS")
			{
				//alert('Successful');
				window.location="class_delete.asp?classid=" + sResult;
			}
			else 
			{
				alert("This class cannot be deleted because there are still people on its roster.");
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
			document.ClassListForm.action="class_list.asp";
			document.ClassListForm.submit();
		}

		function exportToCatalog()
		{
			document.ClassListForm.action="catalogexport.asp";
			document.ClassListForm.submit();
			document.ClassListForm.action="class_list.asp";
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
		<font size="+1"><strong>Recreation: Class\Event Management</strong></font><br />
		<!--<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
	</p>
	<!--END: PAGE TITLE-->


	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
		<fieldset class="filterselection">
		<legend class="filterselection">Search Options&nbsp;</legend>
		<p>
			<form name="ClassListForm" method="post" action="class_list.asp">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td>Season: </td>
				<td colspan="2">
					<% ShowSeasonFilterPicks iClassSeasonId ' In class_global_functions.asp %>
				</td>
			</tr>
			<tr>
				<td>Status:</td>
				<td colspan="2">
					<% DisplayStatusSelect iStatusid %>
				</td>
			</tr>
			<tr>
				<td>Type:</td>
				<td colspan="2">
					<% DisplayTypeSelect iClasstypeid %>
				</td>
			</tr>
			<tr>
				<td>Category:</td>
				<td colspan="2">
					<% DisplayCategorySelectAll iCategoryid  %>
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
				From: <input type="text" name="startdate" value="<%=sStartdate%>" /> <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span> To: <input type="text" name="enddate" value="<%=sEnddate%>" /> <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
			</td>

			</tr>
			<tr>
				<td>&nbsp;</td>
				<td colspan="2">
					<input class="button" type="submit" value="Refresh Results" />
<%					If orghasfeature( "catalog export" ) Then			
						'If UserHasPermission( Session("UserId"), "catalog export" ) Then
%>
							&nbsp;&nbsp;<input class="button" type="button" value="Catalog Export" onclick="exportToCatalog();" />
<%						'End If 
					End If								
%>
				</td>
			</tr>
			</table>
			</form>
		</p>
		</fieldset>
	</div>
	<!--END: FILTER SELECTION-->

	<p>
		<!--<a href="class_list.asp?showdates=<%'sShowDates2%>" ><%'sShowDatesText%></a> -->
  <input type="button" name="showDates" id="showDates" value="<%=sShowDatesText%>" class="button" onclick="location.href='class_list.asp?showdates=<%=sShowDates2%>'" />


	</p>

	<!--BEGIN: CLASS LIST-->

	<%	If Not bQuickLoad Then 
			DisplayClassEvents session("orgid"), iStatusid, iClassTypeId, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId, sSearchName, sSearchActivity 
		Else
			response.write "<strong>To view the class/event list, select from the filter options above then click the &quot;Refresh Results&quot; button.</strong>"
		End If 
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
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB DisplayClassEvents( iorgid, iStatusid, iClassTypeId, iCategoryid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId, sSearchName, sSearchActivity )
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEvents( ByVal iorgid, ByVal iStatusid, ByVal iClassTypeId, ByVal iCategoryid, ByVal iDatefilter, ByVal sStartdate, ByVal sEnddate, ByVal sDefaultRange, ByVal iClassSeasonId, ByVal sSearchName, ByVal sSearchActivity )
	Dim sSql, sWhere, oRs, sFrom, iRowCount, sRowClass, sLink

	sWhere = ""
	sFrom = ""
	iRowCount = 0

	If CLng(iClassSeasonId) <> CLng(0) Then 
		sWhere = sWhere & " and C.classseasonid = " & iClassSeasonId
	End If 
	If CLng(iStatusid) <> 0 Then 
		sWhere = sWhere & " and C.statusid = " & iStatusid
	End If 
	If CLng(iClassTypeId) > CLng(0) Then
		sWhere = sWhere & " and C.classtypeid = " & iClassTypeId
	End If 
	If CLng(iCategoryid) > CLng(0) Then
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


	' GET CLASS\EVENTS FOR ORG that are either single events or are series parents
	sSql = "Select C.classid, C.classname, C.isparent, T.classtypename, C.startdate, C.registrationstartdate, C.publishstartdate,"
	sSql = sSql & " S.statusname, C.isregatta " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S " & sFrom
	sSql = sSql & " Where C.classtypeid = T.classtypeid and C.statusid = S.statusid " & sWhere
	sSql = sSql & " and C.orgid = " & iorgid & " AND C.parentclassid is null ORDER BY C.classname"

'	If UserIsRootAdmin( Session("UserId") ) Then 
'		response.write sSql & "<br />"
'	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then

		' DRAW TABLE WITH CLASSES LISTED
		response.write vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin"">" 
		
		' HEADER ROW
		response.write vbcrlf & "<tr><th>Class Name</th><th style=""text-align: center"">Type</th><th style=""text-align: center"">Status</th>"
		response.write "<th style=""text-align: center"">Publish Date</th><th style=""text-align: center"">Registration Date</th><th style=""text-align: center"">Start Date</th><th style=""text-align: center"">&nbsp;</th></tr>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oRs.EOF

			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sRowClass = " class=""altrow"" "
			Else
				sRowClass = ""
			End If 

			If oRs("isregatta") Then
				sLink = "regattaeventedit.asp?classid=" & oRs("classid")
			Else
				sLink = "edit_class.asp?classid=" & oRs("classid")
			End If 

			' DISPLAY INFORMATION
			response.write vbcrlf & "<tr id=""" & iRowCount & """" & sRowClass & " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			response.write "<td title='Click to Edit' onClick=""location.href='" & sLink & "';"" class=""classname"">&nbsp;<span class=""classname"" id=""classname" & oRs("classid") & """>" & oRs("classname") & "</span></td>"
			response.write "<td title='Click to Edit' onClick=""location.href='" & sLink & "';"" nowrap=""nowrap"" align=""center"">" & oRs("classtypename") & "</td>"
			response.write "<td title='Click to Edit' onClick=""location.href='" & sLink & "';"" nowrap=""nowrap"" align=""center"">" & oRs("statusname") & "</td>"
			response.write "<td title='Click to Edit' onClick=""location.href='" & sLink & "';"" align=""center"">" & oRs("publishstartdate") & "</td>"
			response.write "<td title='Click to Edit' onClick=""location.href='" & sLink & "';"" align=""center"">" & oRs("registrationstartdate") & "</td>"
			response.write "<td title='Click to Edit' onClick=""location.href='" & sLink & "';"" align=""center"">" & oRs("startdate") & "</td>"
			response.write "<td align=""center""><input type=""button"" name=""delete" & oRs("classid") & """ id=""delete" & oRs("classid") & """ value=""Delete"" class=""button"" onclick=""deleteconfirm(" & oRs("classid") & ")"" /></td>"
			response.write "</tr>"

			' DISPLAY CHILD CLASS\EVENTS
			If oRs("isparent") Then
				DisplayChildClassEvents iorgid, oRs("classid"), iRowCount
			End If

			oRs.MoveNext
			response.Flush
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

		oRs.close
		Set oRs = Nothing 
	
	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write vbcrlf & "<p><font color=""red""><b>There are no classes\events created.</b></font></p>"
	
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCHILDCLASSEVENTS(IPARENTID)
'--------------------------------------------------------------------------------------------------
Sub DisplayChildClassEvents( ByVal iorgid, ByVal iparentid, ByRef iRowCount )
	Dim sSql, oClasslist

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT classid, classname, isparent, classtypename, startdate, registrationstartdate, publishstartdate, statusname " 
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_class_status S "
	sSql = sSql & " WHERE C.classtypeid = T.classtypeid "
	sSql = sSql & " AND C.statusid = S.statusid "
	sSql = sSql & " AND C.orgid = " & iorgid
	sSql = sSql & " AND parentclassid = " & iparentid
	sSql = sSql & " ORDER BY publishstartdate"

'	sSql = "Select * FROM egov_class INNER JOIN egov_registration_option ON egov_class.optionid=egov_registration_option.optionid INNER JOIN egov_class_type ON egov_class.classtypeid = egov_class_type.classtypeid where orgid = '" & iorgid & "' AND parentclassid='" & iparentid & "' ORDER BY noenddate DESC,publishstartdate"
	Set oClasslist = Server.CreateObject("ADODB.Recordset")
	oClasslist.Open sSql, Application("DSN"), 0, 1

	If NOT oClasslist.EOF Then

		' LOOP THRU AND DISPLAY CHILD CLASS\EVENTS
		Do While Not oClasslist.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sRowClass = " class=""altrow"" "
			Else
				sRowClass = ""
			End If 
			' DISPLAY CHILD INFORMATION
			response.write "<tr id=""" & iRowCount & """" & sRowClass & " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
			response.write "    <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='edit_class.asp?classid=" & oClasslist("classid") & "';"" class=""classname"">&nbsp;<span class=""classname"" id=""classname"">" & oClasslist("classname") & "</span></td>"  & vbcrlf
			response.write "    <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='edit_class.asp?classid=" & oClasslist("classid") & "';"" nowrap=""nowrap"" align=""center"">" & oClasslist("classtypename") & "</td>" & vbcrlf
			response.write "    <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='edit_class.asp?classid=" & oClasslist("classid") & "';"" nowrap=""nowrap"" align=""center"">" & oClasslist("statusname") & "</td>" & vbcrlf
			response.write "    <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='edit_class.asp?classid=" & oClasslist("classid") & "';"" align=""center"">" & oClasslist("publishstartdate") & "</td>" & vbcrlf
			response.write "    <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='edit_class.asp?classid=" & oClasslist("classid") & "';"" align=""center"">" & oClasslist("registrationstartdate") & "</td>" & vbcrlf
			response.write "    <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='edit_class.asp?classid=" & oClasslist("classid") & "';"" align=""center"">" & oClasslist("startdate") & "</td>" & vbcrlf
   'response.write "    <td nowrap=""nowrap""><a href=""javascript:deleteconfirm(" & oClasslist("classid") & ", '" & FormatForJavaScript(oClasslist("classname")) & "')"">Delete</a></td>" & vbcrlf
			'response.write "    <td align=""center""><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""javascript:deleteconfirm(" & oClasslist("classid") & ", '" & FormatForJavaScript(oClasslist("classname")) & "')""></td>" & vbcrlf
   response.write "    <td align=""center""><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""javascript:deleteconfirm(" & oClasslist("classid") & ")""></td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			oClasslist.MoveNext
		loop 

	'FREE OBJECTS
		oClasslist.close
		Set oClasslist = Nothing 

	End If

End Sub
%>

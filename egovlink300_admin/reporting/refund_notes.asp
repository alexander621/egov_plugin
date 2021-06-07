<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
''--------------------------------------------------------------------------------------------------
''
''--------------------------------------------------------------------------------------------------
'' FILENAME: refund_notes.asp
'' AUTHOR: SteveLoar
'' CREATED: 10/16/2013
'' COPYRIGHT: Copyright 2013 eclink, inc.
''			 All Rights Reserved.
''
'' Description:  This pulls together the refund notes report. Part of a Menlo Park Project.
''
'' MODIFICATION HISTORY
'' 1.0   10/16/2013	Steve Loar - INITIAL VERSION
''
''--------------------------------------------------------------------------------------------------
''
''--------------------------------------------------------------------------------------------------

Dim iClassSeasonId, iCategoryid, fromDate, toDate, sWhereClause, sFrom, sSearchName, sClassName
Dim sActivityNo, iPaymentId, toDateDisplay, iOrderBy, sOrderBy

sWhereClause = ""
sFrom = ""
iOrderBy = 0


' INITIALIZE AND DECLARE VARIABLES'
' SPECIFY FOLDER LEVEL'
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "refund notes rpt" ) Then
	response.redirect sLevel & "../permissiondenied.asp"
End If 


' Process the parameters
If request("classseasonid") = "" Then 
	iClassSeasonId = GetRosterSeasonId()
Else
	iClassSeasonId = clng(request("classseasonid"))
End If 

If request("categoryid") = "" Or CLng(request("categoryid")) = CLng(0) Then 
	iCategoryid = CLng(0)
Else
	iCategoryid = CLng(request("categoryid"))
End If 

If request("searchname") <> "" Then 
	sSearchName = dbsafe(request("searchname"))
End If 

If request("classname") <> "" Then 
	sClassName = dbsafe(request("classname"))
End If 

If request("activityno") <> "" Then 
	sActivityNo = dbsafe(request("activityno"))
End If

If request("paymentid") = "" Then 
	iPaymentId = ""
Else
	iPaymentId = CLng(request("paymentid"))
End If 

fromDate = Request("fromDate")
toDate = Request("toDate")

If request("orderby") <> "" Then 
	iOrderBy = clng(request("orderby"))
End If 

If clng(iOrderBy) = clng(0) Then
	sOrderBy = "paymentdate"
Else
	sOrderBy = "U.userlname, U.userfname, paymentdate"
End If

' BUILD SQL WHERE CLAUSE
If iClassSeasonId > CLng(0) Then
	sWhereClause = sWhereClause & " AND C.classseasonid = " & iClassSeasonId
End If 

If iCategoryid > CLng(0) Then
	sFrom = ", egov_class_category_to_class G "
	sWhereClause = sWhereClause & " AND C.classid = G.classid AND G.categoryid = " & iCategoryid
End If 

If sSearchName <> "" Then
	sWhereClause = sWhereClause & " AND ( U.userfname LIKE '%" & sSearchName & "%' OR U.userlname LIKE '%" & sSearchName & "%' )"
End If 

If sClassName <> "" Then
	sWhereClause = sWhereClause & " AND C.classname LIKE '%" & sClassName & "%' "
End If 

If sActivityNo <> "" Then
	sWhereClause = sWhereClause & " AND T.activityno = '" & sActivityNo & "' "
End If 

If iPaymentId <> "" Then
	sWhereClause = sWhereClause & " AND P.paymentid = " & iPaymentId & " "
End If 

If fromDate <> "" Then 
	sWhereClause = sWhereClause & " AND P.paymentdate >= '" & fromDate & "' "
End If 

If toDate <> "" Then 
	toDateDisplay = toDate
	toDate = DateAdd( "d", 1, toDate )
	sWhereClause = sWhereClause & " AND P.paymentdate < '" & toDate & "' "
End If 

%>
<html lang="en">
<head>
	<meta charset="UTF-8">
  	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="reporting.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../classes/classes.css" />
	<link rel="stylesheet" href="pageprint.css" media="print" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	<script src="scripts/dates.js"></script>
	<script src="../scripts/isvaliddate.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?updatefield=' + ToFrom + '&date=' + $("#" + ToFrom ).val() + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function submitForm( viewToggle )
		{
			var okToSubmit = true;

			// check that paymentid is numeric
			if ( $("#paymentid").val() != "" )
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#paymentid").val());
				if ( !Ok )
				{
					okToSubmit = false;
					inlineMsg("paymentid",'The Receipt No needs to be numeric.',10,'paymentid');
				}
			}

			// check that fromDate and toDate are dates
			if ( $("#fromDate").val() != "" )
			{
				// check that it is a date
				if (! isValidDate($("#fromDate").val()) )
				{
					okToSubmit = false;
					inlineMsg("fromDate",'This Drop Date needs to be a date.',10,'fromDate');
				}
			}

			if ( $("#toDate").val() != "" )
			{
				// check that it is a date
				if (! isValidDate($("#toDate").val()) )
				{
					okToSubmit = false;
					inlineMsg("toDate",'This Drop Date needs to be a date.',10,'toDate');
				}
			}

			if (okToSubmit)
			{
				if ( viewToggle == 'excel' )
				{
					document.frmPFilter.action = "refund_notes_export.asp";
				}
				else
				{
					document.frmPFilter.action = "refund_notes.asp";
				}
				document.frmPFilter.submit();
			}
		}

	  //-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
</div>
<!--END: THIRD PARTY PRINT CONTROL-->
<!--BEGIN PAGE CONTENT-->

<div id="content">
	<div id="centercontent">

<form action="refund_notes.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Refund Notes Report</strong></font></td>
		</tr>
		<tr>
			<td>
				<fieldset>
					<legend><strong>Select</strong></legend>
				
					<!--BEGIN: FILTERS-->
					
					<table border="0" cellpadding="2" cellspacing="0">
						<tr>
							<td><strong>Season:</strong></td>
							<td>
								<% ShowSeasonFilterPicks iClassSeasonId ' In class_global_functions.asp %>
							</td>
						</tr>
						<tr>
							<td><strong>Category:</strong></td>
							<td>
								<% DisplayCategorySelect iCategoryid  %>
							</td>
						</tr>
						<tr>
							<td><strong> Class Name Like:</strong></td>
							<td><input type="text" name="classname" value="<%=sClassName%>" size="75" maxlength="255" /></td>
						</tr>
						<tr>
							<td><strong>Activity No:</strong></td>
							<td><input type="text" name="activityno" value="<%=sActivityNo%>" size="15" maxlength="15" /></td>
						</tr>
						<tr>
							<td><strong>Receipt No:</strong></td>
							<td><input type="text" id="paymentid" name="paymentid" value="<%=iPaymentId%>" size="15" maxlength="15" /></td>
						</tr>
						<tr>
							<td><strong>Name Like:</strong></td>
							<td><input type="text" name="searchname" value="<%=sSearchName%>" size="75" maxlength="255" /></td>
						</tr>
						<tr>
							<td><strong>Drop Date:</strong></td>
							<td>
								<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" placeholder=">= date" />
								<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>		 
								&nbsp;
								<strong>To:</strong>
								<input type="text" id="toDate" name="toDate" value="<%=toDateDisplay%>" size="10" maxlength="10" placeholder="<= date" />
								<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
								&nbsp; <%DrawDateChoices "Dates" %>
							</td>
						</tr>
						<tr>
							<td><strong>Order By:</strong></td>
							<td>
								<% DisplayOrderByPicks iOrderBy %>
							</td>
						</tr>
					</table>
					
					<!--END: FILTERS-->

					<p>
						<input class="button" type="button" value="View Report" onClick="submitForm('view')" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="submitForm('excel')" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				
      
			</td>
		 </tr>
	</table>
  </form>
	</div>
</div>

<!--BEGIN: DISPLAY RESULTS-->
<%
If request.servervariables("REQUEST_METHOD") = "POST" Then 
	Display_Results sWhereClause, sFrom, sOrderBy
Else
	response.write "<div id=""refundnotesblock""><strong>To view the Refund Notes Report, select from the filter options above then click the &quot;View Report&quot; button.</strong></div>"
End If 

%>
<!-- END: DISPLAY RESULTS -->

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

	
<%
'------------------------------------------------------------------------------------------------------------
' void DrawDateChoices sName 
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(document.frmPFilter." & sName & ".value, '" + sName + "');"" id=""" & sName & """ class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 

'------------------------------------------------------------------------------------------------------------
' void Display_Results sWhereClause, sFrom 
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sWhereClause, ByVal sFrom, ByVal sOrderBy )
	Dim sSql, oRs, bgcolor

	sSql = "SELECT P.paymentid, P.paymentdate, P.userid, ISNULL(U.userfname,'' ) AS userfname, ISNULL(U.userlname,'') AS userlname, "
	sSql = sSql & "ISNULL(U.useraddress,'') AS useraddress, U.usercity, U.userstate, U.userzip, ISNULL(U.userhomephone,'') AS userhomephone, "
	sSql = sSql & "C.classname, C.classseasonid, ISNULL(T.activityno,'') AS activityno, D.dropreason, ISNULL(P.notes,'') AS notes "
	sSql = sSql & "FROM egov_class_payment P, egov_users U, egov_class_list L, egov_class C, egov_class_time T, egov_class_dropreasons D " & sFrom
	sSql = sSql & " WHERE P.userid = U.userid AND P.paymentid = L.paymentid	AND L.classid = C.classid "
	sSql = sSql & "AND L.classtimeid = T.timeid AND P.dropreasonid = D.dropreasonid AND P.journalentrytypeid = 2 "
	sSql = sSql & "AND P.orgid = " & Session("orgid") & sWhereClause
	sSql = sSql & " ORDER BY " & sOrderBy

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.Write vbcrlf & "<div id=""refundnotesblock"">"
	'response.write sSql & "<br><br>"
	If oRs.EOF then
		' EMPTY
		response.write "<strong>No Refunds were found that match your selection criteria.</strong>"
	Else
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" id=""refund_notes"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Name</th><th>Address</th><th>Phone</th>"
		response.write "<th>Class Name</th><th>Class #</th><th>Receipt #</th><th>Reason</th><th>Notes</th></tr>"

		bgcolor = "#eeeeee"

		Do While Not oRs.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			response.write "<td nowrap>" & Trim( oRs("userfname") & " " & oRs("userlname") ) & "</td>"
			response.write "<td nowrap>" & oRs("useraddress") & "</td>"
			response.write "<td nowrap>" & FormatPhoneNumber(oRs("userhomephone")) & "</td>"
			response.write "<td>" & oRs("classname") & "</td>"
			response.write "<td align=""center"">" & oRs("activityno") & "</td>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") & """>" & oRs("paymentid") & "</a></td>"
			response.write "<td nowrap>" & oRs("dropreason") & "</td>"
			response.write "<td>" & oRs("notes") & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</table>"
		
	End If

	response.write vbcrlf & "</div>"

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' DisplayOrderByPicks iOrderBy
'------------------------------------------------------------------------------------------------------------
Sub DisplayOrderByPicks( ByVal iOrderBy )

	response.write vbcrlf & "<select name=""orderby"">"
	response.write vbcrlf & "<option value=""0"""
	If clng(iOrderBy) = clng(0) Then
		response.write " selected=""selected"""
	End If
	response.write ">Date</option>"
	response.write vbcrlf & "<option value=""1"""
	If clng(iOrderBy) = clng(1) Then
		response.write " selected=""selected"""
	End If
	response.write ">Name</option>"
	response.write vbcrlf & "</select>"

End Sub


%>	

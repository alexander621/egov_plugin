<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

' GET USER'S PERMISSIONS
blnCanViewAllActionItems = HasPermission("CanViewAllActionItems")
blnCanViewOwnActionItems = HasPermission("CanViewOwnActionItems")
blnCanViewDeptActionItems = HasPermission("CanViewDeptActionItems")


' IF USER HAD SET FILTERS FOR THIS SESSION THEN REMEMBER THEM
If request("useSessions") = 1 then
	' USE/SET FILTERS SET FOR THE SESSION
		recordsPer = session("recordsPer")
		reportType = session("reportType")
			
		orderBy = session("orderBy")
		selectFormId = session("selectFormId")
		selectAssignedto = session("selectAssignedto")
		selectDeptId = session("selectDeptId")
			
		selectUserFName = session("selectUserFName")
		selectUserLName = session("selectUserLName")
			
		fromDate = session("fromDate")
		toDate = session("toDate")
		today = Date()
			
		statusSubmitted = session("statusSubmitted")
		statusInprogress = session("statusInprogress")
		statusWaiting = session("statusWaiting")
		statusResolved = session("statusResolved")
		statusDismissed = session("statusDismissed")
Else
	' USE/SET DEFAULT FILTERS
		recordsPer = request("recordsPer")
		reportType = request("reportType")
			
		orderBy = request("orderBy")
		selectFormId = request("selectFormId")
			
		If (NOT blnCanViewAllActionItems) AND (NOT blnCanViewDeptActionItems) AND blnCanViewOwnActionItems Then
			selectAssignedto = Session("FullName")
		Else
			selectAssignedto = request("selectAssignedto")
		End If 
			
		selectDeptId = request("selectDeptId")
			
		selectUserFName = request("selectUserFName")
		selectUserLName = request("selectUserLName")
			
		fromDate = Request("fromDate")
		toDate = Request("toDate")
		today = Date()
			
		statusSubmitted = request("statusSubmitted")
		statusInprogress = request("statusInprogress")
		statusWaiting = request("statusWaiting")
		statusResolved = request("statusResolved")
		statusDismissed = request("statusDismissed")
End If


' SET REPORT TYPE (LIST,SUMMARY, OR DETAIL) FILTER
If reportType = "" or IsNull(reportType) Then reportType = "List" End If

' SET ORDER BY COLUMN FILTER
If orderBy = "" or IsNull(orderBy) Then 
	' USE SELECT ORDER BY COLUMN
	orderBy = request("groupBy")

	' CHECK TO SEE IF ORDER BY HAS VALUE IF NOT DEFAULT TO SUBMIT_DATE
	If orderBy = "" or IsNull(orderBy) Then
		orderBy = "submit_Date" 
	End if
End If

' SET FORMID FILTER
If selectFormId = "" or IsNull(selectFormId) Then selectFormId = "all" End If

' SET EMPLOYEE ASSIGNED FILTER
If selectAssignedto = "" or IsNull(selectAssignedto) Then selectAssignedto = "all" End If

' SET DEPARTMENT FILTER
If selectDeptId = "" or IsNull(selectDeptId) Then selectDeptId = "all" End If

' SET USER FIRST NAME FILTER
If selectUserFName = "" or IsNull(selectUserFName) Then selectUserFName = "all" End If

' SET LAST NAME FILTER
If selectUserLName = "" or IsNull(selectUserLName) Then selectUserLName = "all" End If

' SET TODATE FILTER
If toDate = "" or IsNull(toDate) Then toDate = dateAdd("d",0,today) End If
toDate = dateAdd("d",1,toDate)

' SET FROMDATE FILTER
If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("yyyy",-1,today) End If

' SET RECORDS PER PAGE FILTER
if recordsPer = "" or IsNull(recordsPer) Or clng(recordsPer) = 0 Then recordsPer = 25 End If

' SET FILTER STATUS
noStatus = true
If statusSubmitted = "yes" Then 
	noStatus = false
ELSE
	statusSubmitted = "no"
End If
If statusInprogress = "yes" Then 
	noStatus = false
ELSE
	statusInprogress = "no"
End If
If statusWaiting = "yes" Then 
	noStatus = false
ELSE
	statusWaiting = "no"
End If
If statusResolved = "yes" Then 
	noStatus = false
ELSE
	statusResolved = "no"
End If
If statusDismissed = "yes" Then 
	noStatus = false
ELSE
	statusDismissed = "no"
End If
if noStatus = true then
	statusDismissed = "yes"
	statusResolved = "yes"
	statusWaiting = "yes"
	statusInprogress = "yes"
	statusSubmitted = "yes"
end if

' SET SESSION VARIABLES FOR REMEMBERING DURING THIS SESSION
Session("reportType") = reportType
Session("orderBy") = orderBy
Session("selectFormId") =  selectFormId
Session("selectAssignedto") = selectAssignedto 
Session("selectDeptId") = selectDeptId

Session("toDate") = toDate 
Session("fromDate") = fromDate
Session("recordsPer") = recordsPer

Session("noStatus") = noStatus
Session("statusDismissed") = statusDismissed
Session("statusResolved") = statusResolved
Session("statusWaiting") = statusWaiting
Session("statusInprogress") =  statusInprogress
Session("statusSubmitted") =  statusSubmitted

Session("selectUserFName") =  selectUserFName
Session("selectUserLName") =  selectUserLName
%>


<html>
<head>
  <title><%=langBSActionLine%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  
  <SCRIPT LANGUAGE="JavaScript">
  
  function checkStat() {
  if ( !(form1.statusSubmitted.checked) &&  !(form1.statusInprogress.checked) && !(form1.statusWaiting.checked) && !(form1.statusResolved.checked) && !(form1.statusDismissed.checked)) {
		alert("You must select the status.");

		form1.statusSubmitted.focus();
		return false;
	}
  }
 
  function CheckAllStatus(checkSt) {
		//if (document.form1.CheckAllStat.checked) {
		if (checkSt) {
			document.form1.statusSubmitted.checked = true;
			document.form1.statusInprogress.checked = true;
			document.form1.statusWaiting.checked = true;
			document.form1.statusResolved.checked = true;
			document.form1.statusDismissed.checked = true;
		} else {
			document.form1.statusSubmitted.checked = false;
			document.form1.statusInprogress.checked = false;
			document.form1.statusWaiting.checked = false;
			document.form1.statusResolved.checked = false;
			document.form1.statusDismissed.checked = false;
		}
  }
 
 function submitForm(){
		 if (document.form1.reportType.value == "Summary") {
				document.forms[0].action = "action_line_summary.asp"
				document.forms[0].submit();
			} else {
				document.forms[0].action = "action_line_list.asp"
				document.forms[0].submit();
			}
 }
 </SCRIPT>
 
 <script language="Javascript">
  <!--
    function doCalendar(ToFrom) {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }
  //-->
  </script>


</head>


<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">


<%DrawTabs tabActionline,1%>


  <table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
    <tr>
      <td><font size="+1"><b>(E-Gov Request Manager) - Manage Action Line Requests</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToStart%></a></td>
    </tr>

    <tr>
	  <td>
		  
		  <!--BEGIN: FILTER SELECTION-->
		  <fieldset>
			  <legend><b>Search/Sorting Option(s)</b></legend>
			  <form name=form1  onSubmit="return checkStat()">
		  		  
		  
		  <table border=0 bordercolor=red>
		  <tr>
		  <!--ASSIGNED USER FILTER-->
		  <td valign=top nowrap>
			 <%
			 ' DISPLAY ASSIGNED TO SELECTION
			 If blnCanViewAllActionItems Then		  
				response.write "<b>Assigned To: "
				If NOT blnCanViewAllActionItems Then 
					' DISPLAY CURRENTLY LOGGED IN ADMINISTRATOR
					response.write "(User " & session("userID") & ")"
				End If 
				' DRAW LIST OF EMPLOYEES
				DrawAssignedEmployeeSelection(session("orgid"))
			 End If
			 %>		  
			<!--ORDER BY FILTER-->
			 <b>Order By: 
			<select name="orderBy">
				<% if orderBy = "submit_Date" then select1 = "selected" else select1="" %>
						<option value="submit_Date" <%=select1%>>Date Descending</option>
				<% if orderBy = "action_Formid" then select1 = "selected" else select1="" %>
						<option value="action_Formid" <%=select1%>>Category</option>
				<% if orderBy = "deptId" then select1 = "selected" else select1="" %>
						<option value="deptId" <%=select1%>>Department</option>
				<% if orderBy = "assigned_Name" then select1 = "selected" else select1="" %>
						<option value="assigned_Name" <%=select1%>>Assigned To</option>			
 		    </select>
		  </td>

        </tr>
		<tr>
			<td valign=top nowrap>
			<!--STATUS FILTER-->
			<%
			If statusSubmitted = "yes" then check1 = "checked"
			If statusInprogress = "yes" Then check2 = "checked" 
			If statusWaiting = "yes" Then check3 = "checked"
			If statusResolved= "yes" Then check4 = "checked"
			If statusDismissed = "yes" Then check5 = "checked"
			%>
			
			<b>Status:</b> 
			  <input type=checkbox name="statusSubmitted" value="yes" <%=check1%>>Submitted
			 <input type=checkbox name="statusInprogress" value="yes" <%=check2%>>In Progress
			 <input type=checkbox name="statusWaiting" value="yes" <%=check3%>>Waiting
			 <input type=checkbox name="statusResolved" value="yes" <%=check4%>>Resolved
			 <input type=checkbox name="statusDismissed" value="yes" <%=check5%>>Dismissed
			</td></tr>		  
		  
		  <tr>
		  <td valign=top nowrap>
		  <!--CATEGORY FILTER-->
		    <b>Category: 
		    <select name="selectFormId"><option value="">All Categories</option><% fnListForms()%></select>
			</td>
      </tr>
      
      <tr>
		  <td valign=top nowrap>
		  <!--DEPARTMENT FILTER-->
			<% 
			If blnCanViewAllActionItems OR blnCanViewDeptActionItems  Then 
				response.write "<b>Department: </b> "
				response.write "<select name=""selectDeptId""><option value=""all"">All Departments</option>"
				'	GET A LIST OF ALL AVAILABLE DEPARTMENTS FOR THIS USER
				fnListDepts selectDeptId 
				response.write "</select>&nbsp;&nbsp;&nbsp;"
		    End If
			%>
		    

		 <!--REPORT TYPE FILTER-->		 
		    <b>Report Type: 
			    <select name="reportType">
			  <% if reportType = "Detail" or reportType = "DrillThru" then %>
						<option value="List">List</option>
						<option value="Summary">Summary</option>
						<option value="Detail" selected>Detail</option>
			   <% elseif reportType = "Summary" then %>
						<option value="List">List</option>
						<option value="Summary" selected>Summary</option>
						<option value="Detail">Detail</option>
			   <% else %>
						<option value="List">List</option>
						<option value="Summary">Summary</option>
						<option value="Detail">Detail</option>
			  <% end if %>
			  </select>
			</td>
      </tr>
      
      <tr>
		  <td valign=top nowrap>
		  <!--DATE RANGE FILTER-->
      <b>From: 
					  <input type=text name="fromDate" value="<%=fromDate%>">
					  <a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border=0></a>		 
				   &nbsp; 
					<b>To:</b> 
					  <input type=text name="toDate" value="<%=dateAdd("d",-1,toDate)%>">
					  <a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border=0></a>
					  
				</td>
				</tr>	  
      
      
		  <tr>	
		  <!--SUBMITTED BY FILTER-->

	<td valign=top nowrap>
      <b>Submitted By: &nbsp;&nbsp;
					  First: <input id="subfirstname" type=text name="selectUserFName" value="<% if selectUserFName <> "all" then response.write selectUserFName %>" size=12>
				   &nbsp; 
					<b>Last:</b> 
					  <input id="sublastname" type=text name="selectUserLName" value="<% if selectUserLName <> "all" then response.write selectUserLName %>" size=12>
					  
				</td>
      </tr>	  

      
		  <tr>	
		  <!--RECORDS PER PAGE FILTER-->
		  <td valign=top>
			   <input type=button onclick="javascript:submitForm();" value=" SEARCH ">
			   
			    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			    
			    <b>Records per Page: 
						<input type=text name="recordsPer" value="<%=recordsPer%>" size=2>
		  </td>
		    </tr>
			</table>
			</form>
			</fieldset>
			<!--END: FILTER SELECTION-->


			 </td>
    </tr>


    <tr>
      <td valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
      <form name=requestlist action=# method="POST">
		<% List_Action_Requests(sSortBy) %>
	  </form>
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>

</body>
</html>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' FUNCTION LIST_ACTION_REQUESTS(SSORTBY)
'------------------------------------------------------------------------------------------------------------
Function List_Action_Requests(sSortBy)


Dim statArray(5)
	i = 0
If statusSubmitted = "yes" then 
	statArray(i) = " status='SUBMITTED' OR"
	i = i + 1
End If
If statusInprogress = "yes" Then 
	statArray(i) = " status='INPROGRESS' OR"
	i = i + 1
End If
If statusWaiting = "yes" Then
	statArray(i) = " status='WAITING' OR"
	i = i + 1
End If
If statusResolved= "yes" Then
	statArray(i) = " status='RESOLVED' OR"
	i = i + 1
End If
If statusDismissed = "yes" Then 
	statArray(i) = " status='DISMISSED' OR"
	i = i + 1
End If

for u = 0 to ubound(statArray)
	varStatClause = varStatClause & "" & statArray(u)
next
lenStatClause = len(varStatClause) - 3
if lenStatClause > 1 then
	varStatClause = left(varStatClause,lenStatClause)
end if



'If start date is before from date AND finish date is NOT before from date
'OR
'If start date is NOT before from date AND start date is NOT after to date

varWhereClause = " WHERE  (egov_action_request_view.orgid=" & session("orgid") & ") AND ( "    ''IsNull(complete_date,'" & Now & "')
		varWhereClause = varWhereClause & " (submit_Date >= '" & fromDate & "' AND submit_Date < '" & toDate & "') OR "
		varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
		varWhereClause = varWhereClause & " (submit_Date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "
varWhereClause = varWhereClause & " ) AND (" & varStatClause & ") "

''**varWhereClause = " WHERE (submit_Date >= '" & fromDate & "' AND submit_Date < '" & toDate & "') AND (" & varStatClause & ")"

If selectFormId <> "all" then 
	If left(selectFormId,1)="C" then 
		sSQLb = "SELECT action_form_id FROM egov_forms_to_categories where form_category_id = " & right(selectFormId,len(selectFormId)-1)
		Set oCategories = Server.CreateObject("ADODB.Recordset")
		
		oCategories.Open sSQLb, Application("DSN"), 3, 1
		
		if oCategories.EOF then
			varWhereClause = varWhereClause & " AND form_category_id=999999"		
		else
			do while not oCategories.EOF
					CatArray = CatArray & oCategories("action_form_id") & ","
			oCategories.MoveNext
			Loop
			CatArray = left(CatArray,(len(CatArray)-1))
			
		end if
		varWhereClause = varWhereClause & " AND action_Formid IN (" & CatArray & ") "
	else
		varWhereClause = varWhereClause & " AND action_Formid = " & selectFormId
	end if
end if

If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '" & selectAssignedto & "'" 

If blnCanViewDeptActionItems AND NOT blnCanViewAllActionItems Then
	If selectDeptId <> "all" then 
		varWhereClause = varWhereClause & " AND deptID = '" & selectDeptId & "'" 
	Else
		varWhereClause = varWhereClause & " AND deptID IN (" & GetGroups(session("user_id")) & ") OR (assignedemployeeid = '" & session("userid") & "') " 
	End If
Else
	If selectDeptId <> "all" then 
		varWhereClause = varWhereClause & " AND deptID = '" & selectDeptId & "'" 
	End If
End If


If selectUserFName <> "all" then varWhereClause = varWhereClause & " AND UserFName LIKE '" & selectUserFName & "%'"
If selectUserLName <> "all" then varWhereClause = varWhereClause & " AND UserLName LIKE '" & selectUserLName & "%'"

	sSQL = "SELECT userlname,userfname,action_autoid,action_formTitle,DateDiff(d,submit_date,complete_date) AS totalDays,submit_date,complete_date,deptID,groupname as deptName,status,assignedName as assigned_Name,assignedemployeeid FROM egov_action_request_view left outer join groups on deptId=groupId" & varWhereClause & " AND (egov_action_request_view.orgid=" & session("orgid") & ") ORDER BY " & orderBy
	if orderBy = "submit_Date" then
		sSQL = sSQL & " desc"
	end if

' SET GOOGLE MAP QUERY
session("MAP_QUERY") = sSQL

Set oRequests = Server.CreateObject("ADODB.Recordset")


	 ' SET PAGE SIZE AND RECORDSET PARAMETERS
	 oRequests.PageSize = recordsPer
	 oRequests.CacheSize = recordsPer
	 oRequests.CursorLocation = 3
	 
	 ' OPEN RECORDSET
	 oRequests.Open sSQL, Application("DSN"), 3, 1

lastTitle = "Test"
lastDate = "1/1/02"
lastDept = 11798
lastDeptName = "Test"
lastAssigned = "bubba"
displayLastTitle = "Test"
lastSubmitted = "bubba"

	 
if oRequests.EOF=false then
	 ' SET PAGE TO VIEW
	 if request("useSessions")=1 then
	 							If Len(Session("pagenum")) <> 0 then
										oRequests.AbsolutePage = clng(Session("pagenum"))	
										Session("pageNum") = clng(Session("pagenum"))		
								Else
										oRequests.AbsolutePage = 1
										Session("pageNum") = 1
								End If
								'Response.write "Issue 1"
	 else
						 If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
								oRequests.AbsolutePage = 1
								Session("pageNum") = 1
								'Response.write "Issue 2"
						 Else
								If clng(Request("pagenum")) <= oRequests.PageCount Then
										oRequests.AbsolutePage = Request("pagenum")
										Session("pageNum") = Request("pagenum")
										'Response.write "Issue 3: " & oRequests.PageCount & "-" & clng(Request("pagenum"))
								Else
										oRequests.AbsolutePage = 1
										Session("pageNum") = 1
										'Response.write "Issue 4"
								End If
								
						 End If
	 end if
	 
	 
	 ' DISPLAY RECORD STATISTICS
	  Dim abspage, pagecnt
		abspage = oRequests.AbsolutePage
		pagecnt = oRequests.PageCount
	  
	  If request("selectAssignedto") <> "" Then
		 sQueryString = replace(request.querystring,"pagenum","HFe301") ' REPLACE PAGENUM FIELD WITH RANDOM FIELD FOR NAVIGATION PURPOSES
	  Else
		 sQueryString = "filter=false"
	  End If
	  
	  
	  Response.Write "<Br><font size=3 color=3399ff><i><b>" & Ucase(Replace(ReportType,"DrillThru","Drill Through")) & " REPORT</b></i></font>"
	  
	  
	  ' ADD MAP LINK
	  response.write " - <a href=""../maps/beta_map.asp"" target=""_blank"">Map It!</a>"

	 
	  Response.Write "<Br><b>Page <font color=blue>" & oRequests.AbsolutePage & "</font>  " & vbcrlf
	  Response.Write "of <font color=blue> " & oRequests.PageCount & "</font></b> &nbsp;|&nbsp; " & vbcrlf
	  Response.Write "<b><font color=blue>" & oRequests.RecordCount & "</font> total Action Item Requests</b>"

	 ' DISPLAY FORWARD AND BACKWARD NAVIGATION TOP AND PRINT PAGE LINK
 	  Response.write "<div><table width=""100%""><tr><td valign=top><table><tr><td><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></td><td width=450 align=right><a href=""action_line_list_print.asp?orderBy=" & orderBy & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName &"&toDate=" & toDate & "&fromDate=" & fromDate  & "&reportType=" & reportType & """ target=new>Open New Printer Friendly Results Window</a></td></tr></table></div>"
	  
	  Response.Write "<table cellspacing=0 cellpadding=5 class=tablelist width=""100%"">"
	  Response.Write "<tr class=tablelist>"
	  
	      ' CHANGE TO BCANEDIT LATER
	      If 1=1 Then 
			' Response.Write "<th><input class=""listCheck"" type=checkbox name=""chkSelectAll"" onClick=""selectAll('requestlist', this.checked)""></th>"
	      Else
	            ' Response.Write "<th>&nbsp;</th>"
	      End If
	
	  Response.Write "<th>Action Line Category</th><th>Date submitted</th>"
	  if ReportType="Detail" or ReportType="DrillThru" then
	  	Response.Write "<th>Date Completed</th><th>Days open*/To complete</th>"
	  end if
	  Response.Write "<th>Status</th><th>Submitted by</th><th>Assigned to</th><th>Department</th></tr>"



 '/////////////////
'sSQLtotl = "SELECT userlname,userfname,action_autoid,action_formTitle,DateDiff(d,submit_date,complete_date) AS totalDays,submit_date,complete_date,deptID,groupname as deptName,status,assigned_Name FROM egov_action_request_view left outer join groups on deptId=groupId" & varWhereClause
sSQLtotl = "SELECT action_autoid,action_formTitle,DateDiff(d,submit_date,complete_date) AS totalDays,submit_date,complete_date,deptID,groupname as deptName,status,assigned_Name FROM egov_action_request_view left outer join groups on deptId=groupId" & varWhereClause
Set oTotals = Server.CreateObject("ADODB.Recordset")
oTotals.Open sSQLtotl, Application("DSN"), 3, 1
	 '/////////////////

sSQLTotal = "SELECT count(*) as numTotal FROM egov_action_request_view " & varWhereClause    '''////& " AND status<>'RESOLVED' AND status<>'DISMISSED'"
response.write "<!--JOHNCOMMENT" & sSQLTotal & "-->"
Set oTotal = Server.CreateObject("ADODB.Recordset")
oTotal.Open sSQLTotal, Application("DSN"), 3, 1

sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') "
Set oOpen = Server.CreateObject("ADODB.Recordset")
oOpen.Open sSQLopen, Application("DSN"), 3, 1			

sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED') "
Set oClosed = Server.CreateObject("ADODB.Recordset")
oClosed.Open sSQLclosed, Application("DSN"), 3, 1

'AVG OPEN	
numOpen = clng(oOpen("numOpen"))
if numOpen<>0 and oOpen("totalDaysOpen")<> 0 then
		avgOpenTotal = oOpen("totalDaysOpen") / numOpen
		avgOpenTotal = formatnumber(avgOpenTotal,1)
else
		avgOpenTotal = ""
end if	
'AVG CLOSED
numClosed = clng(oClosed("numClosed")) 
if numClosed<>0 and oClosed("totalDaysClosed")<> 0 then
		avgClosedTotal = oClosed("totalDaysClosed") / numClosed
		avgClosedTotal = formatnumber(avgClosedTotal,1)
else
		avgClosedTotal = ""
end if					
							'/////////////////



	  
	  if ReportType="Detail" or ReportType="DrillThru" then	
							Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b><font color=navy size=1>Grand Total [" & oRequests.RecordCount & " Requests]</td><td> <b><font color=navy size=1> Submitted: " & oTotal("numTotal") & "</td><td> <b><font color=navy size=1> Open: " & oOpen("numOpen") & "</td><td colspan=2> <b><font color=navy size=1> Avg Time Still Open: " & avgOpenTotal & "</td><td colspan=2> <b><font color=navy size=1> Avg Time To Complete: " & avgClosedTotal & "</td><td> </td></tr>"
		end if
	  
	  
	  '*** LOOP AND DISPLAY THE RECORDS
	  bgcolor = "#eeeeee"
		 For intRec=1 To oRequests.PageSize
		  If Not oRequests.EOF Then

			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If

		  ' GET VALUES
			If oRequests("action_formTitle") <> "" Then
				sTitle = oRequests("action_formTitle")
			Else
				sTitle = "<font color=red><b>???</b></font>"
			End If
			
			''If oRequests("submit_Date") <> "" Then
				sDate = oRequests("submit_date")
				sDate = formatDateTime(sDate,vbShortDate)
			''Else
			''	sDate = "<font color=red><b>???</b></font>"
			''End If
			
			If oRequests("assigned_Name") <> "" Then
				sAssigned = oRequests("assigned_Name")
			Else
				sAssigned = "<font color=red><b>???</b></font>"
			End If
			
			If oRequests("userlname") <> "" Then
				sSubmitted = oRequests("userfname") & " " & oRequests("userlname")
			Else
				sSubmitted = "<font color=red><b>???</b></font>"
			End If
			
			
			
			'If oRequests("deptId") <> "" Then
				sDept = oRequests("deptId")
				sDeptName= oRequests("deptName")
			'Else
			'	sDept = "<font color=red><b>???</b></font>"
			'End If
			

			If oRequests("status") <> "" Then
				sStatus = oRequests("status")
			Else
				sStatus = "<font color=red><b>???</b></font>"
			End If

			If oRequests("submit_date") <> "" Then
				datSubmitDate = oRequests("submit_date")
			Else
				datSubmitDate = "<font color=red><b>???</b></font>"
			End If
			
			If oRequests("complete_date") <> "" Then
				datResolveDate = oRequests("complete_date")
				datResolveDate = formatdatetime(datResolveDate,vbShortDate)
			Else
				datResolveDate = "<font color=red><b>???</b></font>"
			End If

			lngTrackingNumber = oRequests("action_autoid") & replace(FormatDateTime(oRequests("submit_date"),4),":","")
			
			
''//////////
'**INSERT SUBTOTAL ROW IF DETAIL REPORT**
			if ReportType="Detail" or ReportType="DrillThru" then
			
			
						if orderBy = "submit_Date" then
							if DateDiff("d",sDate,lastDate) = 0 then
									'NO NEW LINE
							else
									 
								if lastDate <> "1/1/02" then
									
									sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND submitdateshort = '" & lastDate & "'"
									SQLopenText = "" & sSQLopen & ""
									Set oOpenSub = Server.CreateObject("ADODB.Recordset")
									oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			
									
									sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND submitdateshort = '" & lastDate & "'"
									SQLclosedText = "" & sSQLclosed & ""
									Set oClosedSub = Server.CreateObject("ADODB.Recordset")
									oClosedSub.Open sSQLclosed, Application("DSN"), 3, 1
									
									'av OPEN	
									numOpen = clng(oOpenSub("numOpen"))
									if numOpen<>0  then
											avOpenTotal = oOpenSub("totalDaysOpen") / numOpen
											avOpenTotal = formatnumber(avOpenTotal,1)
									else
											avOpenTotal = " - "
									end if	
									'av CLOSED
									numClosed = clng(oClosedSub("numClosed")) 
									if numClosed<>0 then
											avClosedTotal = oClosedSub("totalDaysClosed") / numClosed
											avClosedTotal = formatnumber(avClosedTotal,1)
									else
											avClosedTotal = " - "
									end if	
									
									displayLastTitle = sDate
									displayLastAvgOpen = avOpenTotal
									displayLastAvgClosed =	avClosedTotal


										Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b><font color=navy size=1>Subtotal: " & lastDate & " </td><td> <b><font color=navy size=1> Submitted: " & subTotl & "</td><td> <b> <font color=navy size=1> Open: " & numOpen & "</td><td colspan=2><b><font color=navy size=1> Avg Time Still Open: " & avOpenTotal & "</td><td colspan=3><b><font color=navy size=1> Avg Time To Complete: " & avClosedTotal & "</td></tr>"
									end if
							end if
						elseif orderBy = "action_Formid" then
							if sTitle = lastTitle then
									'NO NEW LINE
							else
									
									if  lastTitle <> "Test" then
										sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND action_FormTitle='" & lastTitle & "'"
										SQLopenText = "" & sSQLopen & ""
										Set oOpenSub = Server.CreateObject("ADODB.Recordset")
										oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			
										
										sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND action_FormTitle='" & lastTitle & "'"
										SQLclosedText = "" & sSQLclosed & ""
										Set oClosedSub = Server.CreateObject("ADODB.Recordset")
										oClosedSub.Open sSQLclosed, Application("DSN"), 3, 1
										
										'av OPEN	
										numOpen = clng(oOpenSub("numOpen"))
										if numOpen<>0  then
												avOpenTotal = oOpenSub("totalDaysOpen") / numOpen
												avOpenTotal = formatnumber(avOpenTotal,1)
										else
												avOpenTotal = " - "
										end if	
										'av CLOSED
										numClosed = clng(oClosedSub("numClosed")) 
										if numClosed<>0 then
												avClosedTotal = oClosedSub("totalDaysClosed") / numClosed
												avClosedTotal = formatnumber(avClosedTotal,1)
										else
												avClosedTotal = " - "
										end if				
									
									displayLastTitle = sTitle
									displayLastAvgOpen = avOpenTotal
									displayLastAvgClosed =	avClosedTotal
										
									
										Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b> <font color=navy size=1>Subtotal: " & lastTitle & " </td><td> <b> <font color=navy size=1> Submitted: " & subTotl & "</td><td> <b> <font color=navy size=1> Open: " & numOpen & "</td><td colspan=2><b><font color=navy size=1> Avg Time Still Open: " & avOpenTotal & "</td><td colspan=3><b><font color=navy size=1> Avg Time To Complete: " & avClosedTotal & "</td></tr>"
									end if
							end if
						elseif orderBy = "deptId" then
							if sDept = lastDept then
									'NO NEW LINE
							else
								if  lastDept <> 11798 then
										
									sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND deptId = '" & lastDept & "'"
									SQLopenText = "" & sSQLopen & ""
									Set oOpenSub = Server.CreateObject("ADODB.Recordset")
									oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			
									
									sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND deptId = '" & lastDept & "'"
									SQLclosedText = "" & sSQLclosed & ""
									Set oClosedSub = Server.CreateObject("ADODB.Recordset")
									oClosedSub.Open sSQLclosed, Application("DSN"), 3, 1
									
									'av OPEN	
									numOpen = clng(oOpenSub("numOpen"))
									if numOpen<>0  then
											avOpenTotal = oOpenSub("totalDaysOpen") / numOpen
											avOpenTotal = formatnumber(avOpenTotal,1)
									else
											avOpenTotal = " - "
									end if	
									'av CLOSED
									numClosed = clng(oClosedSub("numClosed")) 
									if numClosed<>0 then
											avClosedTotal = oClosedSub("totalDaysClosed") / numClosed
											avClosedTotal = formatnumber(avClosedTotal,1)
									else
											avClosedTotal = " - "
									end if		
									
									displayLastTitle = sDeptName
									displayLastAvgOpen = avOpenTotal
									displayLastAvgClosed =	avClosedTotal
									
									Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b> <font color=navy size=1>Subtotal: " & lastDeptName & " </td><td> <b> <font color=navy size=1> Submitted: " & subTotl & "</td><td> <b> <font color=navy size=1> Open: " & numOpen & "</td><td colspan=2><b><font color=navy size=1> Avg Time Still Open: " & avOpenTotal & "</td><td colspan=3><b><font color=navy size=1> Avg Time To Complete: " & avClosedTotal & "</td></tr>"
									end if
							end if
						elseif orderBy = "assigned_Name" then
							if sAssigned = lastAssigned then
									'NO NEW LINE
							else
								  if lastAssigned <> "bubba" then						
										sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND assignedName = '" & lastAssigned & "'"
										SQLopenText = "" & sSQLopen & ""
										Set oOpenSub = Server.CreateObject("ADODB.Recordset")
										oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			
										
										sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND assignedName = '" & lastAssigned & "'"
										SQLclosedText = "" & sSQLclosed & ""
										Set oClosedSub = Server.CreateObject("ADODB.Recordset")
										oClosedSub.Open sSQLclosed, Application("DSN"), 3, 1
										
										'av OPEN	
										numOpen = clng(oOpenSub("numOpen"))
										if numOpen<>0  then
												avOpenTotal = oOpenSub("totalDaysOpen") / numOpen
												avOpenTotal = formatnumber(avOpenTotal,1)
										else
												avOpenTotal = " - "
										end if	
										'av CLOSED
										numClosed = clng(oClosedSub("numClosed")) 
										if numClosed<>0 then
												avClosedTotal = oClosedSub("totalDaysClosed") / numClosed
												avClosedTotal = formatnumber(avClosedTotal,1)
										else
												avClosedTotal = " - "
										end if
										
										displayLastTitle = sAssigned
										displayLastAvgOpen = avOpenTotal
										displayLastAvgClosed =	avClosedTotal		
									
									
										Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b> <font color=navy size=1>Subtotal: " & lastAssigned & " </td><td> <b> <font color=navy size=1> Submitted: " & subTotl & "</td><td> <b> <font color=navy size=1> Open: " & numOpen & "</td><td colspan=2><b><font color=navy size=1> Avg Time Still Open: " & avOpenTotal & "</td><td colspan=3><b><font color=navy size=1> Avg Time To Complete: " & avClosedTotal & "</td></tr>"
									end if
							  end if
						elseif orderBy = "userLname" then
							if sSubmitted = lastSubmitted then
									'NO NEW LINE
							else
								  if lastSubmitted <> "bubba" then						
										sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM gov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND userLname + ' ' + userFname='" & sSubmitted & "'"
										SQLopenText = "" & sSQLopen & ""
										Set oOpenSub = Server.CreateObject("ADODB.Recordset")
										oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			
										
										sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND userLname + ' ' + userFname='" & sSubmitted & "'"
										SQLclosedText = "" & sSQLclosed & ""
										Set oClosedSub = Server.CreateObject("ADODB.Recordset")
										oClosedSub.Open sSQLclosed, Application("DSN"), 3, 1
										
										'av OPEN	
										numOpen = clng(oOpenSub("numOpen"))
										if numOpen<>0  then
												avOpenTotal = oOpenSub("totalDaysOpen") / numOpen
												avOpenTotal = formatnumber(avOpenTotal,1)
										else
												avOpenTotal = " - "
										end if	
										'av CLOSED
										numClosed = clng(oClosedSub("numClosed")) 
										if numClosed<>0 then
												avClosedTotal = oClosedSub("totalDaysClosed") / numClosed
												avClosedTotal = formatnumber(avClosedTotal,1)
										else
												avClosedTotal = " - "
										end if
										
										displayLastTitle = sAssigned
										displayLastAvgOpen = avOpenTotal
										displayLastAvgClosed =	avClosedTotal		
									response.write "<br>" & displayLastTitle
									
										Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b> <font color=navy size=1>Subtotal: " & lastSubmitted & " </td><td> <b> <font color=navy size=1> Submitted: " & subTotl & "</td><td> <b> <font color=navy size=1> Open: " & numOpen & "</td><td colspan=2><b><font color=navy size=1> Avg Time Still Open: " & avOpenTotal & "</td><td colspan=3><b><font color=navy size=1> Avg Time To Complete: " & avClosedTotal & "</td></tr>"
									end if
							  end if
						end if
					end if		
					
			
			Response.Write "<tr bgcolor=" & bgcolor & "  onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"" onClick=""location.href='action_respond.asp?control=" & oRequests("action_autoid") & "';""><!--<td><input type=checkbox name=""del_" & oRequests("action_autoid") & """ >--><td><b>(" & lngTrackingNumber & ") " & sTitle & " </b></td><td align=center> " & formatdatetime(datSubmitDate,vbShortDate) & "</td>"
			
			if ReportType="Detail" or ReportType="DrillThru" then
					Response.Write "<td align=center>" & datResolveDate & "</td><td align=center>"
					
					if oRequests("totalDays")<> "" then
						 Response.Write oRequests("totalDays") & " days</td>"
						 countDays = clng(oRequests("totalDays"))
					else
						 openDays = dateDiff("d",oRequests("submit_Date"),Now)
						 Response.Write openDays & " days<font color=red>* </td>"
						 countDays = clng(openDays)
					end if
			end if
			Response.Write "<td align=center> " & UCASE(sStatus) & "</td><td align=center> " & oRequests("userfname")  & " " & oRequests("userlname") & "</td><td align=right>"
			Response.Write oRequests("assigned_Name") & "</td><td align=center> " & oRequests("deptName") & "</td></tr>"
			
			
			if ReportType="Detail" or ReportType="DrillThru" then
						if orderBy = "submit_Date" then
							  	if DateDiff("d",sDate,lastDate) = 0 and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
										subTotl = subTotl + 1
										openTotl = openTotl + 1
										totalDays = totalDays + countDays
								elseif DateDiff("d",sDate,lastDate) = 0 and UCASE(sStatus) <> "RESOLVED" then
										subTotl = subTotl + 1
										totalDays = totalDays + countDays											
								else
											if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
													openTotl = 1
													subTotl = 1
													totalDays = countDays
											else
													openTotl = 0
													subTotl = 1
													totalDays = countDays
											end if
								  end if
								  
								lastDate = sDate
						elseif orderBy = "action_Formid" then
								if sTitle = lastTitle and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
									subTotl = subTotl + 1
									openTotl = openTotl + 1
									totalDays = totalDays + countDays
									
								elseif sTitle = lastTitle then
									subTotl = subTotl + 1
									totalDays = totalDays + countDays
								else
											if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
													openTotl = 1
													subTotl = 1
													totalDays = countDays
											else
													openTotl = 0
													subTotl = 1
													totalDays = countDays
											end if
								end if
								lastTitle = sTitle
						elseif orderBy = "deptId" then
								if sDept = lastDept and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
										subTotl = subTotl + 1
										openTotl = openTotl + 1
										totalDays = totalDays + countDays
								elseif sDept = lastDept then
										subTotl = subTotl + 1
										totalDays = totalDays + countDays
									else
											if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
													openTotl = 1
													subTotl = 1
													totalDays = countDays
											else
													openTotl = 0
													subTotl = 1
													totalDays = countDays
											end if
									end if
									lastDept = sDept
									lastDeptName = sDeptName
						elseif orderBy = "assigned_Name" then
								if sAssigned  = lastAssigned and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
									subTotl = subTotl + 1
									openTotl = openTotl + 1
									totalDays = totalDays + countDays
								elseif sAssigned  = lastAssigned then
									subTotl = subTotl + 1
									totalDays = totalDays + countDays
								else
											if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
													openTotl = 1
													subTotl = 1
													totalDays = countDays
											else
													openTotl = 0
													subTotl = 1
													totalDays = countDays
											end if
								end if
								lastAssigned  = sAssigned 
					elseif orderBy = "userLname" then
							if sSubmitted  = lastSubmitted and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
									subTotl = subTotl + 1
									openTotl = openTotl + 1
									totalDays = totalDays + countDays
								elseif sSubmitted  = lastSubmitted then
									subTotl = subTotl + 1
									totalDays = totalDays + countDays
								else
											if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
													openTotl = 1
													subTotl = 1
													totalDays = countDays
											else
													openTotl = 0
													subTotl = 1
													totalDays = countDays
											end if
								end if
								lastSubmitted  = sSubmitted 

						end if
			end if
			
			oRequests.MoveNext 

		  End If
		 Next
		 
		if ReportType="Detail" or ReportType="DrillThru" then
				if orderBy = "submit_Date" then
					sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND submitdateshort = '" & lastDate & "'"
					sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND submitdateshort = '" & lastDate & "'"				
				elseif orderBy = "action_Formid" then
					sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND action_FormTitle='" & lastTitle & "'"
					sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND action_FormTitle='" & lastTitle & "'"
				elseif orderBy = "deptId" then
					sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND deptId = '" & lastDept & "'"
					sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND deptId = '" & lastDept & "'"
				elseif orderBy = "assigned_Name" then
					sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND assignedName = '" & lastAssigned & "'"
					sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND assignedName = '" & lastAssigned & "'"
				end if
				
				
				Set oOpenSub = Server.CreateObject("ADODB.Recordset")
				oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			
				
				Set oClosedSub = Server.CreateObject("ADODB.Recordset")
				oClosedSub.Open sSQLclosed, Application("DSN"), 3, 1
				
				'av OPEN	
				numOpen = clng(oOpenSub("numOpen"))
				if numOpen<>0  then
						avOpenTotal = oOpenSub("totalDaysOpen") / numOpen
						avOpenTotal = formatnumber(avOpenTotal,1)
				else
						avOpenTotal = " - "
				end if	
				'av CLOSED
				numClosed = clng(oClosedSub("numClosed"))
				
				if numClosed<>0 then
						avClosedTotal = oClosedSub("totalDaysClosed") / numClosed
						avClosedTotal = formatnumber(avClosedTotal,1)
				else
						avClosedTotal = " - "
				end if	
				
				
				Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b> <font color=navy size=1>Subtotal: " & displayLastTitle & " </td><td> <b> <font color=navy size=1> Submitted: " & subTotl & "</td><td> <b> <font color=navy size=1> Open: " & openTotl & "</td><td colspan=2><b><font color=navy size=1> Avg Time Still Open: " & avOpenTotal & "</td><td colspan=3><b><font color=navy size=1> Avg Time To Complete: " & avClosedTotal & "</td></tr>"
		End If
	
	
		 Response.Write "</table>"

	' DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM
	  'Response.write "<div><table><tr><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"

	  Response.write "<div><table border=0><tr><td valign=top><!--<img src=""../images/small_delete.gif"" align=""absmiddle"">&nbsp;<a href=""javascript:document.all.DelEvent.submit();"">DELETE</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;--><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 align=""absmiddle"" hspace=3 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"
else
	Response.write "<p><b>No records found</p>"
end if

End Function


'--------------------------------------------------------------------------------------------------
'  FUNCTION GETGROUPS(IUSERID)
'--------------------------------------------------------------------------------------------------
Function GetGroups(iUserID)

	sSQL = "SELECT  Groups.GroupID, Users.OrgID, Groups.GroupName, Groups.GroupDescription FROM Users INNER JOIN UsersGroups ON Users.UserID = UsersGroups.UserID INNER JOIN  Groups ON UsersGroups.GroupID = Groups.GroupID WHERE     (Groups.GroupType = 2) AND (Users.OrgID = " & Session("OrgID") & ") AND (Users.UserID = " & Session("UserID") & ") ORDER BY Groups.GroupName"

		Set oDepts = Server.CreateObject("ADODB.Recordset")
		oDepts.Open sSQL, Application("DSN") , 3, 1

	If NOT oDepts.EOF Then

		do while not oDepts.EOF
			sReturnValue = sReturnValue & "'" & oDepts("GroupID") & "',"
			oDepts.MoveNext
		Loop

		sReturnValue = LEFT(sReturnValue,LEN(sReturnValue)-1) 

	End If 
			
	Set oDepts = Nothing

	GetGroups = sReturnValue

End Function


Function fnListForms()
	sLastCategory = "NONE_START"
	sSQL = "SELECT * FROM dbo.egov_FormList WHERE orgid=" &session("orgid") & " order by form_category_Sequence,action_form_name"

	Set oForms = Server.CreateObject("ADODB.Recordset")
	oForms.Open sSQL, Application("DSN") , 3, 1

	If NOT oForms.EOF Then
		
		Do while NOT oForms.EOF 

				sCurrentCategory = oForms("form_category_name")
				If sLastCategory = "NONE_START" Then
					if selectFormId = "C" & oForms("form_category_id") & "" then
						selectA = "selected"
					else
						selectA = ""
					end if
					
					response.write "<option value=C" & oForms("form_category_id") & " " & selectA & ">----Category: " & sCurrentCategory & "</option>"
				End If
	
				if selectFormId = "C" & oForms("form_category_id") & "" then
					selectA = "selected"
				else
					selectA = ""
				end if
				
				If (sCurrentCategory <> sLastCategory) AND (sLastCategory <> "NONE_START") Then
					response.write "<option value=C" & oForms("form_category_id") & " " & selectA & ">----Category: " & sCurrentCategory &  "</option>"
				End If
				
				
				if cStr(selectFormId)=cStr(oForms("action_form_id")) then 
					selectA = "selected"
				else
					selectA = ""
				end if
				response.write "<option value=" & oForms("action_form_id") & " " & selectA & ">" & oForms("action_form_name") &  "</option>" & vbcrlf
			
			oForms.MoveNext
			sLastCategory = sCurrentCategory
		Loop

	End If

	Set oForms = Nothing
	End Function
	

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNLISTDEPTS(ISELECTEDDEPTID)
'------------------------------------------------------------------------------------------------------------
Function fnListDepts(iSelectedDeptID)

  ' SET SELECTED DEPARTMENT ID
  If iSelectedDeptID = "" Then
	iSelectedDeptID = 0
  End If

  ' GET LIST OF DEPARTMENTS
  If blnCanViewAllActionItems Then
	' GET ALL DEPARTMENTS
	sSQL = "select groupid,orgid,groupname,groupdescription from groups where grouptype=2 AND orgid=" & Session("OrgID") & " order by groupname"
  Else
	' GET DEPARTMENTS ASSIGNED TO CURRENTLY LOGGED ON ADMIN USER
	sSQL = "select groupid,orgid,groupname,groupdescription from groups where grouptype=2 AND orgid=" & Session("OrgID") & " AND groupid IN (" & GetGroups(session("user_id")) & ") order by groupname"
  End If

  Set oDepts = Server.CreateObject("ADODB.Recordset")
  oDepts.Open sSQL, Application("DSN") , 3, 1

  ' LOOP THRU GROUPS DISPLAYING AS NECESSARY
	Do while not oDepts.EOF
		' SET SELECTED DEPARTMENT
		If IsNumeric(iSelectedDeptID) Then
			If clng(iSelectedDeptID) = oDepts("groupid") Then
				' SET SELECTED FLAG
				sSelected = " selected" 
			Else 
				' CLEAR SELECTED FLAG
				sSelected = ""
			End If
		End If
		
		' DISPLAY GROUP AS OPTION
		response.write "<option " & sSelected  & " value=" & oDepts("groupid") & ">" & oDepts("groupname") & "</option>"

		oDepts.MoveNext
	Loop	

	' CLEAN UP OBJECTS
	Set oDepts = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' SUB DRAWASSIGNEDEMPLOYEESELECTION(IORGID)
'------------------------------------------------------------------------------------------------------------
Sub DrawAssignedEmployeeSelection(iorgid)

	sSQLassignedto = "SELECT FirstName + ' ' + LastName as assigned_Name, UserID FROM USERS where OrgID = " & Session("OrgID") & " ORDER BY FirstName,LastName"
	Set oAssigned = Server.CreateObject("ADODB.Recordset")
	oAssigned.Open sSQLassignedto, Application("DSN"), 3, 1

	' IF THERE ARE ASSIGNED USERS THEN LIST
	If  NOT oAssigned.EOF Then
	
		' BEGIN SELECTION BOX
		response.write "<select name=""selectAssignedto""><option value=""all"">Anyone</option>"
		
		' LOOP THRU ASSIGNED USERS
		Do While NOT oAssigned.EOF		
		
			' SET SELECT BOX TO DISPLAY CURRENTLY SELECTED NAME
			If selectAssignedto = clng(oAssigned("userid")) then 
				selectAssign = "selected"
			Else
				selectAssign = ""
			End If
			
			' DISPLAY ASSIGNED EMPLOYEE AS OPTION
			response.write "<option value=""" & oAssigned("assigned_Name") & """ " & selectAssign & ">" & oAssigned("assigned_Name") & "</option>" & vbcrlf

			oAssigned.MoveNext
		Loop

		response.write "</select>&nbsp;&nbsp;&nbsp;"
		
		oAssigned.Close
	
	End If

	Set oAssigned = Nothing

End Sub
%>

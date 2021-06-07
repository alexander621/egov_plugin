<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->

<%
sSQLadmin = "SELECT * FROM dbo.UsersGroupsPlus where UserID = " & Session("UserID")
Set oAdmin = Server.CreateObject("ADODB.Recordset")
oAdmin.Open sSQLadmin, Application("DSN"), 3, 1
if oAdmin.EOF then
	ViewAll = 0
elseif oAdmin("GroupName") = "Administrators" then 
	ViewAll = 1 
else 
	ViewAll = 0
end if

if request("useSessions")=1 then
			recordsPer          = session("recordsPer")
			reportType          = session("reportType")
			
			groupBy             = session("groupBy")
			selectFormId        = session("selectFormId")
			selectAssignedto    = session("selectAssignedto")
			selectDeptId        = session("selectDeptId")
			pastDays            = session("pastDays")
			
			selectUserFName     = session("selectUserFName")
			selectUserLName     = session("selectUserLName")
			
			fromDate            = session("fromDate")
			toDate              = session("toDate")
			today               = Date()
			
   statusSubmitted     = session("statusSubmitted")
   statusInprogress    = session("statusInprogress")
	  statusWaiting       = session("statusWaiting")
	  statusResolved      = session("statusResolved")
	  statusDismissed     = session("statusDismissed")

   substatus_hidden    = session("substatus_hidden")

			selectIssueStreet   = session("selectIssueStreet")
   selectContactStreet = session("selectContactSteet")
else
			recordsPer          = request("recordsPer")
			reportType          = request("reportType")
			
			groupBy             = request("groupBy")
			selectFormId        = request("selectFormId")
			selectAssignedto    = request("selectAssignedto")
			selectDeptId        = request("selectDeptId")
			pastDays            = request("pastDays")
			
			selectUserFName     = request("selectUserFName")
			selectUserLName     = request("selectUserLName")
			
			fromDate            = Request("fromDate")
			toDate              = Request("toDate")
			today               = Date()
			
   statusSubmitted     = request("statusSUBMITTED")
   statusInprogress    = request("statusINPROGRESS")
   statusWaiting       = request("statusWAITING")
   statusResolved      = request("statusRESOLVED")
   statusDismissed     = request("statusDISMISSED")

   substatus_hidden    = request("substatus_hidden")

			selectIssueStreet   = request("selectIssueStreet")
   selectContactStreet = request("selectContactStreet")
end if

If reportType = "" or IsNull(reportType) Then reportType = "List" End If

''If orderBy = "" or IsNull(orderBy) Then orderBy = " submit_date" End If
If groupBy = "" or IsNull(groupBy) Then 
  	groupBy = request("orderBy")
  	If groupBy = "" or IsNull(groupBy) Then
		    groupBy = "action_Formid" 
  	End If
end if

If selectFormId = "" or IsNull(selectFormId) Then selectFormId = "all" End If
If selectAssignedto = "" or IsNull(selectAssignedto) Then selectAssignedto = "all" End If
If selectDeptId = "" or IsNull(selectDeptId) Then selectDeptId = "all" End If

'SET CONTACT STREET FILTER
If selectContactStreet = "" or IsNull(selectContactStreet) Then 
  	selectContactStreet = "all" 
End If

'SET ISSUE/PROBLEM LOCATION STREET FILTER
If selectIssueStreet = "" or IsNull(selectIssueStreet) Then 
  	selectIssueStreet = "all" 
End If

If pastDays = "" or IsNull(pastDays) Then pastDays = "all" End If

If selectUserFName = "" or IsNull(selectUserFName) Then selectUserFName = "all" End If
If selectUserLName = "" or IsNull(selectUserLName) Then selectUserLName = "all" End If

If toDate = "" or IsNull(toDate) Then toDate = dateAdd("d",0,today) End If
If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("m",-1,today) End If
toDate = dateAdd("d",1,toDate)

if recordsPer = "" or IsNull(recordsPer) Then recordsPer = 20 End If

'SET FILTER STATUS
'noStatus = "true"

If statusSubmitted = "yes" Then 
   noStatus = "false"
ELSE
   statusSubmitted = "no"
End If

If statusInprogress = "yes" Then 
   noStatus = "false"
ELSE
   statusInprogress = "no"
End If

If statusWaiting = "yes" Then 
   noStatus = "false"
ELSE
   statusWaiting = "no"
End If

If statusResolved = "yes" Then 
   noStatus = "false"
ELSE
   statusResolved = "no"
End If

If statusDismissed = "yes" Then 
   noStatus = "false"
ELSE
   statusDismissed = "no"
End If

'if noStatus = "true" then
if request("init") = "Y" then
   statusSubmitted  = "yes"
   statusInprogress = "yes"
   statusWaiting    = "yes"
   statusResolved   = "yes"
   statusDismissed  = "yes"
end if

sSQLassignedto = "SELECT FirstName + ' ' + LastName as assigned_Name, UserID FROM USERS where OrgID = " & Session("OrgID") & " ORDER BY FirstName"
Set oAssigned = Server.CreateObject("ADODB.Recordset")
oAssigned.Open sSQLassignedto, Application("DSN"), 3, 1

Function fnListForms()
	sLastCategory = "NONE_START"
	sSQL = "SELECT * FROM dbo.egov_FormList order by form_category_Sequence,action_form_name"

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

Function fnListDepts()
	sSQL = "select groupid,orgid,groupname,groupdescription  from groups where orgid=" & Session("OrgID") & " order by groupname"
  	Set oDepts = Server.CreateObject("ADODB.Recordset")
	oDepts.Open sSQL, Application("DSN") , 3, 1
	if selectDeptId = "all" then
		  do while not oDepts.EOF
				response.write "<option value=" & oDepts("groupid") & ">" & oDepts("groupname") & "</option>"
			oDepts.MoveNext
			Loop	
	else
		  do while not oDepts.EOF
			if clng(selectDeptId) = oDepts("groupid") then selected = " selected" else selected = ""
				response.write "<option value=" & oDepts("groupid") & " " & selected & ">" & oDepts("groupname") & "</option>"
			oDepts.MoveNext
			Loop
	end if

	Set oDepts = Nothing
	End Function
%>

<html>
<head>
  <title><%=langBSActionLine%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script> 
<script language="JavaScript">
  function printit(){
 		if (window.print) {
 			window.print() ;
 		} else {
 			var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
 			document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
 			WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box
 			WebBrowser1.outerHTML = "";
 		}
 	}
	</script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  
  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
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
Function List_Action_Requests(sSortBy)

' LIST ACTION REQUESTS 

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
'REDIM statArray(i-1) 

for u = 0 to ubound(statArray)
	varStatClause = varStatClause & "" & statArray(u)
next
lenStatClause = len(varStatClause) - 3
if lenStatClause > 1 then
	varStatClause = left(varStatClause,lenStatClause)
end if

''**varWhereClause = " WHERE (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') AND (" & varStatClause & ")"
varWhereClause = " WHERE egov_action_request_view.orgid=('"&session("orgid")&"') AND ( "    ''IsNull(complete_date,'" & Now & "')
		varWhereClause = varWhereClause & " (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') OR "
		varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
		varWhereClause = varWhereClause & " (submit_date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "

'varWhereClause = varWhereClause & " ) AND (" & varStatClause & ")"

'Sub-Status Filter
if substatus_hidden = "" then
  'If any Statuses are checked
   if i > 0 then
      varWhereClause = varWhereClause & " ) AND (" & varStatClause & ") "
   else
      varWhereClause = varWhereClause & " ) "
   end if
else
  'If any Statuses are checked
   if i > 0 then
      varWhereClause = varWhereClause & " ) AND ((" & varStatClause & ") "
      varWhereClause = varWhereClause & " OR sub_status_id in (" & substatus_hidden & ")) "
   else
      varWhereClause = varWhereClause & " ) AND sub_status_id in (" & substatus_hidden & ") "
   end if
end if


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
		'varWhereClause = varWhereClause & " AND form_category_id = " & right(selectFormId,len(selectFormId)-1)
	else
		varWhereClause = varWhereClause & " AND action_Formid = " & selectFormId
	end if
end if

'If selectAssignedto    <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '"   & selectAssignedto    & "'" 
If selectAssignedto    <> "all" then varWhereClause = varWhereClause & " AND assignedemployeeid = "   & selectAssignedto
If selectDeptId        <> "all" then varWhereClause = varWhereClause & " AND deptID = "           & selectDeptId        & " " 
If selectContactStreet <> "all" Then varWhereClause = varWhereClause & " AND useraddress LIKE '%" & selectContactStreet & "%'"
If selectIssueStreet   <> "all" Then varWhereClause = varWhereClause & " AND streetname LIKE '%"  & selectIssueStreet   & "%'"

''If pastDays <> "all" then varWhereClause = varWhereClause & " AND DateDiff(d,submit_date,complete_date) >= " & pastDays & " " 


if groupBy="submit_date" then
				''sSQL = "SELECT submitdateshort as TheDate,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays,count(*) as numSubmitted FROM egov_action_request_view " & varWhereClause & " GROUP BY submitdateshort ORDER BY submitdateshort desc"
				sSQL = "SELECT submitdateshort as TheDate,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays,count(*) as numSubmitted FROM egov_action_request_view " & varWhereClause & " GROUP BY submitdateshort ORDER BY submitdateshort desc"
elseif groupBy="action_Formid"  then 
				sSQL = "SELECT action_formTitle,action_Formid,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays,count(*) as numSubmitted FROM egov_action_request_view " & varWhereClause & " GROUP BY action_formTitle,action_Formid ORDER BY action_formTitle"
elseif groupBy="deptId"  then 
				sSQL = "SELECT deptID,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays,count(*) as numSubmitted FROM egov_action_request_view " & varWhereClause & " GROUP BY deptID ORDER BY deptID"
elseif groupBy="assigned_Name"  then 
				sSQL = "SELECT assigned_Name,count(*) as numSubmitted,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays FROM egov_action_request_view " & varWhereClause & " GROUP BY assigned_Name ORDER BY assigned_Name"
ElseIf groupBy="streetname"  Then 
	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests,streetaddress,streetnumber,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays,count(*) as numSubmitted FROM egov_action_request_view " & varWhereClause & " GROUP BY streetaddress, streetnumber ORDER BY UPPER(streetaddress), CAST(streetnumber AS int) "
End If 


''response.write sSQL
Set oRequests = Server.CreateObject("ADODB.Recordset")
oRequests.Open sSQL, Application("DSN"), 3, 1


	 
if oRequests.EOF=false then
	 ' REMOVED SET PAGE TO VIEW
	 
	 ' REMOVED DISPLAY RECORD STATISTICS
	   ''Response.Write "<b><font color=blue>" & oRequests.RecordCount & "</font> total Action Item Requests</b>"
  	 ''Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  	 Response.Write " <font color=3399ff><i><b>" & Ucase(Replace(ReportType,"DrillThru","Drill Through")) & " REPORT</b></i></font>"
	 
	 ' REMOVED DISPLAY FORWARD AND BACKWARD NAVIGATION TOP AND PRINT PAGE LINK
 	  
	  Response.Write "<table cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">"
	  Response.Write "<tr class=tablelist>"
	  
		 if groupBy="submit_date" then
		 					Response.Write "<th>Date</th>"
		 elseif groupBy="action_Formid"  then 
		 					Response.Write "<th>Action Line Category</th>"
		 elseif groupBy="deptId"  then 
		 					Response.Write "<th>Department</th>"
		 elseif groupBy="assigned_Name"  then 
		 					Response.Write "<th>Assigned To</th>"
		 ElseIf groupby = "streetname" Then
    				Response.Write "<th>Issue/Problem Location Street Name</th>"
		 end if

	  Response.Write "<th>Submitted</td><th>Open Items</td>"
	  if pastDays <> "all" then
 	  		Response.Write "<th>Open Items Over " & pastDays & " days</td>"
	  end if
	  Response.Write "<th>Avg. Time still Open</td>"
	  Response.Write "<th>Avg. Time to Complete</td></tr>"

totalSubmitted = 0
totalOpen      = 0
totalDays      = 0
totalPast      = 0

	  ' LOOP AND DISPLAY THE RECORDS
	  bgcolor = "#eeeeee"
		 Do while not oRequests.EOF

			If bgcolor="#eeeeee" Then
  				bgcolor="#ffffff" 
			Else
		  		bgcolor="#eeeeee"
			End If

		  ' GET VALUES

		  	if groupBy="submit_date" then
							If oRequests("TheDate") <> "" Then
								sTitle = oRequests("TheDate") 
							Else
								sTitle = "<font color=red><b>???</b></font>"
							End If

							''// OPENS
							sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND  submitdateshort='" & sTitle & "' AND (status<>'RESOLVED' AND status<>'DISMISSED') "
							Set oOpen = Server.CreateObject("ADODB.Recordset")
							oOpen.Open sSQLopen, Application("DSN"), 3, 1
							''// COMPLETES
							sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND  submitdateshort='" & sTitle & "' AND (status='RESOLVED' OR status='DISMISSED') "
							Set oClosed = Server.CreateObject("ADODB.Recordset")
							oClosed.Open sSQLclosed, Application("DSN"), 3, 1
							''// PAST DUES
							if pastDays <> "all" then pastDate = clng(pastDays) else pastDate = 10000
							sSQLpast = "SELECT count(action_autoid) as numPast FROM egov_action_request_view " & varWhereClause & "  AND DateDiff(d,submit_date,'" & Date() & "')>" & pastDate & " AND submitdateshort='" & sTitle & "' AND status<>'RESOLVED' AND status<>'DISMISSED' "
							Set oPast = Server.CreateObject("ADODB.Recordset")
							oPast.Open sSQLpast, Application("DSN"), 3, 1

							detaillink = "action_line_list.asp?orderBy=" & groupBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&toDate=" & oRequests("TheDate") & "&fromDate=" & oRequests("TheDate")  & "&reportType=DrillThru"
	
			elseif groupBy="action_Formid"  then 
							If oRequests("action_formTitle") <> "" Then
								sTitle = oRequests("action_formTitle") 
							Else
								sTitle = "<font color=red><b>???</b></font>"
							End If

							''// OPENS
							sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND  action_formTitle='" & sTitle & "' AND (status<>'RESOLVED' AND status<>'DISMISSED') "
							Set oOpen = Server.CreateObject("ADODB.Recordset")
							oOpen.Open sSQLopen, Application("DSN"), 3, 1
							''// COMPLETES
							sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND  action_formTitle='" & sTitle & "' AND (status='RESOLVED' OR status='DISMISSED') "
							Set oClosed = Server.CreateObject("ADODB.Recordset")
							oClosed.Open sSQLclosed, Application("DSN"), 3, 1
							''// PAST DUES
							if pastDays <> "all" then pastDate = clng(pastDays) else pastDate = 10000
							sSQLpast = "SELECT count(action_autoid) as numPast FROM egov_action_request_view " & varWhereClause & "  AND DateDiff(d,submit_date,'" & Date() & "')>" & pastDate & " AND action_formTitle='" & sTitle & "' AND status<>'RESOLVED' AND status<>'DISMISSED' "
							Set oPast = Server.CreateObject("ADODB.Recordset")
							oPast.Open sSQLpast, Application("DSN"), 3, 1

							detaillink = "action_line_list.asp?orderBy=" & groupBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & oRequests("action_formId") & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&toDate=" & toDate & "&fromDate=" & fromDate  & "&reportType=DrillThru"
		  	elseif groupBy="deptId"  then 
							If oRequests("deptId") <> "" AND IsNull(oRequests("deptId"))=false Then
								sTitle = clng(oRequests("deptId"))
							Else
								sTitle = 0
							End If

								''// OPENS
							sSQLopen = "SELECT count(deptId) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND  deptId=" & sTitle & " AND (status<>'RESOLVED' AND status<>'DISMISSED') "
							Set oOpen = Server.CreateObject("ADODB.Recordset")
							oOpen.Open sSQLopen, Application("DSN"), 3, 1
							''// COMPLETES
							sSQLclosed = "SELECT count(deptId) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND  deptId=" & sTitle & " AND (status='RESOLVED' OR status='DISMISSED') "
							Set oClosed = Server.CreateObject("ADODB.Recordset")
							oClosed.Open sSQLclosed, Application("DSN"), 3, 1
							''// PAST DUES
							if pastDays <> "all" then pastDate = clng(pastDays) else pastDate = 10000
							sSQLpast = "SELECT count(action_autoid) as numPast FROM egov_action_request_view " & varWhereClause & "  AND DateDiff(d,submit_date,'" & Date() & "')>" & pastDate & " AND deptId=" & sTitle & " AND status<>'RESOLVED' AND status<>'DISMISSED' "
							Set oPast = Server.CreateObject("ADODB.Recordset")
							oPast.Open sSQLpast, Application("DSN"), 3, 1
							
							detaillink = "action_line_list.asp?orderBy=" & groupBy & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&toDate=" & toDate & "&fromDate=" & fromDate  & "&reportType=DrillThru&selectDeptId=" & oRequests("DeptID")

							sSQLdeptName = "select groupname  from groups where orgid=" & Session("OrgID") & " AND groupid=" & sTitle
						  	Set oDeptName = Server.CreateObject("ADODB.Recordset")
							oDeptName.Open sSQLdeptName, Application("DSN") , 3, 1
							
							If oDeptName.EOF  Then
								sTitle = "<font color=red><b>???</b></font>"
							else
								sTitle = oDeptName("groupname") 
							End If

			elseif groupBy="assigned_Name"  then 
							If oRequests("assigned_Name") <> "" Then
  								sTitle = oRequests("assigned_Name") 
							Else
		  						sTitle = "<font color=""red""><b>???</b></font>"
							End If

							''// OPENS
							sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND  assigned_Name='" & oRequests("assigned_Name") & "' AND (status<>'RESOLVED' AND status<>'DISMISSED') "
							Set oOpen = Server.CreateObject("ADODB.Recordset")
							oOpen.Open sSQLopen, Application("DSN"), 3, 1
							''// COMPLETES
							sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND  assigned_Name='" & oRequests("assigned_Name") & "' AND (status='RESOLVED' OR status='DISMISSED') "
							Set oClosed = Server.CreateObject("ADODB.Recordset")
							oClosed.Open sSQLclosed, Application("DSN"), 3, 1
							''// PAST DUES
							if pastDays <> "all" then pastDate = clng(pastDays) else pastDate = 10000
							sSQLpast = "SELECT count(action_autoid) as numPast FROM egov_action_request_view " & varWhereClause & "  AND DateDiff(d,submit_date,'" & Date() & "')>" & pastDate & " AND assigned_Name='" & oRequests("assigned_Name") & "' AND status<>'RESOLVED' AND status<>'DISMISSED' "
							Set oPast = Server.CreateObject("ADODB.Recordset")
							oPast.Open sSQLpast, Application("DSN"), 3, 1
			
							detaillink = "action_line_list.asp?orderBy=" & groupBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & oRequests("assigned_Name") & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&toDate=" & toDate & "&fromDate=" & fromDate  & "&reportType=DrillThru"

				ElseIf groupBy = "streetname" Then 
       lcl_street = ""
       if oRequests("streetnumber") <> "" then
          lcl_street = oRequests("streetnumber")
       end if

       if oRequests("streetaddress") <> "" then
          if lcl_street = "" then
             lcl_street = oRequests("streetaddress")
          else
             lcl_street = lcl_street & " " & oRequests("streetaddress")
          end if
							end if

       if lcl_street = "" then
    						sTitle = "<font color=""red""><b>???</b></font>"
       else
          sTitle = lcl_street
       end if
				
							''// OPENS
							sSQLopen = "SELECT count(*) as numOpen, SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen "
       sSQLopen = sSQLopen & " FROM egov_action_request_view "
       sSQLopen = sSQLopen & varWhereClause
       sSQLopen = sSQLopen & " AND (status<>'RESOLVED' AND status<>'DISMISSED') "
       sSQLopen = sSQLopen & " AND isnull(streetaddress,' ') = '" & oRequests("streetaddress") & "'"
       sSQLopen = sSQLopen & " AND isnull(streetnumber, ' ') = '" & oRequests("streetnumber")  & "'"

							Set oOpen = Server.CreateObject("ADODB.Recordset")
							oOpen.Open sSQLopen, Application("DSN"), 0, 1
							oOpen.movefirst

							''// COMPLETES
							sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed "
       sSQLclosed = sSQLclosed & " FROM egov_action_request_view "
       sSQLclosed = sSQLclosed & varWhereClause
       sSQLclosed = sSQLclosed & " AND (status='RESOLVED' OR status='DISMISSED') "
       sSQLclosed = sSQLclosed & " AND isnull(streetaddress,' ') = '" & oRequests("streetaddress") & "'"
       sSQLclosed = sSQLclosed & " AND isnull(streetnumber,' ') = '"  & oRequests("streetnumber")  & "'"
							Set oClosed = Server.CreateObject("ADODB.Recordset")
							oClosed.Open sSQLclosed, Application("DSN"), 0, 1

							''// PAST DUES
							if pastDays <> "all" then pastDate = clng(pastDays) else pastDate = 10000
							sSQLpast = "SELECT count(action_autoid) as numPast "
       sSQLpast = sSQLpast & " FROM egov_action_request_view "
       sSQLpast = sSQLpast & varWhereClause
       sSQLpast = sSQLpast & " AND DateDiff(d,submit_date,'" & Date() & "')>" & pastDate
       sSQLpast = sSQLpast & " AND status<>'RESOLVED' "
       sSQLpast = sSQLpast & " AND status<>'DISMISSED' "
       sSQLpast = sSQLpast & " AND isnull(streetaddress,' ') = '" & oRequests("streetaddress") & "'"
       sSQLpast = sSQLpast & " AND isnull(streetnumber,' ') = '"  & oRequests("streetnumber")  & "'"
							Set oPast = Server.CreateObject("ADODB.Recordset")
							oPast.Open sSQLpast, Application("DSN"), 0, 1
							
							detaillink = "action_line_list.asp?orderBy=" & groupBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & toDate & "&fromDate=" & fromDate & "&selectIssueStreet=" & lcl_street & "&reportType=DrillThru"
			end if

			''//SUBTOTAL SUBMITTED	
			numSubmitted   = oRequests("numSubmitted")
			totalSubmitted = totalSubmitted + numSubmitted
			
			''//SUBTOTAL OPEN	
			numOpen = clng(oOpen("numOpen"))
			if numOpen<>0 and oOpen("totalDaysOpen")<> 0 then
 					avgOpen = oOpen("totalDaysOpen") / numOpen
	 				avgOpen = formatnumber(avgOpen,1)
			else
		  		If numOpen > 0 Then
				  	 'the datediff is 0 but there are some open items, so they are from today.
    					avgOpen = "< 1.0"
				  Else 
   					'response.write("'" & avgOpen & "'")
			    		avgOpen = "None Open"
  				End If 
			end if
			totalOpen = totalOpen + numOpen

			if IsNull(oOpen("totalDaysOpen")) then
			else	
					 totalOpenDays = totalOpenDays + oOpen("totalDaysOpen")
			end if
			
			''//SUBTOTAL CLOSED
			numClosed = clng(oClosed("numClosed")) 
			if numClosed<>0 and oClosed("totalDaysClosed")<> 0 then
 					avgClosed = oClosed("totalDaysClosed") / numClosed
	 				avgClosed = formatnumber(avgClosed,1)
			else
		  		If numClosed > 0 Then
				   	'Handle datediff is 0 but some have been completed, so they were completed the same day.
    					avgClosed = " < 1.0 "
  				Else
		    			avgClosed = " None Completed "
  				End If
			end if
			totalClosed = totalClosed + numClosed

			if IsNull(oClosed("totalDaysClosed")) then
			else		
 					totalClosedDays = totalClosedDays + oClosed("totalDaysClosed")
			end if

			if pastDays <> "all" then
 					numPast = oPast("numPast")
	 				totalPast =totalPast + numPast
			end if					

			response.write "<tr bgcolor=""" & bgcolor & """ onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">"
   response.write "    <td onClick=""location.href='" & detaillink & "';""><b>" & sTitle & " </b></td>"
			response.write "    <td align=""center"" onClick=""location.href='" & detaillink & "';"">" & numSubmitted  & "</td>"
   response.write "    <td align=""center""> " & numOpen  & "</td>"
			if pastDays <> "all" then
  				response.write "<td align=""center""> " & numPast  & "</td>"
			end if
			response.write "<td align=""center""> " & avgOpen  & "</td>"
   response.write "<td align=""center""> " & avgClosed  & "</td></tr>"
	oRequests.MoveNext 
	Loop

			if totalClosed<>0 and totalClosedDays<> 0 then
 					avgClosed = totalClosedDays / totalClosed
	 				avgClosed = formatnumber(avgClosed,1)
			else
		  		If totalClosed > 0 Then
   					'Handle datediff is 0 but some have been completed, so they were completed the same day.
    					avgClosed = " < 1.0 "
				  Else
					    avgClosed = ""
  				End If
			end if

			if totalOpen<>0 and totalOpenDays<> 0 then
 					avgOpen = totalOpenDays / totalOpen
	 				avgOpen = formatnumber(avgOpen,1)
			else
		  		If totalOpen > 0 Then
				   	'the datediff is 0 but there are some open items, so the are from today.
    					avgOpen = "< 1.0"
				  Else 
					   'response.write("'" & avgOpen & "'")
    					avgOpen = ""
				  End If 
			end if

			response.write "<tr bgcolor=""#dddddd"">"
   response.write "    <td style=""padding-left:90px""><b> <font color=""navy"" size=""1"">TOTAL</td>"
   response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" & totalSubmitted & "</td>"
   response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" & totalOpen & "</td>"
			if pastDays <> "all" then
 					response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" & totalPast & "</td>"
			end if
			response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" & avgOpen & " days</td>"
   response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" & avgClosed & " days</td>"
   response.write "</tr>"

		 response.write "</table>"

	'''** REMOVED DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM

else
	Response.write "<p><b>No records found</p>"
end if

End Function
%>


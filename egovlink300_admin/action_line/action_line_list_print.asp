<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->

<%
' GET USER'S PERMISSIONS
blnCanViewAllActionItems  = HasPermission("CanViewAllActionItems")
blnCanViewOwnActionItems  = HasPermission("CanViewOwnActionItems")
blnCanViewDeptActionItems = HasPermission("CanViewDeptActionItems")

if request("useSessions")=1 then
   recordsPer          = session("recordsPer")
   reporttype          = session("reporttype")

   orderBy             = session("orderBy")
   selectFormId        = session("selectFormId")
   selectAssignedto    = session("selectAssignedto")
   selectDeptId        = session("selectDeptId")

   selectUserFName     = session("selectUserFName")
   selectUserLName     = session("selectUserLName")   

   fromDate            = session("fromDate")
   toDate              = session("toDate")
   today               = Date()

'   statusSubmitted  = session("statusSubmitted")
'   statusInprogress = session("statusInprogress")
'   statusWaiting    = session("statusWaiting")
'   statusResolved   = session("statusResolved")
'   statusDismissed  = session("statusDismissed")

   statusSubmitted     = session("statusSubmitted")
   statusInprogress    = session("statusInprogress")
   statusWaiting       = session("statusWaiting")
   statusResolved      = session("statusResolved")
   statusDismissed     = session("statusDismissed")

   substatus_hidden    = session("substatus_hidden")

   selectIssueStreet   = session("selectIssueStreet")
   selectContactStreet = session("selectContactStreet")
else
   recordsPer          = request("recordsPer")
   reporttype          = request("reporttype")

   orderBy             = request("orderBy")
   selectFormId        = request("selectFormId")
   selectAssignedto    = request("selectAssignedto")
   selectDeptId        = request("selectDeptId")

   selectUserFName     = request("selectUserFName")
   selectUserLName     = request("selectUserLName")

   fromDate            = Request("fromDate")
   toDate              = Request("toDate")
   today               = Date()

'   statusSubmitted  = request("statusSubmitted")
'   statusInprogress = request("statusInprogress")
'   statusWaiting    = request("statusWaiting")
'   statusResolved   = request("statusResolved")
'   statusDismissed  = request("statusDismissed")

   statusSubmitted     = request("statusSUBMITTED")
   statusInprogress    = request("statusINPROGRESS")
   statusWaiting       = request("statusWAITING")
   statusResolved      = request("statusRESOLVED")
   statusDismissed     = request("statusDISMISSED")

   substatus_hidden    = request("substatus_hidden")

   selectIssueStreet   = request("selectIssueStreet")
   selectContactStreet = request("selectContactStreet")
end if

If reporttype = "" or IsNull(reporttype) Then reporttype = "List" End If

If orderBy = "" or IsNull(orderBy) Then 
   orderBy = request("orderBy")
   if orderBy = "" or IsNull(orderBy) Then
      orderBy = "submit_date" 
   end if
End If

if selectFormId     = "" or IsNull(selectFormId)     then selectFormId     = "all" end if
if selectAssignedto = "" or IsNull(selectAssignedto) then selectAssignedto = "all" end if
if selectDeptId     = "" or IsNull(selectDeptId)     then selectDeptId     = "all" end if

if selectUserFName  = "" or IsNull(selectUserFName)  then selectUserFName  = "all" end if
if selectUserLName  = "" or IsNull(selectUserLName)  then selectUserLName  = "all" end if

if toDate           = "" or IsNull(toDate)           then toDate           = dateAdd("d",0,today)     end if
if fromDate         = "" or IsNull(fromDate)         then fromDate         = dateAdd("yyyy",-1,today) end if

toDate = dateAdd("d",1,toDate)

if recordsPer = "" or IsNull(recordsPer) then recordsPer = 25 end if

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

%>
<html>
<head>
<title><%=langBSActionLine%></title>
<link href="../global.css" rel="stylesheet" type="text/css">
<script src="../scripts/selectAll.js"></script>
<script language="JavaScript">
  function printit(){
    if(window.print) {
	   window.print() ;
    } else {
       var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
       document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
       WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box
       WebBrowser1.outerHTML = "";
    }
}
  function changeRowColor(pID,pStatus) {
    if(pStatus=="OVER") {
       document.getElementById(pID).style.cursor          = "hand";
       document.getElementById(pID).style.backgroundColor = "#93bee1";
    }else{
       document.getElementById(pID).style.cursor          = "";
	   document.getElementById(pID).style.backgroundColor = "";
	}
  }
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
<!--
  <tr><td><font size="+1"><b>Manage Action Line Requests</b></font></td>
      <td width="350"><B><a href="javascript:printit()">PRINT</a></a> | <b><a href="javascript:window.close()">CLOSE WINDOW</A></td></tr>-->
  <tr><td colspan="2" valign="top">
          <!--BEGIN: ACTION LINE REQUEST LIST -->
          <% List_Action_Requests(sSortBy) %>
          <!-- END: ACTION LINE REQUEST LIST -->
      </td></tr>
</table>
</body>
</html>
<%
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
If statusResolved = "yes" Then
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

''varWhereClause = " WHERE (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') AND (" & varStatClause & ")"

varWhereClause = " WHERE egov_action_request_view.orgid=('"&session("orgid")&"') AND ( "    ''IsNull(complete_date,'" & Now & "')
varWhereClause = varWhereClause & " (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') OR "
varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
varWhereClause = varWhereClause & " (submit_date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "

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

'varWhereClause = varWhereClause & " ) AND (" & varStatClause & ")"

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

'If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '" & selectAssignedto & "'" 
If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assignedemployeeid = " & selectAssignedto

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

'SET ISSUE/PROBLEM LOCATION STREET FILTER
If selectIssueStreet = "" or IsNull(selectIssueStreet) Then selectIssueStreet = "all" End If

'SET CONTACT STREET FILTER
If selectContactStreet = "" or IsNull(selectContactStreet) Then selectContactStreet = "all" End If

'ISSUE/PROBLEM LOCATION STREET NAME FILTER
If selectIssueStreet <> "all" then varWhereClause = varWhereClause & " AND streetname LIKE '%" & selectIssueStreet & "%'"

'CONTACT STREET NAME FILTER
If selectContactStreet <> "all" then varWhereClause = varWhereClause & " AND useraddress LIKE '%" & selectContactStreet & "%'"

'sSQL = "SELECT userlname, userfname, action_autoid, action_formTitle, DateDiff(d,submit_date,complete_date) AS totalDays, "
'sSQL = sSQL & "submit_date, complete_date, deptID, groupname as deptName, status, assigned_Name "
'sSQL = sSQL & "FROM egov_action_request_view left outer join groups on deptId=groupId"
'sSQL = sSQL & varWhereClause &


sSQL = "SELECT userlname, userfname, userhomephone, useraddress, usercity, userstate, userzip, action_autoid, action_formTitle, "
sSQL = sSQL & " DateDiff(d,submit_date,ISNULL(complete_date,'" & today & "')) AS totalDays, "
sSQL = sSQL & " DateDiff(d,ISNULL(adjustedsubmitdate,submit_date),ISNULL(complete_date,'" & today & "')) AS totalDays_adjusted, "
sSQL = sSQL & " [dbo].[getDateDiff_NoWeekend] (egov_action_request_view.orgid,egov_action_request_view.submit_date,ISNULL(egov_action_request_view.complete_date,'" & today & "')) AS totalDays_noweekends, "
sSQL = sSQL & " [dbo].[getDateDiff_NoWeekend] (egov_action_request_view.orgid,ISNULL(egov_action_request_view.adjustedsubmitdate,egov_action_request_view.submit_date),ISNULL(egov_action_request_view.complete_date,'" & today & "')) AS totalDays_noweekends_adjusted, "
sSQL = sSQL & " submit_date, adjustedsubmitdate, complete_date, deptID, groupname as deptName, status, "
sSQL = sSQL & " assignedName as assigned_Name, assignedemployeeid, latitude, longitude, streetnumber, streetaddress, streetname, comment, "
sSQL = sSQL & " allowedunresolveddays, usesafter5adjustment, issuelocationname, usesweekdays, city, state, zip, validstreet, comments, action_form_display_issue, "
sSQL = sSQL & " (select ISNULL(status_name,'') from egov_actionline_requests_statuses where sub_status_id = action_status_id) AS sub_status_name "
sSQL = sSQL & " FROM egov_action_request_view left outer join groups on deptId = groupId "
sSQL = sSQL & varWhereClause
'sSQL = sSQL & " AND (egov_action_request_view.orgid=" & session("orgid") & ")"

if orderBy = "streetname" then
   lcl_order_by = "UPPER(streetaddress), CAST(streetnumber AS int) "
else
   lcl_order_by = orderBy
end if

sSQL = sSQL & " ORDER BY " & lcl_order_by
if orderBy = "submit_date" then
   sSQL = sSQL & " desc"
end if

''**response.write sSQL
Set oRequests = Server.CreateObject("ADODB.Recordset")
' OPEN RECORDSET
oRequests.Open sSQL, Application("DSN"), 3, 1

lastTitle        = "Test"
lastDate         = "1/1/02"
lastDept         = 11798
lastDeptName     = "Test"
lastAssigned     = "bubba"
displayLastTitle = "Test"

if oRequests.EOF=false then
  'DISPLAY RECORD STATISTICS
   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">"
   response.write "  <tr><td><b><font color=""blue"">" & oRequests.RecordCount & "</font> total Action Item Requests</b>"
   response.write           "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
   response.write           "<font color=""#3399ff""><i><b>" & Ucase(Replace(Replace(reporttype,"DrillThru","Drill Through"),"Full"," (FULL)")) & " REPORT</b></i></font>"
   response.write "      </td>"
   response.write "      <td align=""right"">"
      if OrgHasFeature("issue location") then
         response.write "<font style=""color: #FF0000;"">* <small><i>= Non-Listed Street Address</i></small></font>"
      end if
   response.write "      </td>"
   response.write "  </tr>"
   response.write "</table>"
   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">"

   if reporttype <> "ListFull" then
	 'DISPLAY FORWARD AND BACKWARD NAVIGATION TOP
      Response.Write "  <tr valign=""bottom"" class=""tablelist"">" '''**<th>Action Line Category</th><th>Date Submitted</td><th>Name</td><th>Status</td><th>Action</th></tr>"
      Response.Write "      <th>Action Line Category</th>"
      response.write "      <th>Date submitted</th>"

      if reporttype="Detail" or reporttype="DrillThru" then
         Response.Write "<th>Date Completed</th>"
         response.write "<th>Days open*/To complete</th>"
      end if

      Response.Write "<th>Status</th>"

   	  if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
	       Response.Write "<th>Sub-Status</th>"
      end if

      response.write "<th>Submitted by</th>"
      response.write "<th>Contact<br>Street Name</th>"
      response.write "<th>Assigned to</th>"
      response.write "<th>Department</th>"

      if OrgHasFeature("issue location") then
         response.write "<th>Issue/Problem Location<br>Street Name</th>"
      end if
	  response.write "</tr>"
   else
      response.write "<tr class=""tablelist"">"
	  response.write "    <th>&nbsp;</th></tr>"
   end if

'-------------------------------------------------------------------
   sSQLtotl    = "SELECT action_autoid,action_formTitle,DateDiff(d,submit_date,complete_date) AS totalDays,submit_date,complete_date,deptID,groupname as deptName,status,assigned_Name FROM egov_action_request_view left outer join groups on deptId=groupId" & varWhereClause
   Set oTotals = Server.CreateObject("ADODB.Recordset")
   oTotals.Open sSQLtotl, Application("DSN"), 3, 1

   sSQLTotal  = "SELECT count(*) as numTotal FROM egov_action_request_view " & varWhereClause    '''////& " AND status<>'RESOLVED' AND status<>'DISMISSED'"
   Set oTotal = Server.CreateObject("ADODB.Recordset")
   oTotal.Open sSQLTotal, Application("DSN"), 3, 1

   sSQLopen  = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') "
   Set oOpen = Server.CreateObject("ADODB.Recordset")
   oOpen.Open sSQLopen, Application("DSN"), 3, 1			

   sSQLclosed  = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED') "
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

   if reporttype="Detail" or reporttype="DrillThru" then	
      Response.Write "<tr bgcolor=""#dddddd"">"
      response.write "    <td style=""padding-left:90px"" colspan=""2""><b><font color=""navy"" size=""1"">Grand Total</td>"
      response.write "    <td colspan=""2""> <b><font color=""navy"" size=""1""> Submitted: "                          & oTotal("numTotal") & " </td>"
      response.write "    <td colspan=""2""> <b><font color=""navy"" size=""1""> Open: "                               & oOpen("numOpen")   & " </td>"
      response.write "    <td colspan=""2""> <b><font color=""navy"" size=""1""> Avg Time Still Open: "  & avgOpenTotal       & " </td>"
      response.write "    <td colspan=""3""> <b><font color=""navy"" size=""1""> Avg Time To Complete: " & avgClosedTotal     & " </td>"
      response.write "</tr>"
   end if

'LOOP AND DISPLAY THE RECORDS
bgcolor = "#eeeeee"
do while not oRequests.EOF
   If bgcolor="#eeeeee" Then
      bgcolor="#ffffff" 
   else
      bgcolor="#eeeeee"
   End If

  'GET VALUES
   If oRequests("action_formTitle") <> "" Then
      sTitle = oRequests("action_formTitle")
   Else
      sTitle = "<font color=""red""><b>???</b></font>"
   End If

   If oRequests("status") <> "" Then
      sStatus = oRequests("status")
   Else
      sStatus = "<font color=""red""><b>???</b></font>"
   End If

   ''If oRequests("submit_date") <> "" Then
   sDate = oRequests("submit_date")
   sDate = formatDateTime(sDate,vbShortDate)
   ''Else
   ''	sDate = "<font color=""red""><b>???</b></font>"
   ''End If

   If oRequests("assigned_Name") <> "" Then
      sAssigned = oRequests("assigned_Name")
   Else
      sAssigned = "<font color=""red""><b>???</b></font>"
   End If

   'If oRequests("deptId") <> "" Then
   sDept     = oRequests("deptId")
   sDeptName = oRequests("deptName")
   'Else
   '	sDept = "<font color=""red""><b>???</b></font>"
   'End If

   If oRequests("status") <> "" Then
      sStatus = oRequests("status")
   Else
      sStatus = "<font color=""red""><b>???</b></font>"
   End If

   if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
      sSubStatus = oRequests("sub_status_name")
   else
      sSubStatus = ""
   end if
   
   If oRequests("submit_date") <> "" Then
      datSubmitDate = oRequests("submit_date")
   Else
      datSubmitDate = "<font color=""red""><b>???</b></font>"
   End If

   If oRequests("complete_date") <> "" Then
      datResolveDate = oRequests("complete_date")
      datResolveDate = formatdatetime(datResolveDate,vbShortDate)
   Else
      datResolveDate = "<font color=""red""><b>???</b></font>"
   End If

   lngTrackingNumber = oRequests("action_autoid") & replace(FormatDateTime(oRequests("submit_date"),4),":","")

  '**INSERT SUBTOTAL ROW IF DETAIL REPORT**
   if reporttype="Detail" or reporttype="DrillThru" then
      if orderBy = "submit_date" then
         if DateDiff("d",sDate,lastDate) = 0 then
           'NO NEW LINE
         else
            if lastDate <> "1/1/02" then
               sSQLopen     = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND submitdateshort = '" & lastDate & "'"
               SQLopenText  = "<br>" & sSQLopen & "<br>"
               Set oOpenSub = Server.CreateObject("ADODB.Recordset")
               oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			

               sSQLclosed     = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND submitdateshort = '" & lastDate & "'"
               SQLclosedText  = "<br>" & sSQLclosed & "<br>"
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

               Response.Write "<tr bgcolor=""#dddddd"">"
         			   response.write "    <td style=""padding-left:90px"" colspan=""2""> <b> <font color=""navy"" size=""1""> Subtotal: " & lastDate      & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Submitted: "                            & subTotl       & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Open: "                                 & numOpen       & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Avg Time Still Open: "    & avOpenTotal   & " </td>"
          		   response.write "    <td colspan=""3""> <b> <font color=""navy"" size=""1""> Avg Time To Complete: "   & avClosedTotal & " </td>"
               response.write "</tr>"
            end if
         end if
      elseif orderBy = "action_Formid" then
         if sTitle = lastTitle then
           'NO NEW LINE
         else
            if lastTitle <> "Test" then
               sSQLopen     = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND action_FormTitle='" & lastTitle & "'"
               SQLopenText  = "<br>" & sSQLopen & "<br>"
               Set oOpenSub = Server.CreateObject("ADODB.Recordset")
               oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			

               sSQLclosed     = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND action_FormTitle='" & lastTitle & "'"
               SQLclosedText  = "<br>" & sSQLclosed & "<br>"
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

               displayLastTitle     = sTitle
               displayLastAvgOpen   = avOpenTotal
               displayLastAvgClosed = avClosedTotal

               Response.Write "<tr bgcolor=""#dddddd"">"
         			   response.write "    <td style=""padding-left:90px"" colspan=""2""><b> <font color=""navy"" size=""1"">Subtotal: "   & lastTitle     & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Submitted: "                            & subTotl       & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Open: "                                 & numOpen       & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Avg Time Still Open: "    & avOpenTotal   & " </td>"
         			   response.write "    <td colspan=""3""> <b> <font color=""navy"" size=""1""> Avg Time To Complete: "   & avClosedTotal & " </td>"
               response.write "</tr>"
            end if
         end if
      elseif orderBy = "deptId" then
         if sDept = lastDept then
           'NO NEW LINE
         else
            if lastDept <> 11798 then
               sSQLopen     = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND deptId = '" & lastDept & "'"
               Set oOpenSub = Server.CreateObject("ADODB.Recordset")
               oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			

               sSQLclosed     = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND deptId = '" & lastDept & "'"
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

               displayLastTitle     = sDeptName
               displayLastAvgOpen   = avOpenTotal
               displayLastAvgClosed =	avClosedTotal

               Response.Write "<tr bgcolor=""#dddddd"">"
         			   response.write "    <td style=""padding-left:90px"" colspan=""2""><b> <font color=""navy"" size=""1"">Subtotal: " & lastDeptName  & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Submitted: "                          & subTotl       & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Open: "                               & numOpen       & " </td>"
         			   response.write "    <td colspan=""2""><b><font color=""navy"" size=""1""> Avg Time Still Open: "    & avOpenTotal   & " </td>"
         			   response.write "    <td colspan=""3""><b><font color=""navy"" size=""1""> Avg Time To Complete: "   & avClosedTotal & " </td>"
               response.write "</tr>"
            end if
         end if
      elseif orderBy = "assigned_Name" then
         if sAssigned = lastAssigned then
           'NO NEW LINE
         else
            if lastAssigned <> "bubba" then			
               sSQLopen     = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysOpen FROM egov_action_request_view " & varWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') AND assigned_Name = '" & lastAssigned & "'"
               SQLopenText  = "<br>" & sSQLopen & "<br>"
               Set oOpenSub = Server.CreateObject("ADODB.Recordset")
               oOpenSub.Open sSQLopen, Application("DSN"), 3, 1			

               sSQLclosed     = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDaysClosed FROM egov_action_request_view " & varWhereClause & " AND (status='RESOLVED' OR status='DISMISSED')  AND assigned_Name = '" & lastAssigned & "'"
               SQLclosedText  = "<br>" & sSQLclosed & "<br>"
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

               displayLastTitle     = sAssigned
               displayLastAvgOpen   = avOpenTotal
               displayLastAvgClosed = avClosedTotal		

               Response.Write "<tr bgcolor=""#dddddd"">"
         			   response.write "    <td style=""padding-left:90px"" colspan=""2""><b> <font color=""navy"" size=""1"">Subtotal: " & lastAssigned  & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Submitted: "                          & subTotl       & " </td>"
         			   response.write "    <td colspan=""2""> <b> <font color=""navy"" size=""1""> Open: "                               & numOpen       & " </td>"
         			   response.write "    <td colspan=""2""><b><font color=""navy"" size=""1""> Avg Time Still Open: "    & avOpenTotal   & " </td>"
         			   response.write "    <td colspan=""3""><b><font color=""navy"" size=""1""> Avg Time To Complete: "   & avClosedTotal & " </td>"
               response.write "</tr>"
            end if
         end if
      end if
   end if

'----------------------------------------------------
   if reporttype = "ListFull" then
'----------------------------------------------------
		     Dim lcl_allowed_days, lcl_total_days, lcl_total_days_label
             if oRequests("usesafter5adjustment") = True then
			    if oRequests("usesweekdays") = True then
                   lcl_total_days = oRequests("totalDays_noweekends_adjusted")
				else
                   lcl_total_days = oRequests("totalDays_adjusted")
				end if
             else
			    if oRequests("usesweekdays") = True then
                   lcl_total_days = oRequests("totalDays_noweekends")
				else
                   lcl_total_days = oRequests("totalDays")
				end if
			 end if

		     if oRequests("allowedunresolveddays") > 0 then
                lcl_allowed_days = oRequests("allowedunresolveddays")
             else
			    lcl_allowed_days = 0
			 end if

			 if IsNULL(oRequests("complete_date")) then
				lcl_total_days_label = "Days Open"
			 else
				lcl_total_days_label = "Days taken to Complete"
			 end if

            'Build the USER City/State/Zip display value
             lcl_user_city  = oRequests("usercity")
             lcl_user_state = oRequests("userstate")
             lcl_user_zip   = oRequests("userzip")

             lcl_user_csz   = lcl_user_city

            'Add the userstate
             if lcl_user_csz = "" then
                lcl_user_csz = lcl_user_state
             else
                if lcl_user_state <> "" then
                   lcl_user_csz = lcl_user_csz & " / " & lcl_user_state
                else
                   lcl_user_csz = lcl_user_csz & " / -- "
                end if
             end if

            'Add the userzip
             if lcl_user_csz = "" then
                lcl_user_csz = lcl_user_zip
             else
                if lcl_user_zip <> "" then
                   lcl_user_csz = lcl_user_csz & " / " & lcl_user_zip
                end if
             end if

            'Build the ISSUE LOCATION City/State/Zip display value
             lcl_issue_city  = oRequests("city")
             lcl_issue_state = oRequests("state")
             lcl_issue_zip   = oRequests("zip")

             lcl_issue_csz   = lcl_issue_city

            'Add the issuestate
             if lcl_issue_csz = "" then
                lcl_issue_csz = lcl_issue_state
             else
                if lcl_issue_state <> "" then
                   lcl_issue_csz = lcl_issue_csz & " / " & lcl_issue_state
                else
                   lcl_issue_csz = lcl_issue_csz & " / -- "
                end if
             end if

            'Add the issuezip
             if lcl_issue_csz = "" then
                lcl_issue_csz = lcl_issue_zip
             else
                if lcl_issue_zip <> "" then
                   lcl_issue_csz = lcl_issue_csz & " / " & lcl_issue_zip
                end if
             end if

             response.write("<tr bgcolor=""" & bgcolor & """ id=""row_" & oRequests("action_autoid") & """ onMouseOver=""changeRowColor('row_" & oRequests("action_autoid") & "','OVER')"" onMouseOut=""changeRowColor('row_" & oRequests("action_autoid") & "','OUT')"" onClick=""location.href='action_respond.asp?control=" & oRequests("action_autoid") & "';"">")
             response.write("    <td class=""tablelist"">")
             response.write("        <table border=""0"" cellspacing=""0"" cellpadding=""5"">")
             response.write("          <tr><td><font size=""3""><b>" & lngTrackingNumber & "</b></font></td>")
             response.write("              <td><b>Submitted Date: </b>" & formatdatetime(datSubmitDate,vbShortDate)             & "</td>")
             response.write("              <td><b>Status: </b>"         & UCASE(sStatus)                                        & "</td>")

             if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
                response.write("           <td nowrap=""nowrap""><b>Sub-Status: </b>"     & sSubStatus                          & "</td>")
             end if

             response.write("              <td><b>Category: </b>"       & oRequests("action_FormTitle")                         & "</td></tr>")
             response.write("          <tr><td><b>Submitted By: </b>"   & oRequests("userfname") & " " & oRequests("userlname") & "</td>")


             response.write("              <td><b>Phone: </b>"        & FormatPhoneNumber(oRequests("userhomephone"))            & "</td>")
             response.write("              <td><b>Contact Address: </b>" & oRequests("useraddress")                              & "</td>")
             response.write("              <td colspan=""2""><b>Contact City/State/Zip: </b>" & lcl_user_csz                     & "</td></tr>")
             response.write("          <tr><td><b>Assigned To: </b>"  & oRequests("assigned_Name")                               & "</td>")
             response.write("              <td><b>Department: </b>"   & oRequests("deptName")                                    & "</td>")
             response.write("              <td><b>Days Allowed: </b>" & lcl_allowed_days                                         & " day(s)</td>")
             response.write("              <td><b>" & lcl_total_days_label    & ": </b>" & lcl_total_days                        & " day(s)</td></tr>")

             if OrgHasFeature("issue location") then
                sIssueName = oRequests("issuelocationname")
                If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
                   sIssueName = "Issue/Problem Location:"
                End If

                if oRequests("validstreet") <> "Y" AND oRequests("action_form_display_issue") then
                   lcl_valid_street = "<font style=""color: #FF0000;"">&nbsp;*</font>"
                else
                   lcl_valid_street = ""
                end if

             response.write("          <tr><td colspan=""2""><b>" & sIssueName & " </b>" & oRequests("streetname") & lcl_valid_street & "</td>")
             else
             response.write("          <tr><td colspan=""2"">&nbsp;</td>")
         			 end if
             response.write("              <td colspan=""2""><b>City/State/Zip: </b>" & lcl_issue_csz & "</td>")
             response.write("          </tr>")
             response.write("          <tr>")
             response.write("              <td colspan=""6""><b>Additional Info: </b>" & oRequests("comments") & "</td>")
             response.write("          </tr>")
             response.write("        </table>")

 '<p><b>Location of problem.</b><br>City Hall</p><p><b>Description of problem</b><br>I dumped a load of trash. Clean it up.</p>
			 
			 response.write("        <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">")

			'Format the egov_action_line_request.comment
             dim lcl_comment, arrQues, iQues, sQues, sAnswer, sValue
             sValue = oRequests("comment")

            'Format and split the questions and answers
             sValue = replace(sValue,"<B>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>")
             sValue = replace(sValue,"<b>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>")

	         lcl_comment = remove_html_tags(sValue)

			 response.write("          <tr><td><hr size=""1"" width=""100%""></td></tr>")
			 response.write("          <tr><td>")
			 response.write("                  <table border=""0"" cellspacing=""0"" cellpadding=""0"">")
			 response.write("                    <tr valign=""top"">")
			 response.write("                        <td>&nbsp;&nbsp;</td>")
			 response.write("                        <td>" & lcl_comment & "</td></tr>")
			 response.write("                  </table><p>")
			 response.write("              </td></tr>")
			 response.write("          <tr><td>")
											   List_Comments(oRequests("action_autoid"))
			 response.write("              </td></tr>")
             response.write("        </table>")
             response.write("    </td></tr>")


'----------------------------------------------------
   else 'if reporttype <> "ListFull" then
'----------------------------------------------------
      ''Response.Write "<tr bgcolor=" & bgcolor & " onClick=""location.href='action_respond.asp?control=" & oRequests("action_autoid") & "';"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td><b>(" & lngTrackingNumber & ") " & sTitle & " </b></td><td align=""center""> " & datSubmitDate & "</td><td align=""center""> " & oRequests("firstname")  & " " & oRequests("lastname") & "</td><td align=""center""> " & UCASE(sStatus) & "</td><td align=right><a href=""action_respond.asp?control=" & oRequests("action_autoid") & """>Review/Respond</a></tr>"
      Response.Write "<tr bgcolor=""" & bgcolor & """ onClick=""location.href='action_respond.asp?control=" & oRequests("action_autoid") & "';"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">"
      response.write "    <td><b>(" & lngTrackingNumber & ") " & sTitle & " </b></td>"
      response.write "    <td align=""center""> " & formatdatetime(datSubmitDate,vbShortDate) & "</td>"

      if reporttype="Detail" or reporttype="DrillThru" then
         Response.Write "<td align=""center"">" & datResolveDate & "</td><td align=""center"">"
         if oRequests("totalDays")<> "" then
            Response.Write oRequests("totalDays") & " days</td>"
            countDays = clng(oRequests("totalDays"))
         else
            openDays = dateDiff("d",oRequests("submit_date"),Now)
            Response.Write openDays & "<font color=""red"">* </td>"
            countDays = clng(openDays)
         end if
      end if

      Response.Write "<td align=""center"">" & UCASE(sStatus)             & "</td>"

   if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
      response.write "<td align=""center"" nowrap=""nowrap"">" & sSubStatus & "</td>"
   end if

   response.write "<td align=""center"">" & oRequests("userfname")     & " " & oRequests("userlname") & "</td>"
   response.write "<td>"                  & oRequests("useraddress")   & "</td>"
   response.write "<td align=""right"">"  & oRequests("assigned_Name") & "</td>"
   response.write "<td align=""center"">" & oRequests("deptName")      & "</td>"

   if OrgHasFeature("issue location") then
      if oRequests("validstreet") <> "Y" AND oRequests("action_form_display_issue") then
         lcl_valid_street = "<font style=""color: #FF0000;"">&nbsp;*</font>"
      else
         lcl_valid_street = ""
      end if

      response.write("<td>" & oRequests("streetname") & lcl_valid_street & "</td>")
   end if
	  
	  response.write "</tr>"

      if reporttype="Detail" or reporttype="DrillThru" then
         if orderBy = "submit_date" then
            if DateDiff("d",sDate,lastDate) = 0 and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
               subTotl   = subTotl + 1
               openTotl  = openTotl + 1
               totalDays = totalDays + countDays
            elseif DateDiff("d",sDate,lastDate) = 0 and UCASE(sStatus) <> "RESOLVED" then
               subTotl = subTotl + 1
               totalDays = totalDays + countDays											
            else
               if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
                  openTotl  = 1
                  subTotl   = 1
                  totalDays = countDays
               else
                  openTotl  = 0
                  subTotl   = 1
                  totalDays = countDays
               end if
            end if
            lastDate = sDate
         elseif orderBy = "action_Formid" then
            if sTitle = lastTitle and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
               subTotl   = subTotl + 1
               openTotl  = openTotl + 1
               totalDays = totalDays + countDays
            elseif sTitle = lastTitle then
               subTotl   = subTotl + 1
               totalDays = totalDays + countDays
            else
               if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
                  openTotl  = 1
                  subTotl   = 1
                  totalDays = countDays
               else
                  openTotl  = 0
                  subTotl   = 1
                  totalDays = countDays
               end if
            end if
            lastTitle = sTitle
         elseif orderBy = "deptId" then
            if sDept = lastDept and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
               subTotl   = subTotl + 1
               openTotl  = openTotl + 1
               totalDays = totalDays + countDays
            elseif sDept = lastDept then
               subTotl   = subTotl + 1
               totalDays = totalDays + countDays
            else
               if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
                  openTotl  = 1
                  subTotl   = 1
                  totalDays = countDays
               else
                  openTotl  = 0
                  subTotl   = 1
                  totalDays = countDays
               end if
            end if
            lastDept     = sDept
            lastDeptName = sDeptName
         elseif orderBy = "assigned_Name" then
            if sAssigned  = lastAssigned and (UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED") then
               subTotl   = subTotl + 1
               openTotl  = openTotl + 1
               totalDays = totalDays + countDays
            elseif sAssigned  = lastAssigned then
               subTotl   = subTotl + 1
               totalDays = totalDays + countDays
            else
               if UCASE(sStatus) <> "RESOLVED" AND UCASE(sStatus) <> "DISMISSED" then
                  openTotl  = 1
                  subTotl   = 1
                  totalDays = countDays
               else
                  openTotl  = 0
                  subTotl   = 1
                  totalDays = countDays
               end if
            end if
            lastAssigned  = sAssigned 
         end if
      end if
   end if

   oRequests.MoveNext

Loop

   if reporttype="Detail" or reporttype="DrillThru" then
      Response.Write "<tr bgcolor=""#dddddd"">"
  	   response.write "   <td style=""padding-left:90px"" colspan=""2""><b> <font color=""navy"" size=""1"">Subtotal: " & displayLastTitle     & " </td>"
   	  response.write "   <td colspan=""2""> <b> <font color=""navy"" size=""1""> Submitted: "                          & subTotl              & " </td>"
   	  response.write "   <td colspan=""2""> <b> <font color=""navy"" size=""1""> Open: "                               & openTotl             & " </td>"
   	  response.write "   <td colspan=""2""><b><font color=""navy"" size=""1""> Avg Time Still Open: "    & displayLastAvgOpen   & " </td>"
   	  response.write "   <td colspan=""3""><b><font color=""navy"" size=""1""> Avg Time To Complete: "   & displayLastAvgClosed & " </td>"
      response.write "</tr>"
   End If

   Response.Write "</table>"

else
   Response.write "<p><b>No records found</p>"
end if

End Function

'--------------------------------------------------------------------------------------------------
' FUNCTION GETGROUPS(IUSERID)
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

function remove_html_tags(p_value)
'<p>
  p_value = replace(p_value,"<P>","")
  p_value = replace(p_value,"</P>",vbcrlf & vbcrlf)
  p_value = replace(p_value,"<p>","")
  p_value = replace(p_value,"</p>",vbcrlf & vbcrlf)

'<br>
  p_value = replace(p_value,"<BR>",vbcrlf)
  p_value = replace(p_value,"</BR>",vbcrlf)
  p_value = replace(p_value,"<br>",vbcrlf)
  p_value = replace(p_value,"</br>",vbcrlf)
  p_value = replace(p_value,"<BR />",vbcrlf)
  p_value = replace(p_value,"</BR />",vbcrlf)
  p_value = replace(p_value,"<br />",vbcrlf)
  p_value = replace(p_value,"</br />",vbcrlf)

  remove_html_tags = p_value

end function

Function List_Comments(iID)

	sSQL = "SELECT *, es.status_name AS sub_status_name "
	sSQL = sSQL & " FROM egov_action_responses egr "
	sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egr.action_userid = egov_users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN users ON egr.action_userid = users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_actionline_requests_statuses AS es "
	sSQL = sSQL &               "ON egr.action_sub_status_id = es.action_status_id "
	sSQL = sSQL & " WHERE egr.action_autoid = " & iID
	sSQL = sSQL & " ORDER BY egr.action_editdate DESC"

	Set oCommentList = Server.CreateObject("ADODB.Recordset")
	oCommentList.Open sSQL, Application("DSN"), 3, 1
    sBGColor = "#E0E0E0"

'	sSQL = "SELECT * FROM egov_action_responses LEFT OUTER JOIN egov_users ON egov_action_responses.action_userid = egov_users.userid LEFT OUTER JOIN users on egov_action_responses.action_userid=users.userid where action_autoid=" & iID & " ORDER BY action_editdate DESC"
'	Set oCommentList = Server.CreateObject("ADODB.Recordset")
'	oCommentList.Open sSQL, Application("DSN"), 3, 1
'    sBGColor = "#E0E0E0"

'	Response.Write "<div style=""border-bottom:solid 1px #000000;""></div>"

	If NOT oCommentList.EOF Then
		Do While NOT oCommentList.EOF 

		   lcl_substatus_name = oCommentList("sub_status_name")
		   
		   if lcl_substatus_name <> "" then
			  lcl_substatus_name = " <i>(" & lcl_substatus_name & ")</i>"
		   end if

			Response.Write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>"
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">"
			response.write "  <tr><td>" & oCommentList("firstname") & " " & oCommentList("lastname") & " - " & UCASE(oCommentList("action_status")) & lcl_substatus_name & " - " &  oCommentList("action_editdate") & ")<br>"
			
			If oCommentList("action_externalcomment") <> "" Then
				response.write "&nbsp;&nbsp;&nbsp;<b>Note to Citizen: </b><i>" & oCommentList("action_externalcomment")  & "</i></td></tr>" 
			End If

			If oCommentList("action_citizen") <> "" Then
				response.write "&nbsp;&nbsp;&nbsp;<b>" & oCommentList("userfname")  & " " & oCommentList("userlname") & " : </b><i>" & oCommentList("action_citizen") 
			End If

			If oCommentList("action_internalcomment") <> "" Then
				response.write "&nbsp;&nbsp;&nbsp;<b>Internal Note: </b><i>" & oCommentList("action_internalcomment")  & "</i></td></tr>" 
			End If
			Response.Write "</table></div>"

			oCommentList.MoveNext

			If sBGColor = "#FFFFFF" Then
				sBGColor = "#E0E0E0"
			Else
				sBGColor = "#FFFFFF"
			End If
		Loop

			' DISPLAY SUBMIT DATE TIME AND USER
			Response.Write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & ";"">"
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">"
			response.write "  <tr><td>" & sSubmitName & " - " & UCASE("SUBMITTED") & " - " &  datSubmitDate & "</td></tr>"
			response.write "</table></div>"

	Else

		' NO ACTIVITY FOR THIS REQUEST
		Response.Write "<div style=""border-bottom:solid 1px #000000;background-color:#e0e0e0"">"
		response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">"
		response.write "  <tr><td><font style=""color:red;font-size:12px;"">&nbsp;&nbsp;&nbsp;<i>No activity Reported.</i></td></tr>"
		response.write "</table></div>"


		' DISPLAY SUBMIT DATE TIME AND USER
		Response.Write "<div style=""border-bottom:solid 1px #000000;background-color:#FFFFFF;"">"
		response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">"
		response.write "  <tr><td>" & sSubmitName & " - " & UCASE("SUBMITTED") & " - " &  datSubmitDate & "</td></tr>"
		response.write "</table></div>"
		
	End If
end function
%>

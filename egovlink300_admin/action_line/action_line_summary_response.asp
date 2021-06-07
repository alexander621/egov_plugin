<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "requests" ) Then
  	response.redirect sLevel & "permissiondenied.asp"
End If 

'GET USER'S PERMISSIONS
 'blnCanViewAllActionItems = HasPermission("CanViewAllActionItems")
 'blnCanViewOwnActionItems = HasPermission("CanViewOwnActionItems")
 'blnCanViewDeptActionItems = HasPermission("CanViewDeptActionItems")

iPermissionLevelId = GetUserPermissionLevel( Session("UserId"), "requests" )
If clng(iPermissionLevelId) > 0 Then 
  	sPermissionLevel = GetPermissionLevelName( iPermissionLevelId )
Else
  	response.redirect sLevel & "permissiondenied.asp"
End If 

'Set to use new permission levels
 blnCanViewAllActionItems  = False  
 blnCanViewOwnActionItems  = False 
 blnCanViewDeptActionItems = False

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
			reporttype          = session("reporttype")
			
			orderBy             = session("orderBy")
			selectFormId        = session("selectFormId")
			selectAssignedto    = session("selectAssignedto")
			selectDeptId        = session("selectDeptId")
			pastDays            = session("pastDays")
			
			selectUserFName     = session("selectUserFName")
			selectUserLName     = session("selectUserLName")
					
			fromDate            = session("fromDate")
			toDate              = session("toDate")
   selectDateType      = session("selectDateType")
			today               = Date()
			
			statusSubmitted     = session("statusSubmitted")
			statusInprogress    = session("statusInprogress")
			statusWaiting       = session("statusWaiting")
			statusResolved      = session("statusResolved")
			statusDismissed     = session("statusDismissed")

			substatus_hidden    = session("substatus_hidden")
			show_hide_substatus = session("show_hide_substatus")

else
			recordsPer          = request("recordsPer")
			reporttype          = request("reporttype")
			
			orderBy             = request("orderBy")
			selectFormId        = request("selectFormId")
			selectAssignedto    = request("selectAssignedto")
			selectDeptId        = request("selectDeptId")
			pastDays            = request("pastDays")
			
			selectUserFName     = request("selectUserFName")
			selectUserLName     = request("selectUserLName")
			
			fromDate            = request("fromDate")
			toDate              = request("toDate")
   selectDateType      = request("selectDateType")
			today               = Date()

   statusSubmitted     = request("statusSUBMITTED")
   statusInprogress    = request("statusINPROGRESS")
   statusWaiting       = request("statusWAITING")
   statusResolved      = request("statusRESOLVED")
   statusDismissed     = request("statusDISMISSED")

   substatus_hidden    = request("substatus_hidden")
   show_hide_substatus = request("show_hide_substatus")
end if

If reporttype = "" or IsNull(reporttype) Then reporttype = "List" End If

''If orderBy = "" or IsNull(orderBy) Then orderBy = " submit_date" End If
If orderBy = "" or IsNull(orderBy) Then 
  	orderBy = request("orderBy")
  	If orderBy = "" or IsNull(orderBy) Then
    		orderBy = "action_Formid" 
  	End If
end if

If selectFormId     = "" or IsNull(selectFormId) Then selectFormId = "all" End If
If selectAssignedto = "" or IsNull(selectAssignedto) Then selectAssignedto = "all" End If
'If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '" & selectAssignedto & "'" 
If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assignedemployeeid = " & selectAssignedto

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
if request("init") = "Y" OR (request("init") = "" AND statusSubmitted = "no" AND statusInprogress = "no" AND statusWaiting = "no" AND statusResolved = "no" AND statusDismissed = "no" AND substatus_hidden = "") then
   lcl_init = "Y"
else
   lcl_init = "N"
end if

if lcl_init = "Y" then
   session("isFromEmail") = ""

   statusSubmitted  = "yes"
   statusInprogress = "yes"
   statusWaiting    = "yes"
   statusResolved   = "yes"
   statusDismissed  = "yes"
end if

session("reporttype")          = reporttype
session("orderBy")             = orderBy
session("selectFormId")        = selectFormId
session("selectAssignedto")    = selectAssignedto 
session("selectDeptId")        = selectDeptId

session("toDate")              = toDate 
session("fromDate")            = fromDate
session("selectDateType")      = selectDateType
session("recordsPer")          = recordsPer

session("noStatus")            = noStatus
session("statusDismissed")     = statusDismissed
session("statusResolved")      = statusResolved
session("statusWaiting")       = statusWaiting
session("statusInprogress")    = statusInprogress
session("statusSubmitted")     = statusSubmitted

session("substatus_hidden")    = substatus_hidden
session("show_hide_substatus") = show_hide_substatus

session("selectUserFName")     = selectUserFName
session("selectUserLName")     = selectUserLName
%>

<html>
<head>
  <title><%=langBSActionLine%></title>
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  
<SCRIPT LANGUAGE="JavaScript">
  function checkStat() {
//  if ( !(form1.statusSubmitted.checked) &&  !(form1.statusInprogress.checked) && !(form1.statusWaiting.checked) && !(form1.statusResolved.checked) && !(form1.statusDismissed.checked)) {
  if ( !(form1.statusSUBMITTED.checked) &&  !(form1.statusINPROGRESS.checked) && !(form1.statusWAITING.checked) && !(form1.statusRESOLVED.checked) && !(form1.statusDISMISSED.checked)) {
		alert("You must select the status.");

//		form1.statusSubmitted.focus();
		form1.statusSUBMITTED.focus();
		return false;
	}
  }
  function CheckAllStatus(checkSt) {
		//if (document.form1.CheckAllStat.checked) {
		if (checkSt) {
			document.form1.statusSUBMITTED.checked  = true;
			document.form1.statusINPROGRESS.checked = true;
			document.form1.statusWAITING.checked    = true;
			document.form1.statusRESOLVED.checked   = true;
			document.form1.statusDISMISSED.checked  = true;
		} else {
			document.form1.statusSUBMITTED.checked  = false;
			document.form1.statusINPROGRESS.checked = false;
			document.form1.statusWAITING.checked    = false;
			document.form1.statusRESOLVED.checked   = false;
			document.form1.statusDISMISSED.checked  = false;
		}
  }
 
 function submitForm(){
		 if (document.form1.reporttype.value == "Summary") {
				document.forms[0].action = "action_line_summary.asp";
				document.forms[0].submit();
		} 
		else if (document.form1.reporttype.value == "ResponseSummary") {
				document.forms[0].action = "action_line_summary_response.asp";
				document.forms[0].submit();
			}
		else if (document.form1.reporttype.value == "responsedetail") {
				document.forms[0].action = "action_line_list_response.asp";
				document.forms[0].submit();
			}
		else if (document.form1.reporttype.value == "ListFull") {
				document.forms[0].action = "action_line_list.asp";
				document.forms[0].submit();
            }
		else {
				document.forms[0].action = "action_line_list.asp"
				document.forms[0].submit();
			}
}

function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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

function changeSubStatus() {
  var list
  var list2
  var i
  var a

  mainlist          = document.getElementById('selStatus');
  sub_list          = document.getElementById('selSubStatus');
  sub_list_row      = document.getElementById('sub_status_row');
  sub_list_row_text = document.getElementById('sub_status_row_text');
  i = 0
<%
  dim oMainStatus, oSubStatus, oSubStatus_Count, line_count, lcl_sub_line_count, lcl_total_count

 'Retrieve all of the MAIN statuses
  sSqlm = "SELECT action_status_id, status_name, orgid, parent_status, display_order, active_flag "
  sSqlm = sSqlm & " FROM egov_actionline_requests_statuses "
  sSqlm = sSqlm & " WHERE orgid = 0 "
  sSqlm = sSqlm & " AND parent_status = 'MAIN' "
  sSqlm = sSqlm & " AND active_flag = 'Y' "
  sSqlm = sSqlm & " ORDER BY display_order "

  Set oMainStatus = Server.CreateObject("ADODB.Recordset")
  oMainStatus.Open sSqlm, Application("DSN"), 3, 1

  If NOT oMainStatus.EOF Then
     line_count = 0
	 do while NOT oMainStatus.EOF
        line_count = line_count + 1

		if line_count = 1 then
%>
           if(document.getElementById('selStatus<%=oMainStatus("status_name")%>').checked==true) {
<%      else %>
           }else if(document.getElementById('selStatus<%=oMainStatus("status_name")%>').checked==true) {
<%
        end if

	   'Get the total count of SubStatuses
        sSqlc = "SELECT count(action_status_id) AS Total_SubStatus FROM egov_actionline_requests_statuses "
        sSqlc = sSqlc & " WHERE orgid = "         & clng(Session("OrgID"))
        sSqlc = sSqlc & " AND parent_status = '"  & oMainStatus("status_name") & "' "
        sSqlc = sSqlc & " AND active_flag = 'Y' "
        Set oSubStatus_Count = Server.CreateObject("ADODB.Recordset")
        oSubStatus_Count.Open sSqlc, Application("DSN"), 3, 1

        lcl_total_count = oSubStatus_Count("Total_SubStatus")

        if lcl_total_count > 0 then
		  
		  'Retrieve all of the Sub-Statuses for each MAIN status for the OrgID and the form
           sSqls = "SELECT action_status_id, status_name "
           sSqls = sSqls & " FROM egov_actionline_requests_statuses "
           sSqls = sSqls & " WHERE orgid = "         & clng(Session("OrgID"))
           sSqls = sSqls & " AND parent_status = '"  & oMainStatus("status_name") & "' "
           sSqls = sSqls & " AND active_flag = 'Y' "
           sSqls = sSqls & " ORDER BY display_order "

           Set oSubStatus = Server.CreateObject("ADODB.Recordset")
           oSubStatus.Open sSqls, Application("DSN"), 3, 1

           If NOT oSubStatus.EOF Then
%>
              sub_list_row.style.display      = "block";
			  sub_list.style.display          = "block";

            //remove the current values
			  for(var i=0; i < sub_list.length; i++) {
                  sub_list.remove(i);
			  }
<%
             'Loop through the sub statuses
              lcl_sub_line_count = 0
			  do while NOT oSubStatus.EOF
%>
            //build the new values
				  document.forms["form1"].selSubStatus.options[<%=lcl_sub_line_count%>] = new Option("<%=oSubStatus("status_name")%>","<%=oSubStatus("action_status_id")%>");
<%
                 lcl_sub_line_count = lcl_sub_line_count + 1
				 oSubStatus.movenext
              loop

			  oSubStatus.Close
			  oSubStatus_Count.Close

			  Set oSubStatus       = Nothing
			  Set oSubStatus_Count = Nothing 
		   
		   else
%>
              sub_list_row.style.display      = "none";
			  sub_list.style.display          = "none";
<%
           end if
        else
%>
              sub_list_row.style.display      = "none";
			  sub_list.style.display          = "none";
<%
		end if

		oMainStatus.movenext
	 loop
%>
           }else{
              sub_list_row.style.display      = "none";
			  sub_list.style.display          = "none";
		   }
<%
  else
  end if

  oMainStatus.Close
  Set oMainStatus   = Nothing 
%>
}

function change_substatus_filter() {
  var lcl_substatus_display = document.getElementById('display_substatus');
  var lcl_substatus_text    = document.getElementById('substatus_hidden');
  var lcl_substatus_value;
<%
 'Get a total count of all of the active sub-statuses for this org.
  sSqlc = "SELECT count(s1.action_status_id) AS total_count "
  sSqlc = sSqlc & " FROM egov_actionline_requests_statuses s1 "
  sSqlc = sSqlc & " WHERE s1.active_flag = 'Y' "
  sSqlc = sSqlc & " AND s1.orgid = " & session("orgid")

  Set oTotal = Server.CreateObject("ADODB.Recordset")
  oTotal.Open sSqlc, Application("DSN"), 3, 1

  lcl_total_substatuses = oTotal("total_count")

  set oTotal = nothing

 '1. Build the javascript that will cycle through all of the sub-status search criteria checkboxes and determine which ones
 'have been checked.
 '2. Clear the hidden field (element_id = substatus_hidden) and the display field (element_id = display_substatus) that hold the values.
 '   the hidden field is used for form and query processing and display field is to show the user which values he/she has selected if the
 '   substatus list has been collapsed.
 '3. For those that have been checked then rebuild the hidden field by cycling through all of the checkboxes.
 '4. This will also rebuild the list when a value unchecked.

 'Get all of the active sub-statuses for this org.
  sSqla = "SELECT s1.action_status_id, s1.status_name, s1.display_order, s1.parent_status, s2.display_order AS parent_display_order "
  sSqla = sSqla & " FROM egov_actionline_requests_statuses s1, egov_actionline_requests_statuses s2 "
  sSqla = sSqla & " WHERE s1.parent_status = s2.status_name "
  sSqla = sSqla & " AND s1.active_flag = 'Y' "
  sSqla = sSqla & " AND s2.active_flag = 'Y' "
  sSqla = sSqla & " AND s1.orgid = " & session("orgid")
  sSqla = sSqla & " ORDER BY 5, 4, 3, 2 "

  Set oChange = Server.CreateObject("ADODB.Recordset")
  oChange.Open sSqla, Application("DSN"), 3, 1

%>
     for(var i=0; i < <%=lcl_total_substatuses%>; i++) {
         lcl_substatus_display.innerHTML  = "";
         lcl_substatus_text.value         = "";
<%
  if not oChange.eof then
     do while not oChange.eof
%>
	     lcl_substatus_value = document.getElementById('SS_<%=oChange("action_status_id")%>');
		 if(lcl_substatus_value.checked==true) {
            if(lcl_substatus_display.innerHTML == "") {
	           lcl_substatus_display.innerHTML = '<%=oChange("status_name")%>';
   	 	       lcl_substatus_text.value        = "("+ lcl_substatus_value.value +")";
	        }else{
               lcl_substatus_display.innerHTML = lcl_substatus_display.innerHTML + ", " + '<%=oChange("status_name")%>';
               lcl_substatus_text.value        = lcl_substatus_text.value        + ", " + "("+ lcl_substatus_value.value +")";
            }
         }
<%
        oChange.movenext
	 loop
  end if

  set oChange = nothing
%>
     }	 
}

function showhide_substatus_criteria() {
  var lcl_button = document.getElementById('selectSubStatus');

  if(lcl_button.style.display == "block") {
     lcl_button.style.display = "none";
     document.getElementById('show_hide_substatus').value = "HIDE";
  }else{
     lcl_button.style.display = "block";
     document.getElementById('show_hide_substatus').value = "SHOW";
  }
}

function show_hide_init(p_value) {
  var lcl_button = document.getElementById('selectSubStatus');

  if(p_value=="SHOW") {
     lcl_button.style.display = "block";
     document.getElementById("show_hide_substatus").value = "SHOW";
  }else if(p_value=="HIDE") {
     lcl_button.style.display = "none";
     document.getElementById("show_hide_substatus").value = "HIDE";
  }
}
</script>

</head>

<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="javascript:document.getElementById('selectSubStatus').style.display='none';<%=lcl_display_substatuses%>"> -->
<%
lcl_onload = ""

'If any substatuses have been checked and posted after clicking the Search button then we need to show the 'display' list
if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
   if substatus_hidden <> "" then
      lcl_display_substatuses = "change_substatus_filter();"
   else
      lcl_display_substatuses = ""
   end if

   lcl_onload = "document.getElementById('selectSubStatus').style.display='none';" & lcl_display_substatuses

  'Initialize the sub-status show/hide
   if show_hide_substatus = "" then
      show_hide_substatus = "HIDE"
   else
      show_hide_substatus = show_hide_substatus
   end if

   lcl_onload = lcl_onload & "show_hide_init('" & show_hide_substatus & "');"

else
   lcl_onload = ""
end if

if lcl_onload <> "" then
   lcl_onload = "onload=""javascript:" & lcl_onload & """ "
else
   lcl_onload = ""
end if
%>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" <%=lcl_onload%>>

<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr><td><font size="+1"><b>(E-Gov Request Manager) - Manage Action Line Requests</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td></tr>
  <tr><td>
          <fieldset>
            <legend><b>Search/Sorting Option(s)</b></legend>
			<form name="form1" onSubmit="return checkStat()">
		  <table border="0" bordercolor="red">
      <tr>
		        <td valign="top" nowrap="nowrap">
              <b>Assigned To: <% if ViewAll=0 then response.write "(User " & session("userID") & ")" %> 
     	        <select name="selectAssignedto">
                <option value="all">Anyone</option> 
             			<%
                  sSQLassignedto = "SELECT FirstName + ' ' + LastName as assigned_Name, UserID "
                  sSQLassignedto = sSQLassignedto & " FROM USERS "
                  sSQLassignedto = sSQLassignedto & " WHERE OrgID = " & Session("OrgID")
                  sSQLassignedto = sSQLassignedto & " ORDER BY FirstName,LastName"

                  Set oAssigned = Server.CreateObject("ADODB.Recordset")
                  oAssigned.Open sSQLassignedto, Application("DSN"), 3, 1

                  do while not oAssigned.EOF		
'                   		if selectAssignedto = oAssigned("assigned_Name") then 
                     if selectAssignedto <> "all" then
                      		if clng(selectAssignedto) = clng(oAssigned("userid")) then 
                      					selectAssign = "selected"
                    				else
                      					selectAssign = ""
                    				end if
                     else
                        selectAssign = ""
                     end if

'                  			response.write "<option value=""" & oAssigned("assigned_Name") & """ " & selectAssign & ">" & oAssigned("assigned_Name") & "</option>" & vbcrlf
                  			response.write "<option value=""" & oAssigned("userid") & """ " & selectAssign & ">" & oAssigned("assigned_Name") & "</option>" & vbcrlf
                  			oAssigned.MoveNext
               			Loop
             			%>
              </select>&nbsp;&nbsp;&nbsp; 
	          		 <b>Group By: 
            		<select name="orderBy">
			<% if orderBy = "submit_date" then select2 = "selected" else select2="" end if %>
           					<option value="submit_date" <%=select2%>>Date</option>
			<% if orderBy = "action_Formid" then select2 = "selected" else select2="" end if %>
           					<option value="action_Formid" <%=select2%>>Category</option>
			<% if orderBy = "deptId" then select2 = "selected" else select2="" end if %>
           					<option value="deptId" <%=select2%>>Department</option>
			<% if orderBy = "assigned_Name" then select2 = "selected" else select2="" end if %>
           					<option value="assigned_Name" <%=select2%>>Assigned To</option>				
 <% If OrgHasFeature("issue location") Then %>
				<% if orderBy = "streetname" then select2 = " selected=""selected"" " else select2="" end if %>
                <option value="streetname" <%=select2%>>Issue/Problem Location Street Name</option>
 <% end if %>
			<% if orderBy = "submittedby" then select2 = " selected=""selected"" " else select2="" end if %>
                <option value="submittedby" <%=select2%>>Submitted By</option>
           	  </select>
   		  </td>
		  </tr>
		<tr>
			<td valign="top" nowrap>
			<%
  			if statusSubmitted  = "yes" then check1 = " checked=""checked"" "
		  	if statusInprogress = "yes" then check2 = " checked=""checked"" " 
  			if statusWaiting    = "yes" then check3 = " checked=""checked"" "
		  	if statusResolved   = "yes" then check4 = " checked=""checked"" "
  			if statusDismissed  = "yes" then check5 = " checked=""checked"" "
			%>
			<b>Status:</b> 

                    <input type="checkbox" name="statusSUBMITTED" id="selStatusSUBMITTED" value="yes" <%=check1%> />Submitted <!--onClick="changeSubStatus()" />Submitted-->
                    <input type="checkbox" name="statusINPROGRESS" id="selStatusINPROGRESS" value="yes" <%=check2%> />In Progress <!--onClick="changeSubStatus()" />In Progress-->
                    <input type="checkbox" name="statusWAITING" id="selStatusWAITING" value="yes" <%=check3%> />Waiting <!--onClick="changeSubStatus()" />Waiting-->
                    <input type="checkbox" name="statusRESOLVED" id="selStatusRESOLVED" value="yes" <%=check4%> />Resolved <!--onClick="changeSubStatus()" />Resolved-->
                    <input type="checkbox" name="statusDISMISSED" id="selStatusDISMISSED" value="yes" <%=check5%> />Dismissed <!--onClick="changeSubStatus()" />Dismissed-->
                 <%
                    if UserHasPermission( Session("UserId"), "action_line_substatus" ) then

                   'Cycle through each main status and determine if there are any active sub-statuses.
				               'Retrieve all of the sub-statuses for the organization for each parent_status
                    sSql1 = "SELECT s1.action_status_id, s1.status_name, s1.orgid, s1.parent_status, s1.display_order, s1.active_flag, "
               					sSql1 = sSql1 & " s2.action_status_id AS parent_status_id, s2.display_order AS parent_display_order "
               					sSql1 = sSql1 & " FROM egov_actionline_requests_statuses s1, egov_actionline_requests_statuses s2"
               					sSql1 = sSql1 & " WHERE s1.parent_status = s2.status_name "
               					sSql1 = sSql1 & " AND s1.active_flag = 'Y' "
               					sSql1 = sSql1 & " AND s2.active_flag = 'Y' "
               					sSql1 = sSql1 & " AND s1.orgid = " & session("orgid")
               					sSql1 = sSql1 & " ORDER BY 8, s1.parent_status, s1.display_order, s1.status_name "

               					Set oExists = Server.CreateObject("ADODB.Recordset")
                    oExists.Open sSql1, Application("DSN"), 3, 1

                    i = 0
               					if not oExists.eof then
                 %>
               					<table border="0" cellspacing="0" cellpadding="2" id="sub_status_row">
                      <tr><td>&nbsp;&nbsp;&nbsp;<strong>Sub-Status:&nbsp;</strong>
                              <input type="hidden" name="substatus_hidden" id="substatus_hidden" value="<%=substatus_hidden%>">
                              <input type="hidden" name="show_hide_substatus" id="show_hide_substatus" value="<%=show_hide_substatus%>">
                              <span id="display_substatus"></span><p>
                              <input type="button" value="Show/Hide Sub-Status List" style="cursor: hand" onclick="javascript:showhide_substatus_criteria();"></td></tr>
                      <tr valign="top">
					      <td>
							  <span id="selectSubStatus">
                                <table border="0" cellspacing="1" cellpadding="0" bgcolor="#000000">
								  <tr><td>
								<table border="0" cellspacing="0" cellpadding="2" bgcolor="#c0c0c0">
								  <tr valign="top">
                 <%
                       lcl_parent_status = ""
					   lcl_line_count    = 0

					  'Loop through all of the Sub-Statuses
					   while not oExists.eof
                          lcl_line_count = lcl_line_count + 1
						  i = i + 1

                          if instr(substatus_hidden,"(" & oExists("action_status_id") & ")") > 0 then
                            'lcl_click = "document.getElementById('SS_" & oExists("action_status_id") & "').checked;"
                             lcl_click = " checked=""checked"" "
                          else
                             lcl_click = ""
						  end if

                          if lcl_line_count = 1 then
						     lcl_parent_status = oExists("parent_status")
                 %>
									  <td>
                                          <table border="0" cellspacing="0" cellpadding="2" bgcolor="#ffffff">
                                            <tr bgcolor="#efefef">
											    <td colspan="2"><b><%=oExists("parent_status")%></b></td></tr>
						                	<tr><td><input type="checkbox" name="SS_<%=oExists("action_status_id")%>" id="SS_<%=oExists("action_status_id")%>" value="<%=oExists("action_status_id")%>" onclick="javascript:change_substatus_filter();"<%=lcl_click%>></td>
          								        <td><%=oExists("status_name")%></td></tr>
                 <%
                          else
   						     if UCASE(lcl_parent_status) <> UCASE(oExists("parent_status")) then
							    if lcl_line_count > 1 then
								   lcl_line_count = 1
                                   lcl_parent_status = oExists("parent_status")
				 %>
                                          </table>
						              </td>
									  <td>
                                          <table border="0" cellspacing="0" cellpadding="2" bgcolor="#ffffff">
                                            <tr bgcolor="#efefef">
											    <td colspan="2"><b><%=oExists("parent_status")%></b></td></tr>
						                	<tr><td><input type="checkbox" name="SS_<%=oExists("action_status_id")%>" id="SS_<%=oExists("action_status_id")%>" value="<%=oExists("action_status_id")%>" onclick="javascript:change_substatus_filter();"<%=lcl_click%>></td>
          								        <td><%=oExists("status_name")%></td></tr>
                 <%
				                end if
                             else
                                lcl_parent_status = lcl_parent_status
                 %>
						                	<tr><td><input type="checkbox" name="SS_<%=oExists("action_status_id")%>" id="SS_<%=oExists("action_status_id")%>" value="<%=oExists("action_status_id")%>" onclick="javascript:change_substatus_filter();"<%=lcl_click%>></td>
          								        <td><%=oExists("status_name")%></td></tr>
                 <%
                             end if
                          end if
						  oExists.movenext
					   wend
                 %>
                                          </table>
						              </td>
								  </tr>
                                </table>
								      </td></tr>
								</table>
                                <p>
							  </span>
				 <% else %>
					<table border="0" cellspacing="0" cellpadding="2" id="sub_status_row">
                      <tr><td>&nbsp;&nbsp;&nbsp;<strong>Sub-Status:&nbsp;</strong>No Sub-Statuses
					          <span id="selectSubStatus"></span>
                              <input type="hidden" name="substatus_hidden" id="substatus_hidden" value="<%=substatus_hidden%>">
                              <input type="hidden" name="show_hide_substatus" id="show_hide_substatus" value="<%=show_hide_substatus%>"></td></tr>
				 <% end if %>
						  </td></tr>
					</table>
                 <% end if 'end of check for permission %>
<!------------------------------------------------------------------->
			</td></tr>
		  
		  <tr>
		  <td valign=top nowrap>
		    <b>Category: 
		    <select name="selectFormId"><option value="">All Categories</option><% fnListForms()%></select>
			</td>
      </tr>
      
      <tr>
		  <td valign=top nowrap>
		    <b>Department: 
		    <select name="selectDeptId">
        <option value="all">All Departments</option>
        <% fnListDepts()%>
      </select>
		    
		    &nbsp;&nbsp;&nbsp; 
		    <b>Report Type: 
			    <select name="reporttype">
              <% if reporttype = "Detail" then %>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					  <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail" selected>Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail" >Response Detail</option>
                   <% End If%>
			   <% elseif reporttype = "Summary" then %>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					  <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary" selected>Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail" >Response Detail</option>
                   <% End If%>
			   <% elseif reporttype = "ResponseSummary" Then%>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					  <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary" selected>Response Summary</option>
                      <option value="responsedetail" >Response Detail</option>
                   <% End If%>
               <% elseif reporttype = "responsedetail" Then%>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					  <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail" selected>Response Detail</option>
                   <% End If%>
			   <% elseif reporttype = "ListFull" then%>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					  <option value="ListFull" selected>List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail" >Response Detail</option>
                   <% End If%>
			   <% else %>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					  <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail" >Response Detail</option>
                   <% End If%>
			   <% end if %>
			  </select>
			</td>
      </tr>
      
      <tr>
		  <td valign="top" nowrap>
    		  <fieldset>      
          <b>From: 
   					  <input type=text name="fromDate" value="<%=fromDate%>">
			    	  <a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border="0"></a>		 
   				   &nbsp; 
			     		<b>To:</b> 
   					  <input type=text name="toDate" value="<%=dateAdd("d",-1,toDate)%>">
			   		  <a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border="0"></a>
          <b>From/To Dates will search on:</b>
          <select name="selectDateType">
          <%
            if UCASE(selectDateType) = "SUBMIT" then
               lcl_selected_active = ""
               lcl_selected_submit = " selected"
            else
               lcl_selected_active = " selected"
               lcl_selected_submit = ""
            end if
          %>
            <option value="active"<%=lcl_selected_active%>>Active Requests</option>
            <option value="submit"<%=lcl_selected_submit%>>Submit Date</option>
          </select>
        </fieldset>
				</td>
      </tr>	
      
      <tr>	
				<td valign=top nowrap>
      <b>Submitted By: &nbsp;&nbsp;
					  First: <input type=text name="selectUserFName" value="<% if selectUserFName <> "all" then response.write selectUserFName %>" size=12>
				   &nbsp; 
					<b>Last:</b> 
					  <input type=text name="selectUserLName" value="<% if selectUserLName <> "all" then response.write selectUserLName %>" size=12>
					  
				</td>
      </tr>	    
      
      
		  <tr>
		  <td>
		  
		  <b>Display Open Over: 
					  <% if pastDays <> "all" then %>					  
					  		<input type=text name="pastDays" value="<%=pastDays%>" size=2>  days
					  <% else %>
					  		<input type=text name="pastDays" value="" size=2> days 
					  <% end if %>
					  		 
				   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  
				  <input type="button" style="cursor: hand" onclick="javascript:submitForm();" value=" SEARCH ">
		  </td>
		    </tr>
			</table>
			</form>
			</fieldset>
			 </td>
    </tr>


    <tr>
      <td valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
      <form name=requestlist action=# method="POST">
		
		<%
		' DISPLAY REPORT TITLE
		Response.Write "<Br><font size=3 color=3399ff><i><b>RESPONSE SUMMARY REPORT</b></i></font>"
		
		' GET REPORT RESULTS
		List_Action_Requests(sSortBy)
		%>

	  </form>
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>

</body>
</html>


<%
' -------------------------------------------------------------------------------------------------
' BEGIN USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


' -------------------------------------------------------------------------------------------------
' FUNCTION LIST_ACTION_REQUESTS(SSORTBY)
'--------------------------------------------------------------------------------------------------
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

for u = 0 to ubound(statArray)
	varStatClause = varStatClause & "" & statArray(u)
next
lenStatClause = len(varStatClause) - 3
if lenStatClause > 1 then
   varStatClause = left(varStatClause,lenStatClause)
end if

'Check the selectDateType to determine how to use the From/To Date fields.
if UCASE(selectDateType) = "ACTIVE" then
   varWhereClause = " WHERE egov_action_request_view.orgid=('"&session("orgid")&"') AND ( "    ''IsNull(complete_date,'" & Now & "')
   varWhereClause = varWhereClause & " (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') OR "
   varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
   varWhereClause = varWhereClause & " (submit_date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "
else 'selectDateType = SUBMIT
   varWhereClause = " WHERE egov_action_request_view.orgid=" & session("orgid")
   varWhereClause = varWhereClause & " AND (submit_date BETWEEN '" & fromDate & "' AND '" & toDate & "'"
end if

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

	else
		varWhereClause = varWhereClause & " AND action_Formid = " & selectFormId
	end if
end if

'If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '" & selectAssignedto & "'"
If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assignedemployeeid = " & selectAssignedto
If selectDeptId     <> "all" then varWhereClause = varWhereClause & " AND deptID = "             & selectDeptId     & " "
If selectUserFName  <> "all" then varWhereClause = varWhereClause & " AND UserFName LIKE '"      & selectUserFName  & "%'"
If selectUserLName  <> "all" then varWhereClause = varWhereClause & " AND UserLName LIKE '"      & selectUserLName  & "%'"

'CREATE SQL BASED ON SELECTED GROUP BY
If orderBy="submit_date" then
 	'SQL FOR SUBMIT DATE
  	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, submitdateshort as TheDate, "
   sSQL = sSQL & " SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays,count(*) as numSubmitted, "
   sSQL = sSQL & " sum(viewedrequests) as numviewedrequests, sum(respondedrequests) as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &   varWhereClause
   sSQL = sSQL & " GROUP BY submitdateshort ORDER BY 3 desc "
elseif orderBy="action_Formid" then 
 	'SQL FOR ACTION FORM
  	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, action_formTitle, "
   sSQL = sSQL & " action_Formid, SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays, "
   sSQL = sSQL & " count(*) as numSubmitted, sum(viewedrequests) as numviewedrequests, sum(respondedrequests) as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &   varWhereClause
   sSQL = sSQL & " GROUP BY action_formTitle, action_Formid "
   sSQL = sSQL & " ORDER BY action_formTitle "
elseif orderBy="deptId" then 
 	'SQL FOR DEPARTMENT
  	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, deptID, "
   sSQL = sSQL & " SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays, count(*) as numSubmitted, "
   sSQL = sSQL & " sum(viewedrequests) as numviewedrequests, sum(respondedrequests) as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &   varWhereClause
   sSQL = sSQL & " GROUP BY deptID "
   sSQL = sSQL & " ORDER BY deptID "
elseif orderBy="assigned_Name" then 
 	'SQL FOR ASSIGNED NAME
  	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, assigned_Name, "
   sSQL = sSQL & " count(*) as numSubmitted, SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays, "
   sSQL = sSQL & " sum(viewedrequests) as numviewedrequests, sum(respondedrequests) as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &   varWhereClause
   sSQL = sSQL & " GROUP BY assigned_Name "
   sSQL = sSQL & " ORDER BY assigned_Name "
elseif orderBy="streetname" then
 	'SQL FOR STREET NAME
  	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, streetaddress, streetnumber, "
   sSQL = sSQL & " SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays, count(*) as numSubmitted, "
   sSQL = sSQL & " sum(viewedrequests) as numviewedrequests, sum(respondedrequests) as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &   varWhereClause
   sSQL = sSQL & " GROUP BY streetaddress, streetnumber "
   sSQL = sSQL & " ORDER BY UPPER(streetaddress), CAST(streetnumber AS int) "
elseif orderBy="submittedby" then
  'SQL FOR SUBMITTED BY
  	sSQL = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, "
   sSQL = sSQL & " userlname, userfname, count(*) as numSubmitted, "
   sSQL = sSQL & " SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays, "
   sSQL = sSQL & " sum(viewedrequests) as numviewedrequests, sum(respondedrequests) as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &   varWhereClause
   sSQL = sSQL & " GROUP BY userlname, userfname "
   sSQL = sSQL & " ORDER BY userlname, userfname "
else
   sSQL = "SELECT 0 as totalresponsetime, 0 as ttlviewedrequests, '' as assigned_name, 0 as numsubmitted, 0 as totaldays, "
   sSQL = sSQL & " 0 as numviewedrequests, 0 as numrespondedrequests "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL & " WHERE 1 = 0 "
end if

'BEGIN DISPLAYING RECORDSET WITH SUMMARRY DATA
Set oRequests = Server.CreateObject("ADODB.Recordset")
oRequests.Open sSQL, Application("DSN"), 3, 1

if oRequests.EOF=false then

	 'BEGIN TABLE ROW HEADINGS		  
 	 response.write "<table cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">" & vbcrlf
	  response.write "  <tr class=""tablelist"">" & vbcrlf
	  
	 'FIRST COLUMN BASED ON GROUP BY SELECTION
	 if orderBy = "submit_date" then
   		response.write "      <th>Date</th>" & vbcrlf
	 elseif orderBy = "action_Formid" then
   		response.write "      <th>Action Line Category</th>" & vbcrlf
	 elseif orderBy = "deptId" then
   		response.write "      <th>Department</th>" & vbcrlf
	 elseif orderBy = "assigned_Name" then
   		response.write "      <th>Assigned To</th>" & vbcrlf
	 elseif orderBy = "streetname" then
  			response.write "      <th>Issue/Problem Location Street Name</th>" & vbcrlf
	 elseif orderBy = "submittedby" then
	  		response.write "      <th>Submitted By</th>" & vbcrlf
	 End if

 'SUBMIT DATE
  response.write "      <th>Submitted</td>" & vbcrlf

 'NUMBER OF REQUESTS NOT VIEWED
	 response.write "      <th>Not Viewed</td>" & vbcrlf

	'NUMBER OF REQUESTS WITH NO ACTION
	 response.write "      <th>No Action</td>" & vbcrlf

 'AVERAGE RESPONSE TIME
  response.write "      <th>Avg. Response Time</td>" & vbcrlf

  response.write "  </tr>" & vbcrlf

 'LOOP AND DISPLAY THE RECORDS
  bgcolor = "#eeeeee"
	 Do while not oRequests.EOF

    'GET BASE VALUES
 		 	iTotalRequestsSubmitted = oRequests("numSubmitted")
	    iRespondedRequests      = oRequests("numrespondedrequests")
  			iViewedRequests         = oRequests("numviewedrequests")

 			'SET COLOR FOR ALTERNATING ROWS 
  			If bgcolor="#eeeeee" Then
        bgcolor="#ffffff"
  			Else
        bgcolor="#eeeeee"
   		End If

    	If orderBy="submit_date" then
 		   		If oRequests("TheDate") <> "" Then
      					sTitle = oRequests("TheDate") 
     			Else
      					sTitle = "<font color=""red""><b>???</b></font>"
     			End If

        detaillink = "action_line_list_response.asp?orderBy=" & orderBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&substatus_hidden=" & substatus_hidden & "&show_hide_substatus=" & show_hide_substatus & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & oRequests("TheDate") & "&fromDate=" & oRequests("TheDate") & "&selectDateType=" & selectDateType & "&reporttype=DrillThru"

  			elseif orderBy="action_Formid"  then 
    				If oRequests("action_formTitle") <> "" Then
      					sTitle = oRequests("action_formTitle") 
    				Else
	     					sTitle = "<font color=""red""><b>???</b></font>"
								End If

     			detaillink = "action_line_list_response.asp?orderBy=" & orderBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & oRequests("action_formId") & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&substatus_hidden=" & substatus_hidden & "&show_hide_substatus=" & show_hide_substatus & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & toDate & "&fromDate=" & fromDate & "&selectDateType=" & selectDateType & "&reporttype=DrillThru"

  			elseif orderBy="deptId"  then 
     			If oRequests("deptId") <> "" AND IsNull(oRequests("deptId"))=false Then
      					sTitle = clng(oRequests("deptId"))
    				Else
      					sTitle = 0
    				End If

     			detaillink = "action_line_list_response.asp?orderBy=" & orderBy & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&substatus_hidden=" & substatus_hidden & "&show_hide_substatus=" & show_hide_substatus & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & toDate & "&fromDate=" & fromDate & "&selectDateType=" & selectDateType & "&reporttype=DrillThru&selectDeptId=" & oRequests("DeptID")

     			sSQLdeptName = "SELECT groupname "
        sSQLdeptName = sSQLdeptName & " FROM groups "
        sSQLdeptName = sSQLdeptName & " WHERE orgid=" & Session("OrgID")
        sSQLdeptName = sSQLdeptName & " AND groupid=" & sTitle

     			Set oDeptName = Server.CreateObject("ADODB.Recordset")
     			oDeptName.Open sSQLdeptName, Application("DSN") , 3, 1

    				If oDeptName.EOF  Then
      					sTitle = "<font color=""red""><b>???</b></font>"
    				else
      					sTitle = oDeptName("groupname") 
     			End If
  			elseif orderBy="assigned_Name"  then 
  			   If oRequests("assigned_Name") <> "" Then
     						sTitle = oRequests("assigned_Name") 
   					Else
			     			sTitle = "<font color=""red""><b>???</b></font>"
   					End If

   					detaillink = "action_line_list_response.asp?orderBy=" & orderBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&substatus_hidden=" & substatus_hidden & "&show_hide_substatus=" & show_hide_substatus & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & toDate & "&fromDate=" & fromDate & "&selectDateType=" & selectDateType & "&reporttype=DrillThru"

  			elseif orderBy="streetname"  then 
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

   					detaillink = "action_line_list_response.asp?orderBy=" & orderBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&substatus_hidden=" & substatus_hidden & "&show_hide_substatus=" & show_hide_substatus & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & toDate & "&fromDate=" & fromDate & "&selectDateType=" & selectDateType & "&reporttype=DrillThru"

  			elseif orderBy="submittedby"  then 
							if oRequests("userlname") <> "" OR oRequests("userfname") <> "" then
  								sTitle = oRequests("userfname") & " " & oRequests("userlname")
							else
		  						sTitle = "<font color=""red""><b>???</b></font>"
							end if

   					detaillink = "action_line_list_response.asp?orderBy=" & orderBy & "&selectDeptId=" & selectDeptId & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&substatus_hidden=" & substatus_hidden & "&show_hide_substatus=" & show_hide_substatus & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName & "&toDate=" & toDate & "&fromDate=" & fromDate & "&selectDateType=" & selectDateType & "&reporttype=DrillThru"
  			end if

			 'DISPLAY FIRST COLUMN VALUE - GROUP BY COLUMN
  			response.write "<tr bgcolor=""" & bgcolor & """ onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">" & vbcrlf
     response.write "    <td onClick=""location.href='" & detaillink & "';""><b>" & sTitle & " </b></td>" & vbcrlf

 			'DISPLAY NUMBER OF REQUESETS SUBMITTED
  			response.write "    <td align=""center"" onClick=""location.href='" & detaillink & "';"">" & iTotalRequestsSubmitted  & "</td>" & vbcrlf

		 	'DISPLAY NUMBER OF REQUESTS NOT VIEWED
   		response.write "    <td align=""center"" onClick=""location.href='" & detaillink & "';""> " & iTotalRequestsSubmitted - iViewedRequests  & "</td>" & vbcrlf

 			'DISPLAY NUMBER OF REQUESTS WITH NO ACTION
  			response.write "    <td align=""center"" onClick=""location.href='" & detaillink & "';""> " & iTotalRequestsSubmitted - iRespondedRequests  & "</td>" & vbcrlf
			
			 'DISPLAY AVERAGE RESPONSE TIME
  			response.write "<td align=""center"" onClick=""location.href='" & detaillink & "';"">" & vbcrlf

  			If iRespondedRequests = 0 OR IsNull(iRespondedRequests) Then
		    		response.write "None Viewed\Responded" & vbcrlf
   		Else
			    	decResponseDays = formatnumber(oRequests("totalresponsetime") / iRespondedRequests, 1)
    				If decResponseDays < 1 Then
      					response.write "< 1" & vbcrlf
    				Else
				      	response.write decResponseDays & vbcrlf
    				End If
  			End If

  			response.write "    </td>" & vbcrlf
		  	response.write "</tr>" & vbcrlf

 			'COMPUTE GRAND TOTAL VALUES
  			iGrandTotalRequestsSubmitted = iGrandTotalRequestsSubmitted + iTotalRequestsSubmitted
		  	iTotalRequestsNotViewed      = iTotalRequestsNotViewed + (iTotalRequestsSubmitted - iViewedRequests)
  			iTotalRequestsNoAction       = iTotalRequestsNoAction + (iTotalRequestsSubmitted - iRespondedRequests)
		  	iTotalResponseTime           = iTotalResponseTime + oRequests("totalresponsetime")
  			iTotalRespondedRequests      = iTotalRespondedRequests + iRespondedRequests

    	oRequests.MoveNext 
  Loop

	'DISPLAY TOTALS ROW			
 	response.write "<tr bgcolor=""#dddddd"">" & vbcrlf
  response.write "    <td style=""padding-left:90px""><b><font color=""navy"" size=""1"">TOTAL</td>" & vbcrlf
	 response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" &  iGrandTotalRequestsSubmitted & "</td>" & vbcrlf
	 response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" &  iTotalRequestsNotViewed & "</td>" & vbcrlf
	 response.write "    <td align=""center""><b> <font color=""navy"" size=""1"">" &  iTotalRequestsNoAction & "</td>" & vbcrlf
	
 	If iTotalRespondedRequests > 0 Then
	   	If iTotalResponseTime/iTotalRespondedRequests < 1 Then
     			response.write  "<td align=""center""><b> <font color=""navy"" size=""1"">< 1</td>" & vbcrlf
   		Else
			     response.write "<td align=""center""><b> <font color=""navy"" size=""1"">" &  formatnumber(iTotalResponseTime/iTotalRespondedRequests,1) & "</td>" & vbcrlf
   		End If
 	Else
    	response.write  "<td align=""center""><b> <font color=""navy"" size=""1"">None Viewed\Responded</td>" & vbcrlf
  End If

  response.Write "</table>"

Else

	'NO DATA FOUND FOR GIVEN PARAMETERS
 	response.write "<p><b>No records found</b></p>" & vbcrlf

End If

End Function


' -------------------------------------------------------------------------------------------------
' FUNCTION FNLISTFORMS()
'--------------------------------------------------------------------------------------------------
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
	

' -------------------------------------------------------------------------------------------------
' FUNCTION FNLISTDEPTS()
'--------------------------------------------------------------------------------------------------
Function fnListDepts()
	sSQL = "select groupid,orgid,groupname,groupdescription  from groups where grouptype=2 and orgid=" & Session("OrgID") & " order by groupname"
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


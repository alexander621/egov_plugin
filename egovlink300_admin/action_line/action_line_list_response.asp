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

' GET USER'S PERMISSIONS
'blnCanViewAllActionItems = HasPermission("CanViewAllActionItems")
'blnCanViewOwnActionItems = HasPermission("CanViewOwnActionItems")
'blnCanViewDeptActionItems = HasPermission("CanViewDeptActionItems")

iPermissionLevelId = GetUserPermissionLevel( Session("UserId"), "requests" )
If clng(iPermissionLevelId) > 0 Then 
  	sPermissionLevel = GetPermissionLevelName( iPermissionLevelId )
Else
	  response.redirect sLevel & "permissiondenied.asp"
End If 

' Set to use new permission levels
blnCanViewAllActionItems  = False  
blnCanViewOwnActionItems  = False 
blnCanViewDeptActionItems = False

' IF USER HAD SET FILTERS FOR THIS SESSION THEN REMEMBER THEM
If request("useSessions") = 1 then
	' USE/SET FILTERS SET FOR THE SESSION
		recordsPer              = session("recordsPer")
		reportType              = session("reportType")
			
		orderBy                 = session("orderBy")
		selectFormId            = session("selectFormId")
		selectAssignedto        = session("selectAssignedto")
		selectDeptId            = session("selectDeptId")
			
		selectUserFName         = session("selectUserFName")
		selectUserLName         = session("selectUserLName")
			
		fromDate                = session("fromDate")
		toDate                  = session("toDate")
  selectDateType          = session("selectDateType")
		today                   = Date()
			
		statusSubmitted         = session("statusSubmitted")
		statusInprogress        = session("statusInprogress")
		statusWaiting           = session("statusWaiting")
		statusResolved          = session("statusResolved")
		statusDismissed         = session("statusDismissed")

  substatus_hidden        = session("substatus_hidden")
		show_hide_substatus     = session("show_hide_substatus")

  selectIssueStreetNumber = session("selectIssueStreetNumber")
		selectIssueStreet       = session("selectIssueStreet")
  selectContactStreet     = session("selectContactStreet")
Else
	' USE/SET DEFAULT FILTERS
		recordsPer              = request("recordsPer")
		reportType              = request("reportType")
			
		orderBy                 = request("orderBy")
		selectFormId            = request("selectFormId")
			
		If (NOT blnCanViewAllActionItems) AND (NOT blnCanViewDeptActionItems) AND blnCanViewOwnActionItems Then
      if sPermissionLevel = "View Dept - Edit Dept" OR sPermissionLevel = "View Dept - Edit Own" then
    		   selectAssignedto = request("selectAssignedto")
      else
    		   selectAssignedto = Session("userid")
      end if
		Else
		   selectAssignedto = request("selectAssignedto")
		End If 
			
		selectDeptId            = request("selectDeptId")
			
		selectUserFName        = request("selectUserFName")
		selectUserLName         = request("selectUserLName")
			
		fromDate                = request("fromDate")
		toDate                  = request("toDate")
  selectDateType          = request("selectDateType")
		today                   = Date()
			
		statusSubmitted         = request("statusSUBMITTED")
		statusInprogress        = request("statusINPROGRESS")
		statusWaiting           = request("statusWAITING")
		statusResolved          = request("statusRESOLVED")
		statusDismissed         = request("statusDISMISSED")

		substatus_hidden        = request("substatus_hidden")
		show_hide_substatus     = request("show_hide_substatus")

  selectIssueStreetNumber = request("selectIssueStreetNumber")
		selectIssueStreet       = request("selectIssueStreet")
  selectContactStreet     = request("selectContactStreet")
End If


' SET REPORT TYPE (LIST,SUMMARY, OR DETAIL) FILTER
If reportType = "" or IsNull(reportType) Then reportType = "List" End If

' SET ORDER BY COLUMN FILTER
If orderBy = "" or IsNull(orderBy) Then 
 	'USE SELECT ORDER BY COLUMN
  	orderBy = request("groupBy")

 	'CHECK TO SEE IF ORDER BY HAS VALUE IF NOT DEFAULT TO SUBMIT_DATE
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

' SET CONTACT STREET FILTER
If selectContactStreet = "" or IsNull(selectContactStreet) Then selectContactStreet = "all" End If

' SET ISSUE/PROBLEM LOCATION STREET FILTER
'If selectIssueStreet = "" or IsNull(selectIssueStreet) Then selectIssueStreet = "all" End If

' SET TODATE FILTER
If toDate = "" or IsNull(toDate) Then toDate = dateAdd("d",0,today) End If
toDate = dateAdd("d",1,toDate)

' SET FROMDATE FILTER
If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("yyyy",-1,today) End If

' SET RECORDS PER PAGE FILTER
if recordsPer = "" or IsNull(recordsPer) Or clng(recordsPer) = 0 Then recordsPer = 25 End If

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

' SET SESSION VARIABLES FOR REMEMBERING DURING THIS SESSION
session("reportType")              = reportType
session("orderBy")                 = orderBy
session("selectFormId")            = selectFormId
session("selectAssignedto")        = selectAssignedto 
session("selectDeptId")            = selectDeptId

session("toDate")                  = toDate
session("fromDate")                = fromDate
session("selectDateType")          = selectDateType
session("recordsPer")              = recordsPer

session("noStatus")                = noStatus
session("statusDismissed")         = statusDismissed
session("statusResolved")          = statusResolved
session("statusWaiting")           = statusWaiting
session("statusInprogress")        = statusInprogress
session("statusSubmitted")         = statusSubmitted

session("substatus_hidden")        = substatus_hidden
session("show_hide_substatus")     = show_hide_substatus

session("selectUserFName")         = selectUserFName
session("selectUserLName")         = selectUserLName

session("selectIssueStreetNumber") = selectIssueStreetNumber
session("selectIssueStreet")       = selectIssueStreet
session("selectContactStreet")     = selectContactStreet
%>

<html>
<head>
  <title><%=langBSActionLine%></title>
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  
<script language="javascript">
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
		 if (document.form1.reportType.value == "Summary") {
				document.forms[0].action = "action_line_summary.asp";
				document.forms[0].submit();
		} 
		else if (document.form1.reportType.value == "ResponseSummary") {
				document.forms[0].action = "action_line_summary_response.asp";
				document.forms[0].submit();
			}
		else if (document.form1.reportType.value == "responsedetail") {
				document.forms[0].action = "action_line_list_response.asp";
				document.forms[0].submit();
			}
		else if (document.form1.reportType.value == "ListFull") {
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

<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"> -->
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
  <tr><td><font size="+1"><b>(E-Gov Request Manager) - Manage Action Line Requests</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langBackToStart%></a></td></tr>
  <tr><td>
		  <!--BEGIN: FILTER SELECTION-->
      		  <fieldset>
       			  <legend><b>Search/Sorting Option(s)</b></legend>
     			  <form name="form1"  onSubmit="return checkStat()">
      		  <table border="0" bordercolor="red">
      		    <tr>
            		  <!--ASSIGNED USER FILTER-->
            		  <td valign="top" nowrap>
        		  			 <%
        		  			   'DISPLAY ASSIGNED TO SELECTION
'        		    			 If blnCanViewAllActionItems Then		  
        		       				response.write "<strong>Assigned To: </strong>"

                     if  NOT blnCanViewAllActionItems _
                     AND NOT blnCanViewDeptActionItems then
         		     	   	 		'DISPLAY CURRENTLY LOGGED IN ADMINISTRATOR
         		      			   		response.write "(User " & session("userID") & ")"
                         response.write "<input type=""hidden"" name=""selectAssignedto"" value=""" & selectAssignedto & """>"
                     else
                		  				'DRAW LIST OF EMPLOYEES
                		   				DrawAssignedEmployeeSelection(session("orgid"))
           		    			 end if
              	 %>
        		    	     <!--ORDER BY FILTER-->
			                 <strong>Order By: </strong>
                 			<select name="orderBy">
 			<% if orderBy = "assigned_Name" then select1 = "selected" else select1="" %>
                 					<option value="assigned_Name" <%=select1%>>Assigned To</option>	
 			<% if orderBy = "action_Formid" then select1 = "selected" else select1="" %>
                 					<option value="action_Formid" <%=select1%>>Category</option>
			 <% if orderBy = "submit_Date" then select1 = "selected" else select1="" %>
                 					<option value="submit_Date" <%=select1%>>Date Descending</option>
 			<% if orderBy = "deptId" then select1 = "selected" else select1="" %>
                 					<option value="deptId" <%=select1%>>Department</option>
 <% If OrgHasFeature("issue location") Then %>
  		<% if orderBy = "streetname" then select1 = "selected" else select1="" %>
                 					<option value="streetname" <%=select1%>>Issue/Problem Location Street Name</option>
    <% end if %>
				<% if orderBy = "submittedby" then select1 = " selected=""selected"" " else select1="" end if %>
                      <option value="submittedby" <%=select1%>>Submitted By</option>
                    </select>
                </td>
            </tr>
            <tr>
                <td valign="top" nowrap>
                    <!--STATUS FILTER-->
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
                              <input type="button" value="Show/Hide Sub-Status List" onclick="javascript:showhide_substatus_criteria();"></td></tr>
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
								                            </td>
                                </tr>
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
                </td></tr>
            <tr>
            		  <td valign="top" nowrap>
                		  <!--CATEGORY FILTER-->
               		   <b>Category: 
                	   <select name="selectFormId"><option value="">All Categories</option><% fnListForms()%></select>
               	</td>
            </tr>
            <tr>
            		  <td valign="top" nowrap>
                		  <!--DEPARTMENT FILTER-->
                 			<% 
                   			If blnCanViewAllActionItems OR blnCanViewDeptActionItems  Then 
                     				response.write "<strong>Department: </strong> "
                     				response.write "<select name=""selectDeptId""><option value=""all"">All Departments</option>"

                    				'GET A LIST OF ALL AVAILABLE DEPARTMENTS FOR THIS USER
                     				fnListDepts selectDeptId 

                     				response.write "</select>&nbsp;&nbsp;&nbsp;"
                		    End If
                  		%>
                  	 <!--REPORT TYPE FILTER-->		 
                    <b>Report Type: 
             			    <select name="reportType">
      <% if reportType = "Detail" then %>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
				                  <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail" selected>Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail">Response Detail</option>
                   <% End If%>
			   <% elseif reportType = "Summary" then %>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					                 <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary" selected>Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail">Response Detail</option>
                   <% End If%>
			   <% elseif reportType = "ResponseSummary" Then%>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					                 <option value="ListFull">List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary" selected>Response Summary</option>
                      <option value="responsedetail">Response Detail</option>
                   <% End If%>
               <% elseif reportType = "responsedetail" Then%>
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
			   <% elseif reportType = "ListFull" then%>
                      <option value="List">List</option>
                   <% If OrgHasFeature("actionline_listfull") Then %>
					                 <option value="ListFull" selected>List (Full)</option>
                   <% end if %>
                      <option value="Summary">Summary</option>
                      <option value="Detail">Detail</option>
                   <% If OrgHasFeature("responsetimereporting") Then %>
                      <option value="ResponseSummary">Response Summary</option>
                      <option value="responsedetail">Response Detail</option>
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
                      <option value="responsedetail">Response Detail</option>
                   <% End If%>
			   <% end if %>
             		     </select>
             		 </td>
            </tr>
            <tr>
            		  <td valign="top" nowrap>
                		  <!--DATE RANGE FILTER-->
                		  <fieldset>
                      <b>From: 
		  	              	  <input type="text" name="fromDate" value="<%=fromDate%>">
               					  <a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border=0></a>		 
               				   &nbsp; 
                 					<b>To:</b> 
               					  <input type="text" name="toDate" value="<%=dateAdd("d",-1,toDate)%>">
               					  <a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border=0></a>
                      &nbsp;
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
         		     <!--SUBMITTED BY FILTER-->
                <td valign="top" nowrap>
                    <b>Submitted By: &nbsp;&nbsp;
              				  First: <input id="subfirstname" type="text" name="selectUserFName" value="<% if selectUserFName <> "all" then response.write selectUserFName %>" size=12>
           			  	   &nbsp; 
                				<b>Last:</b> 
           					    <input id="sublastname" type="text" name="selectUserLName" value="<% if selectUserLName <> "all" then response.write selectUserLName %>" size=12>
             			</td>
            </tr>
            <tr>
                <td valign="top" nowrap>
              		    <b>Contact Street Name: <input type="text" name="selectContactStreet" value="<%=selectContactStreet%>">
                </td>
            </tr>

            <!--BEGIN: STREET FILTER-->
            <% If OrgHasFeature("issue location") Then %>
            <tr>
                <td valign="top" nowrap>
                    <strong>Issue/Problem Location: &nbsp;&nbsp;
                    Street Number: <input type="text" name="selectIssueStreetNumber" value="<%=selectIssueStreetNumber%>" size="10" maxlength="150">
                    &nbsp;
                    Street Name: </strong> <input type="text" name="selectIssueStreet" value="<%=selectIssueStreet%>" size="30" maxlength="300">
                </td>
            </tr>
            <% End If %>
            <!--END: STREET FILTER-->

         		 <!--RECORDS PER PAGE FILTER-->
          	 <tr>
            		  <td valign="top">
              			   <input type="button" style="cursor: hand" onclick="javascript:submitForm();" value=" SEARCH ">
          		        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              		    <b>Records per Page: 
                				<input type="text" name="recordsPer" value="<%=recordsPer%>" size="5" maxlength="4">
            		  </td>
       		   </tr>
      		  </table>
			       </form>
       			</fieldset>
       <% if OrgHasFeature("issue location") then %>
          <div align="right">
            <font style="color: #FF0000;">* <small><i>= Non-Listed Street Address</i></small></font>
          </div>
       <% end if %>
    			<!--END: FILTER SELECTION-->
  			 </td>
  </tr>
  <tr>
      <td valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
          <form name="requestlist" action="#" method="POST">
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

'Check the selectDateType to determine how to use the From/To Date fields.
if UCASE(selectDateType) = "ACTIVE" then
   varWhereClause = " WHERE  (egov_action_request_view.orgid=" & session("orgid") & ") AND ( "    ''IsNull(complete_date,'" & Now & "')
   varWhereClause = varWhereClause & " (submit_Date >= '" & fromDate & "' AND submit_Date < '" & toDate & "') OR "
   varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
   varWhereClause = varWhereClause & " (submit_Date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "
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

'varWhereClause = varWhereClause & " ) AND (" & varStatClause & ") "

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

'If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '" & selectAssignedto & "'" 
If selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assignedemployeeid = " & selectAssignedto

If blnCanViewDeptActionItems AND NOT blnCanViewAllActionItems Then
	If selectDeptId <> "all" then 
		varWhereClause = varWhereClause & " AND deptID = '" & selectDeptId & "'" 
	Else
     		varWhereClause = REPLACE(varWhereClause,"AND (" & varStatClause & ")","") & " AND (((" & varStatClause & ") AND deptID IN (" & GetGroups(session("userid")) & ")) OR ((assignedemployeeid = '" & session("userid") & "') and (" & varStatClause  & "))) "
'		varWhereClause = varWhereClause & " AND deptID IN (" & GetGroups(session("userid")) & ") OR (assignedemployeeid = '" & session("userid") & "') " 
	End If
Else
	If selectDeptId <> "all" then 
		varWhereClause = varWhereClause & " AND deptID = '" & selectDeptId & "'" 
	End If
End If

If selectUserFName <> "all" then varWhereClause = varWhereClause & " AND UserFName LIKE '" & selectUserFName & "%'"
If selectUserLName <> "all" then varWhereClause = varWhereClause & " AND UserLName LIKE '" & selectUserLName & "%'"

'CONTACT STREET NAME FILTER
If selectContactStreet <> "all" then varWhereClause = varWhereClause & " AND useraddress LIKE '%" & selectContactStreet & "%'"

'ISSUE/PROBLEM LOCATION STREET NAME FILTER
'If selectIssueStreet <> "all" then varWhereClause = varWhereClause & " AND streetname LIKE '%" & selectIssueStreet & "%'"
 lcl_search_address = ""
 if selectIssueStreetNumber <> "" AND NOT isnull(selectIssueStreetNumber) then
    lcl_search_address = selectIssueStreetNumber & "%"
 end if

 lcl_search_address = lcl_search_address & selectIssueStreet

 if lcl_search_address <> "" then
    varWhereClause = varWhereClause & " AND UPPER(streetname) LIKE ('%" & UCASE(lcl_search_address) & "%')"
 end if

	sSQL = "SELECT userlname, userfname, useraddress, usercity, userstate, userhomephone, action_autoid, action_formTitle, "
 sSQL = sSQL & " DateDiff(d,submit_date,ISNULL(complete_date,'" & today & "')) AS totalDays, "
 sSQL = sSQL & " DateDiff(d,ISNULL(adjustedsubmitdate,submit_date),ISNULL(complete_date,'" & today & "')) AS totalDays_adjusted, "
 sSQL = sSQL & " [dbo].[getDateDiff_NoWeekend] (egov_action_request_view.orgid,egov_action_request_view.submit_date,ISNULL(egov_action_request_view.complete_date,'" & today & "')) AS totalDays_noweekends, "
 sSQL = sSQL & " [dbo].[getDateDiff_NoWeekend] (egov_action_request_view.orgid,ISNULL(egov_action_request_view.adjustedsubmitdate,egov_action_request_view.submit_date),ISNULL(egov_action_request_view.complete_date,'" & today & "')) AS totalDays_noweekends_adjusted, "
	sSQL = sSQL & " submit_date, complete_date, deptID, groupname as deptName, status, assignedName as assigned_Name, assignedemployeeid, "
	sSQL = sSQL & " streetnumber, streetaddress, streetname, responsedatetime, firstactiondate, responsetime, issuelocationname, "
 sSQL = sSQL & " city, state, zip, validstreet, comments, action_form_display_issue, "
 sSQL = sSQL & " (select ISNULL(status_name,'') from egov_actionline_requests_statuses where sub_status_id = action_status_id) AS sub_status_name "
	sSQL = sSQL & " FROM egov_action_request_view "
 sSQL = sSQL & " LEFT OUTER JOIN groups on deptId=groupId"
 sSQL = sSQL &   varWhereClause
 sSQL = sSQL & " AND (egov_action_request_view.orgid=" & session("orgid") & ") "

 if orderBy = "streetname" then
    lcl_order_by = "UPPER(streetaddress), CAST(streetnumber AS int) "
 elseif orderBy = "submittedby" then
    lcl_order_by = "UPPER(userlname), UPPER(userfname) "
 else
    lcl_order_by = orderBy
 end if

 sSQL = sSQL & " ORDER BY " & lcl_order_by

 if orderBy = "submit_date" then
    sSQL = sSQL & " desc"
 end if

' BEGIN: STORE QUERY FOR EXPORT TO CSV
session("DISPLAYQUERY") = "SELECT "
session("DISPLAYQUERY") = session("DISPLAYQUERY") & " CAST(action_autoid AS varchar) + RIGHT('0'+ CAST(DATEPART(hh, submit_date) AS varchar),2) + RIGHT('0' + CAST(DATEPART(mi, submit_date) AS varchar),2) as [E-Gov Link Tracking ID],action_formtitle as [Action Form Name],submit_date as [Date Submitted],status as [Status],userfname + ' ' + userlname as [Submitted By],assigned_Name as [Assigned To],groupname as [Department],comment as [Form Values]  "
session("DISPLAYQUERY") = session("DISPLAYQUERY") & " " & RIGHT(sSQL,Len(sSQL)-instr(sSQL,"FROM")+1)
' END: STORE QUERY FOR EXPORT TO CSV

Set oRequests = Server.CreateObject("ADODB.Recordset")

	 ' SET PAGE SIZE AND RECORDSET PARAMETERS
	 oRequests.PageSize       = recordsPer
	 oRequests.CacheSize      = recordsPer
	 oRequests.CursorLocation = 3

	 ' OPEN RECORDSET
	 oRequests.Open sSQL, Application("DSN"), 3, 1

	lastTitle        = "Test Title"
	lastDate         = "1/1/02"
	lastDept         = 11798
	lastDeptName     = "Test Department"
	lastAssigned     = "John Doe"
	displayLastTitle = "Test Department Last"
	lastSubmitted    = "John Last Doe"

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

 'DISPLAY RECORD STATISTICS
	 Dim abspage, pagecnt
		abspage = oRequests.AbsolutePage
		pagecnt = oRequests.PageCount

	 If request("selectAssignedto") <> "" Then
		   sQueryString = replace(request.querystring,"pagenum","HFe301") ' REPLACE PAGENUM FIELD WITH RANDOM FIELD FOR NAVIGATION PURPOSES
  Else
  		 sQueryString = "filter=false"
  End If

  response.write "<br><font size=""3"" color=""3399ff""><i><b>RESPONSE DETAIL REPORT</b></i></font>"
  response.write "<br><b>Page <font color=""blue"">" & oRequests.AbsolutePage & "</font>  " & vbcrlf
  response.write "of <font color=""blue""> " & oRequests.PageCount & "</font></b> &nbsp;|&nbsp; " & vbcrlf
  response.write "<b><font color=""blue"">" & oRequests.RecordCount & "</font> total Action Item Requests</b>"

 'BEGIN: INSERT LINK TO EXPORT RESULTS
  If OrgHasFeature("action export") Then
   		If request.querystring("reportType") = "" or LCASE(request.querystring("reportType")) = "list" Then
			    'DISPLAY EXPORT BUTTON ONLY FOR LIST VIEW
     			response.write "&nbsp;&nbsp;<input type=""button"" class=""excelexport"" value=""Download as CSV"" onClick=""location.href='../export/csv_export.asp'"">"
   		End If
  End If
 'END: INSERT LINK TO EXPORT RESULTS

 'DISPLAY FORWARD AND BACKWARD NAVIGATION TOP AND PRINT PAGE LINK
  response.write "<div>"
  response.write "<table width=""100%"">"
  response.write "  <tr>"
  response.write "      <td valign=""top"">"
  response.write "          <table>"
  response.write "                 <tr>"
  response.write "                     <td><a href=""action_line_list_response.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=""0"" src=""../images/arrow_back.gif""></a></td>"
  response.write "                     <td valign=""top""><a href=""action_line_list_response.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td>"
  response.write "                     <td valign=""top"">&nbsp;" & "<a href=""action_line_list_response.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td>"
  response.write "                     <td valign=top><a href=""action_line_list_response.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=""0"" src=""../images/arrow_forward.gif"" valign=""bottom""></a></td>"
  response.write "                 </tr>"
  response.write "          </table>"
  response.write "      </td>"
  response.write "      <td width=""450"" align=""right"">"
  response.write "          <!--<a href=""action_line_list_print.asp?orderBy=" & orderBy & "&selectFormId=" & selectFormId & "&selectAssignedto=" & selectAssignedto & "&statusSubmitted=" & statusSubmitted & "&statusInprogress=" & statusInprogress & "&statusWaiting=" & statusWaiting & "&statusResolved=" & statusResolved & "&statusDismissed=" & statusDismissed & "&selectUserLName=" & selectUserLName & "&selectUserFName=" & selectUserFName &"&toDate=" & toDate & "&fromDate=" & fromDate & "&selectDateType=" & selectDateType & "&reportType=" & reportType & """ target=new>Open New Printer Friendly Results Window</a>-->"
  response.write "      </td>"
  response.write "  </tr>"
  response.write "</table>"
  response.write "</div>"
  response.write "<table cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">"
  response.write "  <tr valign=""bottom"" class=""tablelist"">"

	      ' CHANGE TO BCANEDIT LATER
	      If 1=1 Then 
			' Response.Write "<th><input class=""listCheck"" type=checkbox name=""chkSelectAll"" onClick=""selectAll('requestlist', this.checked)""></th>"
	      Else
	            ' Response.Write "<th>&nbsp;</th>"
	      End If
	
	  response.write "<th>Action Line Category</th>"
	  response.write "<th>Date submitted</th>"

	  if ReportType="Detail" or ReportType="DrillThru" or ReportType="responsedetail" then
  	  	response.write "<th>First Viewed</th>"
    		response.write "<th>First Action</th>"
    		response.write "<th>Days to Respond</th>"
	  end if

	  response.write "<th>Status</th>"

	  if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
      response.write "<th>Sub-Status</th>"
   end if

	  response.write "<th>Submitted by</th>"
	  response.write "<th>Contact<br>Street Name</th>"
	  response.write "<th>Assigned to</th>"
	  response.write "<th>Department</th>"

	  If OrgHasFeature("issue location") Then
    		response.write "<th>Issue/Problem Location<br>Street Name</th>"
	  End If

	 response.write "</tr>"

sSQLtotl    = "SELECT action_autoid,action_formTitle,DateDiff(d,submit_date,complete_date) AS totalDays,submit_date,complete_date,deptID,groupname as deptName,status,assigned_Name FROM egov_action_request_view left outer join groups on deptId=groupId" & varWhereClause
Set oTotals = Server.CreateObject("ADODB.Recordset")
oTotals.Open sSQLtotl, Application("DSN"), 3, 1
	 '/////////////////

sSQLTotal  = "SELECT count(*) as numTotal FROM egov_action_request_view " & varWhereClause   
response.write "<!--JOHNCOMMENT" & sSQLTotal & "-->"
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

	' GRAND TOTAL FOR RECORDS
	If ReportType="Detail" or ReportType="DrillThru" or ReportType="responsedetail" then	

		'Response.Write "<tr bgcolor=#dddddd><td style=""padding-left:90px""><b><font color=navy size=1>Grand Total [" & oRequests.RecordCount & " Requests]</td><td> <b><font color=navy size=1> Submitted: " & oTotal("numTotal") & "</td><td> <b><font color=navy size=1> Open: " & oOpen("numOpen") & "</td><td colspan=2> <b><font color=navy size=1> Avg Time Still Open: " & avgOpenTotal & "</td><td colspan=2> <b><font color=navy size=1> Avg Time To Complete: " & avgClosedTotal & "</td><td>&nbsp;</td><td>&nbsp;</td><td colspan=3>&nbsp;</td></tr>"

	End if

'*** LOOP AND DISPLAY THE RECORDS
 bgcolor = "#eeeeee"
' iNumSubmittedRequests = -1
 iNumSubmittedRequests = 0

 For intRec=1 To oRequests.PageSize
 		  If Not oRequests.EOF Then

     			If bgcolor="#eeeeee" Then
       				bgcolor="#ffffff" 
     			Else
       				bgcolor="#eeeeee"
     			End If

    			'TRACK NUMBER OF SUBMITTED REQUESTS
     			iNumSubmittedRequests = iNumSubmittedRequests + 1

    		 'GET VALUES
     			If oRequests("action_formTitle") <> "" Then
       				sTitle = oRequests("action_formTitle")
     			Else
       				sTitle = "<font color=""red""><b>???</b></font>"
     			End If

     		'SUBMIT DATE
			     sDate = oRequests("submit_date")
     			sDate = formatDateTime(sDate,vbShortDate)

     			If oRequests("assigned_Name") <> "" Then
       				sAssigned = oRequests("assigned_Name")
     			Else
       				sAssigned = "<font color=""red""><b>???</b></font>"
     			End If

     			If oRequests("userlname") <> "" Then
  			     	sSubmitted = oRequests("userfname") & " " & oRequests("userlname")
     			Else
		  	     	sSubmitted = "<font color=""red""><b>???</b></font>"
     			End If

    			'DEPARTMENT INFORMATION	
     			sDept     = oRequests("deptId")
     			sDeptName = oRequests("deptName")
     			'Else
     			'	sDept = "<font color=red><b>???</b></font>"
     			'End If

						 'FIRST VIEWED DATETIME
  						datFirstViewed = "N\A"
		  				If oRequests("responsedatetime") <> "" AND NOT IsNull(oRequests("responsedatetime") <> "") Then
			    				datFirstViewed =  oRequests("responsedatetime")
						  End If

						 'FIRST ACTION DATETIME
			  			datFirstAction = "N\A"
			  			If oRequests("firstactiondate") <> "" AND NOT IsNull(oRequests("firstactiondate") <> "") Then
			    				datFirstAction =  oRequests("firstactiondate")
			  			End If

    			'FIRST ACTION DATETIME
			  			If oRequests("responsetime") <> "" Then
				       iResponseTime =  oRequests("responsetime")
      		End If

						 'STATUS
			     If oRequests("status") <> "" Then
       				sStatus = oRequests("status")
     			Else
       				sStatus = "<font color=""red""><b>???</b></font>"
     			End If

       'SUB-STATUS
        if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
           sSubStatus = oRequests("sub_status_name")
        else
           sSubStatus = ""
        end if

     		'SUBMIT DATE
     			If oRequests("submit_date") <> "" Then
       				datSubmitDate = oRequests("submit_date")
      		Else
     				datSubmitDate = "<font color=""red""><b>???</b></font>"
     			End If

     		'COMPLETE DATE
     			If oRequests("complete_date") <> "" Then
       				datResolveDate = oRequests("complete_date")
     		  		datResolveDate = formatdatetime(datResolveDate,vbShortDate)
     			Else
     				  datResolveDate = "<font color=red><b>???</b></font>"
     			End If

     		'TRACKING NUMBER
     			lngTrackingNumber = oRequests("action_autoid") & replace(FormatDateTime(oRequests("submit_date"),4),":","")

     		'SUBTOTAL ROW FOR DETAIL REPORT
     			If ReportType="Detail" or ReportType="DrillThru" or ReportType="responsedetail" then

      					If orderBy = "submit_Date" then

     						 	'SUBTOTAL ROW ORDER BY SUBMIT DATE
       							If DateDiff("d",sDate,lastDate) = 0 then
							       		'NO NEW LINE
       							Else
         								If lastDate <> "1/1/02" then
           								'GET VIEW AND NO ACTION TOTALS
           									sSQL = "SELECT count(*) as numSubmitted, sum(viewedrequests) as numviewedrequests, "
                    sSQL = sSQL & " sum(respondedrequests) as numrespondedrequests "
                    sSQL = sSQL & " FROM egov_action_request_view "
                    sSQL = sSQL &   varWhereClause
                    sSQL = sSQL & " AND submitdateshort = '" & lastDate & "'"

  																		Set oSubTotals = Server.CreateObject("ADODB.Recordset")
  																		oSubTotals.Open sSQL, Application("DSN"), 3, 1	

  																		If NOT oSubTotals.EOF Then
  									  										iNumSubmittedRequests = oSubTotals("numSubmitted")
  									  										iNumResponsesViewed   = oSubTotals("numviewedrequests")
  									  										iNumResponsesNoAction = oSubTotals("numrespondedrequests")
  									  										oSubTotals.Close
  																		End If

  																		Set oSubTotals = Nothing

  																	'DISPLAY SUBTOTAL ROW
  																		response.write "<tr bgcolor=""#dddddd"" STYLE=""COLOR:NAVY;"">" & vbcrlf
  																		response.write "    <td align=""center""><b>Subtotals:</td>" & vbcrlf
  																		response.write "    <td> <b>Submitted: "              & iNumSubmittedRequests                         & "</td>" & vbcrlf
  																		response.write "    <td> <b>Not Viewed: "             & iNumSubmittedRequests - iNumResponsesViewed   & "</td>" & vbcrlf
  																		response.write "    <td colspan=""2""><b>No Action: " & iNumSubmittedRequests - iNumResponsesNoAction & "</td>" & vbcrlf
  																		response.write "    <td colspan=""7"">&nbsp;</td>" & vbcrlf
  																		response.write "</tr>" & vbcrlf
  															End if
  												End if

     						elseif orderBy = "action_Formid" then
      							'SUBTOTAL ROW ORDER BY ACTION FORM
       							If sTitle = lastTitle then
							       		'NO NEW LINE
       							Else
        									If lastTitle <> "Test Title" then
         										'GET VIEW AND NO ACTION TOTALS
          										sSQL = "SELECT count(*) as numSubmitted, sum(viewedrequests) as numviewedrequests, "
                    sSQL = sSQL & " sum(respondedrequests) as numrespondedrequests "
                    sSQL = sSQL & " FROM egov_action_request_view "
                    sSQL = sSQL &   varWhereClause
                    sSQL = sSQL & " AND action_FormTitle='" & lastTitle & "'"

          										Set oSubTotals = Server.CreateObject("ADODB.Recordset")
          										oSubTotals.Open sSQL, Application("DSN"), 3, 1	

          										If NOT oSubTotals.EOF Then
										            	iNumSubmittedRequests = oSubTotals("numSubmitted")
            											iNumResponsesViewed   = oSubTotals("numviewedrequests")
            											iNumResponsesNoAction = oSubTotals("numrespondedrequests")
            											oSubTotals.Close
          										End If

          										Set oSubTotals = Nothing

          									'DISPLAY SUBTOTAL ROW
          										response.write "<tr bgcolor=""#dddddd"" STYLE=""COLOR:NAVY;"">" & vbcrlf
          										response.write "    <td align=""center""><b>Subtotals:</td>" & vbcrlf
          										response.write "    <td> <b>Submitted: "              & iNumSubmittedRequests                         & "</td>" & vbcrlf
          										response.write "    <td> <b>Not Viewed: "             & iNumSubmittedRequests - iNumResponsesViewed   & "</td>" & vbcrlf
          										response.write "    <td colspan=""2""><b>No Action: " & iNumSubmittedRequests - iNumResponsesNoAction & "</td>" & vbcrlf
          						    response.write "    <td colspan=""7"">&nbsp;</td>" & vbcrlf
          										response.write "</tr>" & vbcrlf
            					End If
          				end if

     						elseif orderBy = "deptId" then
   							   'SUBTOTAL ROW ORDER BY DEPARTMENT ID
       							if sDept = lastDept then
         							'NO NEW LINE
       							else
         								if lastDept <> 11798 then
          									'GET VIEW AND NO ACTION TOTALS
     															sSQL = "SELECT count(*) as numSubmitted, sum(viewedrequests) as numviewedrequests, "
                    sSQL = sSQL & " sum(respondedrequests) as numrespondedrequests "
                    sSQL = sSQL & " FROM egov_action_request_view "
                    sSQL = sSQL &   varWhereClause
                    sSQL = sSQL & " AND deptId = '" & lastDept & "'"

          										Set oSubTotals = Server.CreateObject("ADODB.Recordset")
          										oSubTotals.Open sSQL, Application("DSN"), 3, 1	

          										If NOT oSubTotals.EOF Then
            											iNumSubmittedRequests = oSubTotals("numSubmitted")
            											iNumResponsesViewed   = oSubTotals("numviewedrequests")
            											iNumResponsesNoAction = oSubTotals("numrespondedrequests")
            											oSubTotals.Close
          										End If

          										Set oSubTotals = Nothing

          									'DISPLAY SUBTOTAL ROW
          										response.write "<tr bgcolor=""#dddddd"" STYLE=""COLOR:NAVY;"">" & vbcrlf
          										response.write "    <td align=""center""><b>Subtotals:</td>" & vbcrlf
          										response.write "    <td> <b>Submitted: "              & iNumSubmittedRequests                         & "</td>" & vbcrlf
          										response.write "    <td> <b>Not Viewed: "             & iNumSubmittedRequests - iNumResponsesViewed   & "</td>" & vbcrlf
          										response.write "    <td colspan=""2""><b>No Action: " & iNumSubmittedRequests - iNumResponsesNoAction & "</td>" & vbcrlf
          						    response.write "    <td colspan=""7"">&nbsp;</td>" & vbcrlf
          										response.write "</tr>" & vbcrlf
            					end if
          				end if
     						elseif orderBy = "assigned_Name" then
       						'SUBTOTAL ROW ORDER BY ASSIGNED NAME
       							if sAssigned = lastAssigned then
							        	'NO NEW LINE
       							else
							       	  If lastAssigned <> "John Doe" then						
         										'GET VIEW AND NO ACTION TOTALS
          										sSQL = "SELECT count(*) as numSubmitted, sum(viewedrequests) as numviewedrequests, "
                    sSQL = sSQL & " sum(respondedrequests) as numrespondedrequests "
                    sSQL = sSQL & " FROM egov_action_request_view "
                    sSQL = sSQL &   varWhereClause
                    sSQL = sSQL & " AND assignedName = '" & lastAssigned & "'"

          										Set oSubTotals = Server.CreateObject("ADODB.Recordset")
          										oSubTotals.Open sSQL, Application("DSN"), 3, 1	

          										If NOT oSubTotals.EOF Then
										            	iNumSubmittedRequests = oSubTotals("numSubmitted")
          											  iNumResponsesViewed   = oSubTotals("numviewedrequests")
										           	 iNumResponsesNoAction = oSubTotals("numrespondedrequests")
            											oSubTotals.Close
          										End If

          										Set oSubTotals = Nothing

          									'DISPLAY SUBTOTAL ROW
          										response.write "<tr bgcolor=""#dddddd"" STYLE=""COLOR:NAVY;"">" & vbcrlf
          										response.write "    <td align=""center""><b>Subtotals:</td>" & vbcrlf
          										response.write "    <td> <b>Submitted: "              & iNumSubmittedRequests                         & "</td>" & vbcrlf
          										response.write "    <td> <b>Not Viewed: "             & iNumSubmittedRequests - iNumResponsesViewed   & "</td>" & vbcrlf
          										response.write "    <td colspan=""2""><b>No Action: " & iNumSubmittedRequests - iNumResponsesNoAction & "</td>" & vbcrlf
          						    response.write "    <td colspan=""7"">&nbsp;</td>" & vbcrlf
          										response.write "</tr>" & vbcrlf
        									end if
       					  End If
     						elseif orderBy = "submittedby" then
     							 'SUBTOTAL ROW ORDER BY USERNAME
        						if sSubmitted = lastSubmitted then
   		     						'NO NEW LINE
     			  				else
					       			  if lastSubmitted <> "John Last Doe" then						
							  	         'GET VIEW AND NO ACTION TOTALS
          										sSQL = "SELECT count(*) as numSubmitted,sum(viewedrequests) as numviewedrequests, "
                    sSQL = sSQL & " sum(respondedrequests) as numrespondedrequests "
                    sSQL = sSQL & " FROM egov_action_request_view "
                    sSQL = sSQL &   varWhereClause
                    sSQL = sSQL & " AND userLname + ' ' + userFname='" & sSubmitted & "'"

    		 			     					Set oSubTotals = Server.CreateObject("ADODB.Recordset")
		    	 							     oSubTotals.Open sSQL, Application("DSN"), 3, 1	

     				  		   				If NOT oSubTotals.EOF Then
  			     	  			   				iNumSubmittedRequests = oSubTotals("numSubmitted")
   		  	       								iNumResponsesViewed   = oSubTotals("numviewedrequests")
			   	  	  			     			iNumResponsesNoAction = oSubTotals("numrespondedrequests")
     						     	  				oSubTotals.Close
          										End If

     		  	   							Set oSubTotals = Nothing

    	      								'DISPLAY SUBTOTAL ROW
          										response.write "<tr bgcolor=""#dddddd"" STYLE=""COLOR:NAVY;"">" & vbcrlf
        	  									response.write "    <td align=""center""><b>Subtotals:</td>" & vbcrlf
			          							response.write "    <td> <b>Submitted: "              & iNumSubmittedRequests                         & "</td>" & vbcrlf
     						     				response.write "    <td> <b>Not Viewed: "             & iNumSubmittedRequests - iNumResponsesViewed   & "</td>" & vbcrlf
 				     					    	response.write "    <td colspan=""2""><b>No Action: " & iNumSubmittedRequests - iNumResponsesNoAction & "</td>" & vbcrlf
	   					           response.write "    <td colspan=""7"">&nbsp;</td>" & vbcrlf
       				   						response.write "</tr>" & vbcrlf
    	     							end if
      			  			end if
     						end if
        end if

			' BEGIN DISPLAY DETAILS FOR EACH ROW
			Response.Write "<tr bgcolor=""" & bgcolor & """ onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"" onClick=""location.href='action_respond.asp?control=" & oRequests("action_autoid") & "';""><!--<td><input type=checkbox name=""del_" & oRequests("action_autoid") & """ >--><td><b>(" & lngTrackingNumber & ") " & sTitle & " </b></td><td align=center> " & formatdatetime(datSubmitDate,vbShortDate) & "</td>"
			
			If ReportType="Detail" or ReportType="DrillThru" or ReportType="responsedetail" then
					
					' FIRST VIEWED
					response.write "<td>" & datFirstViewed & "</td>" & vbcrlf

					' FIRST ACTION
					response.write "<td>" & datFirstAction & "</td>" & vbcrlf

					' RESPOND TIME
					If iReponseTime < 1 Then
						response.write "<td align=""center"">< 1</td>" & vbcrlf
					Else
						response.write "<td align=""center"">" & iResponseTime & "</td>" & vbcrlf
					End If
					

			End if

			response.write "<td align=""center""> " & UCASE(sStatus) & "</td>" & vbcrlf

   if UserHasPermission( Session("UserId"), "action_line_substatus" ) then
      response.write "<td align=""center""> " & sSubStatus & "</td>" & vbcrlf
   end if

			response.write "<td align=""center""> " & oRequests("userfname")     & " " & oRequests("userlname") & "</td>" & vbcrlf
			response.write "<td>"                   & oRequests("useraddress")   & "</td>" & vbcrlf
			response.write "<td align=""right"">"   & oRequests("assigned_Name") & "</td>" & vbcrlf
			response.write "<td align=""center""> " & oRequests("deptName")      & "</td>" & vbcrlf

   if OrgHasFeature("issue location") then
      if oRequests("validstreet") <> "Y" AND oRequests("action_form_display_issue") then
         lcl_valid_street = "<font style=""color: #FF0000;"">&nbsp;*</font>"
      else
         lcl_valid_street = ""
      end if
      response.write "<td>" & oRequests("streetname") & lcl_valid_street & "</td>" & vbcrlf
   end if


'			if OrgHasFeature("issue location") then
'      if oRequests("streetnumber") <> "" then
'         lcl_valid_street = "<font style=""color: #FF0000;"">&nbsp;*</font>"
'      else
'         lcl_valid_street = ""
'      end if
'			  	response.write "<td>" & oRequests("streetname")& lcl_valid_street & "</td>" & vbcrlf
'			end if
			
			response.write "</tr>" & vbcrlf
			
			
			if ReportType="Detail" or ReportType="DrillThru" or ReportType="responsedetail" then
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
		 
		if ReportType="Detail" or ReportType="DrillThru" or ReportType="responsedetail" then
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
				
				' SUB TOTALS FOR LAST GROUPING
				' DISPLAY SUBTOTAL ROW
				response.Write "<tr bgcolor=""#dddddd"" STYLE=""COLOR:NAVY;"">" & vbcrlf
				response.write "    <td align=""center""><b>Subtotals:</b></td>" & vbcrlf
				response.write "    <td> <b>Submitted: "              & iNumSubmittedRequests                         & "</td>" & vbcrlf
				response.write "    <td> <b>Not Viewed: "             & iNumSubmittedRequests - iNumResponsesViewed   & "</td>" & vbcrlf
				response.write "    <td colspan=""2""><b>No Action: " & iNumSubmittedRequests - iNumResponsesNoAction & "</td>" & vbcrlf
				response.write "    <td colspan=""7"">&nbsp;</td>" & vbcrlf
				response.write "</tr>" & vbcrlf

				If OrgHasFeature("issue location") Then
  					response.write "<td>&nbsp;</td>" & vbcrlf
				End If

				response.write "</tr>" & vbcrlf
		End If
  response.write "</table>" & vbcrlf
  response.write "<div>" & vbcrlf
  response.write "<table border=""0"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "<!--<img src=""../images/small_delete.gif"" align=""absmiddle"">&nbsp;<a href=""javascript:document.all.DelEvent.submit();"">DELETE</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->" & vbcrlf
  response.write "          <a href=""action_line_list_response.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=""0"" align=""absmiddle"" hspace=""3"" src=""../images/arrow_back.gif""></a>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <a href=""action_line_list_response.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          &nbsp;" & "<a href=""action_line_list_response.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <a href=""action_line_list_response.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=""0"" src=""../images/arrow_forward.gif"" valign=bottom></a>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>"
else
 	response.write "<p><b>No records found</p>" & vbcrlf
end if

End Function


'--------------------------------------------------------------------------------------------------
'  FUNCTION GETGROUPS(IUSERID)
'--------------------------------------------------------------------------------------------------
Function GetGroups(iUserID)

	sSQL = "SELECT Groups.GroupID, Users.OrgID, Groups.GroupName, Groups.GroupDescription "
 sSQL = sSQL & " FROM Users "
 sSQL = sSQL & " INNER JOIN UsersGroups ON Users.UserID = UsersGroups.UserID "
 sSQL = sSQL & " INNER JOIN  Groups ON UsersGroups.GroupID = Groups.GroupID "
 sSQL = sSQL & " WHERE (Groups.GroupType = 2) "
 sSQL = sSQL & " AND (Users.OrgID = " & Session("OrgID") & ") "
' sSQL = sSQL & " AND (Users.UserID = " & Session("UserID") & ") "
 sSQL = sSQL & " AND (Users.UserID = " & iUserID & ") "
 sSQL = sSQL & " ORDER BY Groups.GroupName "

	Set oDepts = Server.CreateObject("ADODB.Recordset")
	oDepts.Open sSQL, Application("DSN") , 3, 1

	If NOT oDepts.EOF Then

  		do while not oDepts.EOF
    			sReturnValue = sReturnValue & "'" & oDepts("GroupID") & "',"
    			oDepts.MoveNext
  		loop

  		sReturnValue = LEFT(sReturnValue,LEN(sReturnValue)-1) 

 else
    sReturnValue = 0

	End If 
			
	Set oDepts = Nothing

	GetGroups = sReturnValue

End Function

'--------------------------------------------------------------------------------------------------
'  FUNCTION FNLISTFORMS()
'--------------------------------------------------------------------------------------------------
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
					
					response.write "<option value=C" & oForms("form_category_id") & " " & selectA & ">----Category: " & sCurrentCategory & "</option>" & vbcrlf
				End If
	
				if selectFormId = "C" & oForms("form_category_id") & "" then
					selectA = "selected"
				else
					selectA = ""
				end if
				
				If (sCurrentCategory <> sLastCategory) AND (sLastCategory <> "NONE_START") Then
					response.write "<option value=C" & oForms("form_category_id") & " " & selectA & ">----Category: " & sCurrentCategory &  "</option>" & vbcrlf
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
	sSQL = "select groupid,orgid,groupname,groupdescription from groups where grouptype=2 AND orgid=" & Session("OrgID") & " AND groupid IN (" & GetGroups(session("userid")) & ") order by groupname"
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
		response.write "<option " & sSelected  & " value=" & oDepts("groupid") & ">" & oDepts("groupname") & "</option>" & vbcrlf

		oDepts.MoveNext
	Loop	

	' CLEAN UP OBJECTS
	Set oDepts = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' SUB DRAWASSIGNEDEMPLOYEESELECTION(IORGID)
'------------------------------------------------------------------------------------------------------------
Sub DrawAssignedEmployeeSelection(iorgid)

	sSQLassignedto = "SELECT FirstName + ' ' + LastName as assigned_Name, UserID "
 sSQLassignedto = sSQLassignedto & " FROM USERS "
 sSQLassignedto = sSQLassignedto & " WHERE OrgID = " & Session("OrgID")
 sSQLassignedto = sSQLassignedto & " ORDER BY FirstName,LastName"

	Set oAssigned = Server.CreateObject("ADODB.Recordset")
	oAssigned.Open sSQLassignedto, Application("DSN"), 3, 1

'IF THERE ARE ASSIGNED USERS THEN LIST
	If NOT oAssigned.EOF Then
 		'BEGIN SELECTION BOX
  		response.write "<select name=""selectAssignedto"">" & vbcrlf
    response.write "  <option value=""all"">Anyone</option>" & vbcrlf

 		'LOOP THRU ASSIGNED USERS
   	Do While NOT oAssigned.EOF		

   			'SET SELECT BOX TO DISPLAY CURRENTLY SELECTED NAME
       if selectAssignedto <> "all" then
       			If clng(selectAssignedto) = clng(oAssigned("userid")) then 
			   	      selectAssign = "selected"
       			Else
         				selectAssign = ""
       			End If
       else
          selectAssign = ""
       end if
			
    		'DISPLAY ASSIGNED EMPLOYEE AS OPTION
'    			response.write "<option value=""" & oAssigned("assigned_Name") & """ " & selectAssign & ">" & oAssigned("assigned_Name") & "</option>" & vbcrlf
    			response.write "<option value=""" & oAssigned("userid") & """ " & selectAssign & ">" & oAssigned("assigned_Name") & "</option>" & vbcrlf

    			oAssigned.MoveNext
   	Loop

  		response.write "</select>&nbsp;&nbsp;&nbsp;" & vbcrlf

  		oAssigned.Close

	End If

	Set oAssigned = Nothing

End Sub
%>

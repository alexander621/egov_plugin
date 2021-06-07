<!-- #include file="../includes/common.asp" //-->
<%
Response.Buffer      = True
Response.Expires     = -1
Server.ScriptTimeout = 600  'in secs.  10 min.

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: EVAL_EXPORT_FORM.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0 01/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/17/06	STEVE LOAR - SECURITY, HEADER AND NAV CHANGED
' 1.2 08/18/08 David Boyer - Added Status and Sub-Status are search criteria
' 1.3 06/23/09 David Boyer - Added date validation in search criteria
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../" ' Override of value from common.asp

'SECURITY CHEC
 if not UserHasPermission( Session("UserId"), "action export" ) then
 	  response.redirect sLevel & "permissiondenied.asp"
 end if

'GET SEARCH DATE RANGE
'CURRENT MONTH START TILL TODAY
 datStartDate = Month(Date()) & "/1/" & Year(Date())
 datEndDate   = Date()

'Check for org features
 lcl_orghasfeature_action_line_substatus = orghasfeature("action_line_substatus")

'Check for user permissions
 lcl_userhaspermission_action_line_substatus = userhaspermission(session("userid"),"action_line_substatus")
%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="JavaScript">
<!--
function doCalendar(ToFrom) {
		w = (screen.width - 350)/2;
		h = (screen.height - 350)/2;
		eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function validateFields() {
  var lcl_return_false = 0

		var daterege   = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var dateFromOk = daterege.test(document.getElementById("fromDate").value);
		var dateToOk   = daterege.test(document.getElementById("toDate").value);

  if (document.getElementById("toDate").value != "") {
   		if (! dateToOk ) {
         document.getElementById("toDate").focus();
         inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "Submitted To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
         lcl_return_false = lcl_return_false + 1;
     }else{
         clearMsg("toDateCalPop");
     }
  }else{
     document.getElementById("toDate").focus();
     inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Required Field Missing: </strong> Submitted To Date',10,'toDateCalPop');
     lcl_return_false = lcl_return_false + 1;
  }

  if (document.getElementById("fromDate").value != "") {
   		if (! dateFromOk ) {
         document.getElementById("fromDate").focus();
         inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "Submitted From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
         lcl_return_false = lcl_return_false + 1;
     }else{
         clearMsg("fromDateCalPop");
     }
  }else{
     document.getElementById("fromDate").focus();
     inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Required Field Missing: </strong> Submitted From Date',10,'fromDateCalPop');
     lcl_return_false = lcl_return_false + 1;
  }

  if(lcl_return_false > 0) {
     return false;
  }else{
     return true;
  }
}

function submitForm(p_version) {
  if(validateFields()) {
     var x=document.getElementById("csv_export")
     if (p_version=="2.0") {
         x.action="eval_export.asp";
     }else if(p_version=="3.0") {
         x.action="eval_export_dtb.asp";
     }else{
         x.action="eval_export_v1.asp";
     }

     x.submit()
  }
}
//-->
</script>
  <script language="Javascript" src="../reporting/scripts/dates.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
</head>

<body>
 
<%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">
<h3>Action Line CSV Data Export (<%=getEgovWebsiteURL%>)</h3>
<form name="frmPFilter" id="csv_export" action="" method="post">
<div class="shadow">
<table class="tablelist" cellpadding="5" cellspacing="0" style="padding-left:10px;">
  <tr><th class="corrections" align="left" colspan="2">&nbsp;Action Line CSV Data Export</th></tr>
		<tr><td>&nbsp;
		       	<!--BEGIN: SEARCH OPTIONS-->
      				<table border="0">
        				<tr>
                <td align="right"><strong>Select Form:</strong></td>
          						<td colspan="2"><% subDisplayFormSelection() %></td>
         			</tr>
       					<tr>
          						<td align="right"><strong>Submitted From:</strong></td>
    												<td>
    						   							<input type="text" name="fromDate" id="fromDate" value="<%=datStartDate%>" size="10" maxlength="10" onchange="clearMsg('fromDateCalPop');" />
    						   							<a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border="0" name="fromDateCalPop" id="fromDateCalPop" onclick="clearMsg('fromDateCalPop');" /></a>
    						   							&nbsp;<strong>To:</strong>
    						   							<input type="text" name="toDate" id="toDate" value="<%=datEndDate%>" size="10" maxlength="10" onchange="clearMsg('toDateCalPop');" />
    						   							<a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border="0" name="toDateCalPop" id="toDateCalPop" onclick="clearMsg('toDateCalPop');" /></a>
    												</td>
    												<td><%DrawDateChoices("Dates")%></td>
    								</tr>
          </table>
      				<table border="0" cellpadding="15">
       					<tr>
                <td>
                  <%
            				   'Retrieve all of the MAIN statuses.
                    sSQL1 = "SELECT action_status_id, status_name, display_order "
               					sSQL1 = sSQL1 & " FROM egov_actionline_requests_statuses "
               					sSQL1 = sSQL1 & " WHERE active_flag = 'Y' "
                    sSQL1 = sSQL1 & " AND parent_status = 'MAIN' "
               					sSQL1 = sSQL1 & " AND orgid IN (0," & session("orgid") & ") "
               					sSQL1 = sSQL1 & " ORDER BY display_order, status_name "

               					Set oMain = Server.CreateObject("ADODB.Recordset")
                    oMain.Open sSQL1, Application("DSN"), 3, 1

                    i = 0
               					if not oMain.eof then
                       lcl_display_substatus = "N"
                       lcl_label             = "Status"

                      'Check to see if the org and user have the "Action Line - Sub-Status" feature.
                       if lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then
                          lcl_display_substatus = "Y"
                          lcl_label             = lcl_label & "/Sub-Status"
                       end if

                       response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#c0c0c0"">" & vbcrlf
                       response.write "  <caption align=""left"" style=""font-size: 11px; font-family: verdana,tahoma;""><strong>" & lcl_label & ":</strong></caption>" & vbcrlf
                       response.write "  <tr valign=""top"">" & vbcrlf

                       lcl_parent_status = ""
               					   lcl_line_count    = 0

                				  'Loop through all of the Sub-Statuses
                				   while not oMain.eof
                          lcl_line_count = lcl_line_count + 1
                						    i = i + 1

                          lcl_click_status    = " checked=""checked"""
                          lcl_click_substatus = ""
             						       lcl_parent_status   = oMain("status_name")

                          response.write "      <td>" & vbcrlf
                          response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" bgcolor=""#ffffff"">" & vbcrlf
                          response.write "            <tr bgcolor=""#efefef"">" & vbcrlf
                          response.write "                <td colspan=""2"">" & vbcrlf
                          response.write "                    <input type=""checkbox"" name=""p_status_" & oMain("action_status_id") & """ id=""p_status_" & oMain("action_status_id") & """ value=""" & oMain("status_name") & """ " & lcl_click_status & ">" & vbcrlf
                          response.write "                    <strong>" & oMain("status_name") & "</strong>" & vbcrlf
                          response.write "                </td>" & vbcrlf
                          response.write "            </tr>" & vbcrlf

                          if lcl_display_substatus = "Y" then

                            'Retrieve all of the sub-statuses for the organization for each main status if any exist.
                             sSQL2 = "SELECT action_status_id, status_name, display_order, active_flag "
                             sSQL2 = sSQL2 & " FROM egov_actionline_requests_statuses "
                             sSQL2 = sSQL2 & " WHERE UPPER(parent_status) = '" & UCASE(oMain("status_name")) & "' "
                             sSQL2 = sSQL2 & " AND active_flag = 'Y' "
                             sSQL2 = sSQL2 & " AND orgid = " & session("orgid")
                             sSQL2 = sSQL2 & " ORDER BY display_order, status_name "

                             Set oSub = Server.CreateObject("ADODB.Recordset")
                             oSub.Open sSQL2, Application("DSN"), 3, 1

                             if not oSub.eof then
                                while not oSub.eof
                                   response.write "            <tr bgcolor=""#ffffff"">" & vbcrlf
                                   response.write "                <td colspan=""2"">" & vbcrlf
                                   response.write "                    <input type=""checkbox"" name=""p_substatus_" & oSub("action_status_id") & """ id=""p_substatus_" & oSub("action_status_id") & """ value=""" & oSub("action_status_id") & """ " & lcl_click_substatus & "" & vbcrlf
                                   response.write "                    <strong>" & oSub("status_name") & "</strong>" & vbcrlf
                                   response.write "                </td>" & vbcrlf
                                   response.write "            </tr>" & vbcrlf
                                   oSub.movenext
                                wend
                             end if

                             set oSub = nothing

                          end if

                          response.write "          </table>" & vbcrlf
                          response.write "      </td>" & vbcrlf

                   				   oMain.movenext
                				   wend

                       oMain.close
                       set oMain = nothing

                       response.write "  </tr>" & vbcrlf
                       response.write "</table>" & vbcrlf
                    end if
                  %>
                </td>
            </tr>
          </table>
      				<table border="0">
       					<tr>
   				       		<td align="right">
                    <input type="button" value="Download CSV File" class="button" onClick="submitForm('1.0')" />
                    <input type="button" value="Download CSV File (expanded)" class="button" onClick="submitForm('2.0')" />
                    <!--
                    <input type="button" value="Download CSV File" class="button" onClick="javascript:submitForm('1.0')" />
                    <input type="button" value="Download CSV File (expanded)" class="button" onClick="javascript:submitForm('2.0')" />
                    -->
                </td>
       					</tr>
      				</table>
        		</p>
       			<!--END: SEARCH OPTIONS-->
      </td></tr>
</table>
</div>
</form>
  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
Sub subDisplayFormSelection()
	Dim sSql, iOrgId, oFormList

	iorgID = session("orgid")
	sSQL = "SELECT action_form_id, action_form_name, action_form_type "
 sSQL = sSQL & " FROM egov_action_request_forms "
 sSQL = sSQL & " WHERE (action_form_type <> 2) "
 sSQL = sSQL & " AND orgid = " & iorgID
 sSQL = sSQL & " ORDER BY action_form_type, action_form_name"
	
	Set oFormList = Server.CreateObject("ADODB.Recordset")
	oFormList.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oFormList.EOF Then
		
  		response.write "<select name=""iformid"">" & vbcrlf
		  while NOT oFormList.eof
		 	  'DISPLAY SELECT OPTION
  			  response.write "  <option value=""" & oFormList("action_form_id") & """>" & UCASE(oFormList("action_form_name")) & "</option>" & vbcrlf
		  	  oFormList.movenext
  		wend
  		response.write "</select>" & vbcrlf
	
	End If
	oFormList.close
	Set oFormList = Nothing 

End Sub

'------------------------------------------------------------------------------
Function DrawDateChoices(sName)

	response.write "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=""calendarinput"" Name=""" & sName & """>" & vbcrlf
	response.write "  <option value=""0"">Or Select Date Range from Dropdown...</option>" & vbcrlf
	response.write "  <option value=""11"">This Week</option>" & vbcrlf
	response.write "  <option value=""12"">Last Week</option>" & vbcrlf
	response.write "  <option value=""1"">This Month</option>" & vbcrlf
	response.write "  <option value=""2"">Last Month</option>" & vbcrlf
	response.write "  <option value=""3"">This Quarter</option>" & vbcrlf
	response.write "  <option value=""4"">Last Quarter</option>" & vbcrlf
	response.write "  <option value=""6"">Year to Date</option>" & vbcrlf
	response.write "  <option value=""5"">Last Year</option>" & vbcrlf
	response.write "  <option value=""7"">All Dates to Date</option>" & vbcrlf
	response.write "</select>" & vbcrlf

End Function
%>
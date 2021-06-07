<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="subscriptionslog_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: subscriptionslog_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the rss feeds
'
' MODIFICATION HISTORY
' 1.0 06/29/09 David Boyer - Initial Version
' 1.1 08/05/09 David Boyer - Added "listtype"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptionslog_maint") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if request("listtype") <> "" then
    lcl_list_type = ucase(request("listtype"))
 else
    lcl_list_type = ""
 end if

 if lcl_list_type = "BID" then
    lcl_pagetitle      = "Bid Postings"
    lcl_userpermission = "subscriptionslog_maint_bids"
 elseif lcl_list_type = "JOB" then
    lcl_pagetitle      = "Job Postings"
    lcl_userpermission = "subscriptionslog_maint_jobs"
 else
    lcl_pagetitle      = "Distribution Lists"
    lcl_userpermission = "subscriptionslog_maint"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Retrieve the search options
 today                  = Date()
 lcl_sc_fromdate        = request("fromDate")
 lcl_sc_todate          = request("toDate")
 lcl_sc_sentbyuserid    = request("sc_sentbyuserid")
 lcl_sc_email_fromname  = request("sc_email_fromname")
 lcl_sc_email_fromemail = request("sc_email_fromemail")
 lcl_sc_email_subject   = request("sc_email_subject")
 lcl_sc_email_format    = request("sc_email_format")

'From Date (last year)
 if lcl_sc_fromdate = "" or IsNull(lcl_sc_fromdate) then
    lcl_sc_fromdate = dateAdd("yyyy",-1,today)
 end if

'To Date (get today's date)
 if lcl_sc_todate = "" or IsNull(lcl_sc_todate) then
    lcl_sc_todate = dateAdd("d",0,today)
 end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%> - Send Log}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
 	<script language="javascript" src="../scripts/getdates.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--
function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("../calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function confirm_delete(idl_logid, iTotalItems) {
  var lcl_rssTitle = document.getElementById("logfile"+idl_logid).innerHTML;

  if(iTotalItems > 0) {
     lcl_msg  = '"' + lcl_rssTitle + '" cannot be deleted as there are RSS Items associated to it.\n';
     lcl_msg += 'Set the RSS Feed to "inactive".';

     alert(lcl_msg);
  }else{
    	if (confirm("Are you sure you want to delete '" + lcl_rssTitle + "' ?")) { 
  	   			//DELETE HAS BEEN VERIFIED
   		  		location.href='subscriptionslog_action.asp?user_action=DELETE&dl_logid=' + idl_logid + '&listtype=<%=lcl_list_type%>';
     }
		}
}

function validateFields() {

  var lcl_return_false = 0;
		var daterege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;

		var dateToOk   = daterege.test(document.getElementById("toDate").value);
		var dateFromOk = daterege.test(document.getElementById("fromDate").value);

		if (! dateToOk ) {
      document.getElementById("toDate").focus();
      inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
      lcl_return_false = lcl_return_false + 1;
  }else{
      clearMsg("toDateCalPop");
  }

		if (! dateFromOk ) {
      document.getElementById("fromDate").focus();
      inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
      lcl_return_false = lcl_return_false + 1;
  }else{
      clearMsg("fromDateCalPop");
  }

  if(lcl_return_false > 0) {
     return false;
  }else{
     document.getElementById("searchLog").submit();
     return true;
  }

}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;width:900px;">
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr>
                  <td><font size="+1"><strong><%=Session("sOrgName")%>&nbsp;<%=lcl_pagetitle%> - Send Log</strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
              </tr>
              <tr>
                  <td colspan="2">
                      <fieldset>
                        <legend>Search Options&nbsp;</legend>
                        <table border="0" cellspacing="0" cellpadding="2" style="margin-top:5px;">
                          <form name="searchLog" id="searchLog" method="post" action="subscriptionslog_list.asp">
                            <input type="hidden" name="listtype" id="listtype" value="<%=lcl_list_type%>" size="5" maxlength="20" />
                          <tr>
                              <td nowrap="nowrap"><strong>Sent Date</strong>&nbsp;&nbsp;From: </td>
                              <td>
                                  <input type="text" name="fromDate" id="fromDate" size="10" maxlength="10" value="<%=lcl_sc_fromdate%>" onchange="clearMsg('fromDateCalPop');" />
                                  <a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" id="fromDateCalPop" border="0" onclick="clearMsg('fromDateCalPop');" /></a>
                              </td>
                              <td nowrap="nowrap">
                                  To: <input type="text" name="toDate" id="toDate" size="10" maxlength="10" value="<%=lcl_sc_todate%>" onchange="clearMsg('toDateCalPop');" />
                                  <a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" id="toDateCalPop" border="0" onclick="clearMsg('toDateCalPop');" /></a>
                              </td>
                              <td colspan="2">
                                  <% DrawDateChoices "Date", "" %>
                              </td>
                          </tr>
                          <tr>
                              <td nowrap="nowrap">Sent By:</td>
                              <td colspan="2">
                                  <select name="sc_sentbyuserid" id="sc_sentbyuserid">
                                    <option value=""></option>
                                    <% displaySearchCriteriaSentBy session("orgid"), lcl_sc_sentbyuserid %>
                                  </select>
                              </td>
                              <td nowrap="nowrap">Email Format:</td>
                              <td>
                                  <select name="sc_email_format" id="sc_email_format">
                                    <% displaySearchCriteriaEmailFormats lcl_sc_email_format %>
                                  </select>
                              </td>
                          </tr>
                          <tr>
                              <td>From Name: </td>
                              <td colspan="2"><input type="text" name="sc_email_fromname" size="40" maxlength="150" value="<%=lcl_sc_email_fromname%>" /></td>
                              <td nowrap="nowrap">From Email:</td>
                              <td><input type="text" name="sc_email_fromemail" size="40" maxlength="150" value="<%=lcl_sc_email_fromemail%>" /></td>
                          </tr>
                          <tr>
                              <td>Subject: </td>
                              <td colspan="4"><input type="text" name="sc_email_subject" size="80" maxlength="150" value="<%=lcl_sc_email_subject%>" /></td>
                          </tr>
                          <tr>
                              <td colspan="5"><input type="button" name="searchButton" id="searchButton" value="Search" class="button" onclick="validateFields();" /></td>
                          </tr>
                          </form>
                        </table>
                      </fieldset>
                  </td>
              </tr>
            </table>
            <br />
            <p><% listLogFile session("orgid"), lcl_list_type, lcl_sc_fromdate, lcl_sc_todate, lcl_sc_sentbyuserid, _
                              lcl_sc_email_fromname, lcl_sc_email_fromemail, lcl_sc_email_subject, lcl_sc_email_format %></p>
          </div>
      </td>
  </tr>
</table>

  </div>
</div>
	
<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
sub listLogFile(iorgid, iListType, iSCFromDate, iSCToDate, iSCSentByUserID, iSCFromName, iSCFromEmail, iSCSubject, iSCEmail_Format)
 	Dim iRowCount

 	iRowCount = 0

  if iListType <> "" then
     sListType = "'" & UCASE(dbsafe(iListType)) & "'" 
  else
     sListType = "NULL"
  end if

 'Used the SQL BETWEEN function which also performs a LESS THAN on the "todate"
 'Adding a day to the "toDate" returns the information INCLUDING the "toDate"
  if iSCToDate <> "" then
     lcl_query_toDate = dateAdd("d",1,iSCToDate)
  end if

 'Determine if the user is a "root admin"
  lcl_isRootAdmin = false

  if UserIsRootAdmin(session("userid")) then
     lcl_isRootAdmin = true
  end if

  sSQL = "SELECT dl_logid, "
  sSQL = sSQL & " sentbyuserid, "
  sSQL = sSQL & " sentdate, "
  sSQL = sSQL & " completedate, "
  sSQL = sSQL & " sendstatus, "
  sSQL = sSQL & " email_fromname, "
  sSQL = sSQL & " email_fromemail, "
  sSQL = sSQL & " email_subject, "
  sSQL = sSQL & " email_body, "
  sSQL = sSQL & " email_format, "
  sSQL = sSQL & " dl_listids, "
  sSQL = sSQL & " (select firstname + ' ' + lastname "
  sSQL = sSQL &  " from users "
  sSQL = sSQL &  " where userid = sentbyuserid) AS sentbyusername "
  sSQL = sSQL & " FROM egov_class_distributionlist_log "
  sSQL = sSQL & " WHERE orgid = " & iorgid

  if iListType <> "" then
     sSQL = sSQL & " AND UPPER(distributionlisttype) = " & sListType
  else
     sSQL = sSQL & " AND (distributionlisttype IS NULL OR distributionlisttype = '') "
  end if

 'Sent Date
  if iSCFromDate <> "" AND iSCToDate <> "" then
     sSQL = sSQL & " AND sentdate BETWEEN '" & iSCFromDate & "' AND '" & lcl_query_toDate & "' "
  else
     if iSCFromDate <> "" AND iSCToDate = "" then
        sSQL = sSQL & " AND sentdate >= '" & iSCFromDate & "' "
     elseif iSCToDate <> "" AND iSCFromDate = "" then
        sSQL = sSQL & " AND sentdate < '" & lcl_query_toDate & "' "
     end if
  end if

 'Sent By User ID
  if trim(iSCSentByUserID) <> "" then
     sSQL = sSQL & " AND sentbyuserid = " & iSCSentByUserID
  end if

 'From Name
  if trim(iSCFromName) <> "" then
     sSQL = sSQL & " AND UPPER(email_fromname) LIKE ('%" & UCASE(dbsafe(trim(iSCFromName))) & "%') "
  end if

 'From Email
  if trim(iSCFromEmail) <> "" then
     sSQL = sSQL & " AND UPPER(email_fromemail) LIKE ('%" & UCASE(dbsafe(trim(iSCFromEmail))) & "%') "
  end if

 'Subject
  if trim(iSCSubject) <> "" then
     sSQL = sSQL & " AND UPPER(email_subject) LIKE ('%" & UCASE(dbsafe(trim(iSCSubject))) & "%') "
  end if

 'Email Format
  if iSCEmail_Format <> "" then
     sSQL = sSQL & " AND UPPER(email_format) = '" & UCASE(iSCEmail_Format) & "' "
  end if

  sSQL = sSQL & " ORDER BY sentdate DESC "

 	set oListLog = Server.CreateObject("ADODB.Recordset")
	 oListLog.Open sSQL, Application("DSN"), 3, 1
	
 	if not oListLog.eof then
   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""3"" class=""tablelist"" border=""0"" style=""width:900px"">" & vbcrlf
   		response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <th nowrap=""nowrap"">Sent Date</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"">Sent By</th>" & vbcrlf
     'response.write "      <th nowrap=""nowrap"">Email Format</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"">From Name</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"">From Email</th>" & vbcrlf
     response.write "      <th>Subject</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"">Completed Date</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     lcl_bgcolor             = "#ffffff"

     do while not oListLog.eof
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount    = iRowCount + 1

       'Setup the onclick
        lcl_row_onclick = "location.href='subscriptionslog_maint.asp?dl_logid=" & oListLog("dl_logid") & "&listtype=" & iListType & "';"

       'Determine what the email format is
        lcl_emailformat = getEmailFormatDesc(oListLog("email_format"))

       'Set up the "sent by username"
        lcl_sentbyusername = trim(oListLog("sentbyusername"))

        if lcl_sentbyusername = "" then
           lcl_sentbyusername = "<span style=""color:#800000;"">N/A</span>"
        end if

        if lcl_isRootAdmin then
           lcl_sentbyusername = "[<a href=""../dirs/update_user.asp?userid=" & oListLog("sentbyuserid") & "&currentpage=1"" target=""_blank"">" & oListLog("sentbyuserid") & "</a>] " & lcl_sentbyusername
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" nowrap=""nowrap"" onClick=""" & lcl_row_onclick & """><span id=""logfile" & oListLog("dl_logid") & """>" & oListLog("sentdate") & "</span></td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" nowrap=""nowrap"">" & lcl_sentbyusername & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" nowrap=""nowrap"" onClick=""" & lcl_row_onclick & """>" & oListLog("email_fromname")  & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" nowrap=""nowrap"" onClick=""" & lcl_row_onclick & """>" & oListLog("email_fromemail") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" nowrap=""nowrap"" onClick=""" & lcl_row_onclick & """>" & oListLog("email_subject")   & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" nowrap=""nowrap"" onClick=""" & lcl_row_onclick & """>" & oListLog("completedate")    & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" align=""center"">" & vbcrlf
        response.write "          <input type=""button"" name=""delete" & iRowCount & """ id=""delete" & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oListLog("dl_logid") & "');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

       oListLog.movenext
   loop

 		response.write "</table>" & vbcrlf
	  response.write "</div>" & vbcrlf

 else
  		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No log entries have been created.</p>" & vbcrlf
	end if

	oListLog.close
	set oListLog = nothing 

end sub

'------------------------------------------------------------------------------
function getTotalRSSItems(idl_logid)

  lcl_return = 0

  if idl_logid <> "" then
     sSQL = "SELECT count(rssid) AS total_rss "
     sSQL = sSQL & " FROM egov_rss "
     sSQL = sSQL & " WHERE orgid = " & session("orgid")
     sSQL = sSQL & " AND dl_logid = "  & idl_logid

    	set oRSSCnt = Server.CreateObject("ADODB.Recordset")
   	 oRSSCnt.Open sSQL, Application("DSN"), 3, 1

     if not oRSSCnt.eof then
        lcl_return = oRSSCnt("total_rss")
     end if

     oRSSCnt.close
     set oRSSCnt = nothing

  end if

  getTotalRSSItems = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "RSS Feed does not exist..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub displaySearchCriteriaSentBy(iOrgID, iSentByID)

  sSQL = "SELECT distinct sentbyuserid, "
  sSQL = sSQL & " (select firstname from users where userid = sentbyuserid) AS sentbyuserfirstname, "
  sSQL = sSQL & " (select lastname from users where userid = sentbyuserid) AS sentbyuserlastname "
  sSQL = sSQL & " FROM egov_class_distributionlist_log "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " ORDER BY 3, 2 "

 	set oSCSentByIDs = Server.CreateObject("ADODB.Recordset")
	 oSCSentByIDs.Open sSQL, Application("DSN"), 3, 1

  if not oSCSentByIDs.eof then
     do while not oSCSentByIDs.eof
        if iSentByID <> "" then
           if CLng(iSentByID) = CLng(oSCSentByIDs("sentbyuserid")) then
              lcl_selected_sentbyuserid = " selected=""selected"""
           else
              lcl_selected_sentbyuserid = ""
           end if
        else
           lcl_selected_sentbyuserid = ""
        end if

        response.write "  <option value=""" & oSCSentByIDs("sentbyuserid") & """" & lcl_selected_sentbyuserid & ">" & oSCSentByIDs("sentbyuserfirstname") & " " & oSCSentByIDs("sentbyuserlastname") & "</option>" & vbcrlf

        oSCSentByIDs.movenext
     loop
  end if

  oSCSentByIDs.close
  set oSCSentByIDs = nothing

end sub

'------------------------------------------------------------------------------
sub displaySearchCriteriaEmailFormats(iSelectedEmailFormat)

  lcl_selected_none      = " selected=""selected"""
  lcl_selected_plaintext = ""
  lcl_selected_htmlplain = ""
  lcl_selected_html      = ""

  if iSelectedEmailFormat <> "" then
     lcl_emailformat = CStr(iSelectedEmailFormat)

     if lcl_emailformat = "1" then
        lcl_selected_plaintext = " selected=""selected"""
     elseif lcl_emailformat = "2" then
        lcl_selected_htmlplain = " selected=""selected"""
     elseif lcl_emailformat = "3" then
        lcl_selected_html = " selected=""selected"""
     end if
  end if

  response.write "  <option value="""""  & lcl_selected_none      & "></option>" & vbcrlf
  response.write "  <option value=""1""" & lcl_selected_plaintext & ">Plain Text Only</option>"     & vbcrlf
  response.write "  <option value=""2""" & lcl_selected_htmlplain & ">HTML And Plain Text</option>" & vbcrlf
  response.write "  <option value=""3""" & lcl_selected_html      & ">HTML Only</option>"           & vbcrlf

end sub
%>
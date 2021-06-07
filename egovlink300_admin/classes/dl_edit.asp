<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DL_LIST.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 05/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  05/10/06  JOHN STULLENBERGER - INITIAL VERSION
' 1.2	 10/05/06	 Steve Loar - Security, Header and nav changed
' 1.3  01/28/08  David Boyer - Incorporated isFeatureOffline check
' 1.4  01/28/08  David Boyer - Incorporated Job/Bid Postings
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptions,job_postings,bid_postings") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

'INITIALIZE VARIABLES
 Dim sName, sDescription,blnDisplay
 Dim iDLid

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=hide

'Check the type of list and then check for the permission
 if UCASE(request("sc_list_type")) = "JOB" then
    lcl_permission = "job_postings"
 elseif UCASE(request("sc_list_type")) = "BID" then
    lcl_permission = "bid_postings"
 else
    lcl_permission = "distribution lists"
 end if

 if not UserHasPermission( Session("UserId"), lcl_permission ) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the search parameters
 lcl_sc_name              = request("sc_name")
 lcl_sc_description       = request("sc_description")
 lcl_sc_publicly_viewable = request("sc_publicly_viewable")
 lcl_sc_show_postings     = request("sc_show_postings")
 lcl_sc_list_type         = request("sc_list_type")
 lcl_sc_orderby           = request("sc_orderby")

'GET DL ID
if request("dlid") = "" OR NOT isnumeric(request("dlid")) OR request("dlid") = 0 then
	 '-- ADD --------------------------------
	  idlID = 0

   if lcl_sc_list_type = "JOB" then
      sTitle    = "Add New Job Posting Category"
      lcl_label = "Job Posting Category"
   elseif lcl_sc_list_type = "BID" then
      sTitle    = "Add New Bid Posting Category/Sub-Category"
      lcl_label = "Bid Posting Category/Sub-Category"
   else
   	  sTitle    = "Add New Distribution List"
      lcl_label = "Distribution List"
   end if

	  sLinkText = "Add"
else
	 '-- EDIT -------------------------------
	  idlID = request("dlid")
   if lcl_sc_list_type = "JOB" then
      sTitle    = "Edit Job Posting Category"
      lcl_label = "Job Bosting"
   elseif lcl_sc_list_type = "BID" then
      sTitle    = "Edit Bid Posting Category/Sub-Category"
      lcl_label = "Bid Posting"
   else
      sTitle    = "Edit Distribution List"
      lcl_label = "Distribution List"
   end if

	  sLinkText = "Save"
end if

'Retrieve data for this list
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_class_distributionlist "
 sSQL = sSQL & " WHERE distributionlistid = " & idlID

	set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN") , 3, 1

	If NOT oValues.EOF Then
  		sName        = oValues("distributionlistname")
		  sDescription = oValues("distributionlistdescription")
  		blnDisplay   = oValues("distributionlistdisplay")
    sParentID    = oValues("parentid")
 else
    sName        = ""
    sDescription = ""
    blnDisplay   = False
    sParentID    = ""
	End If

	oValues.close
	Set oValues = nothing

If blnDisplay Then
  	blnDisplay = " CHECKED "
Else
  	blnDisplay = " "
End If

'Build the BODY onload
 lcl_onload = ""
 lcl_onload = lcl_onload & "setMaxLength();"
%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="javascript" src="tablesort.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>

<style>
 		input {width:300px;}
</style>
</head>
<body onload="<%=lcl_onload%>">
<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
<font size="+1"><b><%=sTitle%></b></font><br>
<a href="dl_mgmt.asp?sc_name=<%=lcl_sc_name%>&sc_description=<%=lcl_sc_description%>&sc_publicly_viewable=<%=lcl_sc_publicly_viewable%>&sc_show_postings=<%=lcl_sc_show_postings%>&sc_list_type=<%=lcl_sc_list_type%>&sc_orderby=<%=lcl_sc_orderby%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
<p>

<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr>
      <td>
          <div id="functionlinks">
          		<a href="dl_mgmt.asp?sc_name=<%=lcl_sc_name%>&sc_description=<%=lcl_sc_description%>&sc_publicly_viewable=<%=lcl_sc_publicly_viewable%>&sc_show_postings=<%=lcl_sc_show_postings%>&sc_list_type=<%=lcl_sc_list_type%>&sc_orderby=<%=lcl_sc_orderby%>"><img src="../images/cancel.gif" align="absmiddle" border="0">&nbsp;Cancel</a>&nbsp;&nbsp;
          		<a href="javascript:document.frmdl.submit();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;<%=sLinkText%></a>&nbsp;&nbsp;
          </div>
      </td>
      <td align="right">
      <%
        lcl_message = ""

        if request("success") = "SU" then
           lcl_message = "<b style=""color:#FF0000"">*** Successfully Updated... ***</b>"
        elseif request("success") = "EU" then
           lcl_message = "<b style=""color:#FF0000"">*** This Bid Posting has sub-categories associated to it.  Category cannot be modified. ***</b>"
        else
           lcl_message = "&nbsp;"
        end if

        if lcl_message <> "" then
           response.write lcl_message
        end if
      %>
      </td>
  </tr>
</table>
<div class="shadow">
<table cellpadding="5" cellspacing="0" border="0" class="locationlist">
  <tr><th><%=lcl_label%> Information</th></tr>
  <tr>
      <td>
          <table border="0" cellpadding="5" cellspacing="0">
            <form name="frmdl" action="dl_save.asp" method="post">
              <input type="<%=lcl_hidden%>" name="idlid" value="<%=idlID%>" />
              <input type="<%=lcl_hidden%>" name="sc_name" value="<%=lcl_sc_name%>" size="15" maxlength="512">
              <input type="<%=lcl_hidden%>" name="sc_descripion" value="<%=lcl_sc_description%>" size="15" maxlength="1024">
              <input type="<%=lcl_hidden%>" name="sc_publicly_viewable" value="<%=lcl_sc_publicly_viewable%>" size="15" maxlength="5">
              <input type="<%=lcl_hidden%>" name="sc_show_postings" value="<%=lcl_sc_show_postings%>" size="15" maxlength="5">
              <input type="<%=lcl_hidden%>" name="sc_list_type" value="<%=lcl_sc_list_type%>" size="15" maxlength="100">
              <input type="<%=lcl_hidden%>" name="sc_orderby" value="<%=lcl_sc_orderby%>" size="15" maxlength="50">
          		<tr>
          	   		<td colspan="2">
                				<table border="0" cellspacing="0" cellpadding="0">
                 					<tr>
                    						<td>Name:</td>
                          <td><input type="text" name="sName" maxlength="150" value="<%=sName%>" ></td>
                 					</tr>
                      <%
                        'Only display this option if the listtype = "BID"
                         if lcl_sc_list_type = "BID" then
                      %>
                      <tr>
                          <td>Category:</td>
                          <td>
                              <select name="sParentID">
                                <option value=""></option>
                              <%
                               'Retreive all bid postings that DO NOT have a parentid
                               	sSQLs = "SELECT distributionlistid, distributionlistname "
                                sSQLs = sSQLs & " FROM egov_class_distributionlist "
                                sSQLs = sSQLs & " WHERE orgid = "  & session("orgid")
                                if clng(idlID) > 0 then
                                   sSQLs = sSQLs & " AND distributionlistid <> " & idlID
                                end if
                                sSQLs = sSQLs & " AND (parentid IS NULL OR parentid < 0) "
                                sSQLs = sSQLs & " AND distributionlisttype = 'BID' "
                                sSQLs = sSQLs & " ORDER BY UPPER(distributionlistname) "

                                Set rss = Server.CreateObject("ADODB.Recordset")
                                rss.Open sSQLs, Application("DSN"), 3, 1

                                if not rss.eof then
                                   while not rss.eof
                                      if isnull(sParentID) then
                                         lcl_selected = ""
                                      else
                                         if sParentID = rss("distributionlistid") then
                                            lcl_selected = " selected"
                                         else
                                            lcl_selected = ""
                                         end if
                                      end if

                                      response.write "  <option value=""" & rss("distributionlistid") & """" & lcl_selected & ">" & rss("distributionlistname") & "</option>" & vbcrlf
                                      rss.movenext
                                   wend
                                end if
                              %>
                              </select>
                          </td>
                      </tr>
                      <% else %>
                      <input type="<%=lcl_hidden%>" name="sParentID" value="<%=sParentID%>" size="3" maxlength="5">
                      <% end if %>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    Description:<br>
                    <textarea name="sDescription" maxlength="1024"><%=sDescription%></textarea>
                </td>
            </tr>
        		  <tr>
                <td>Display on Registration Form:</td>
                <td><input type="checkbox" name="blnDisplay" value="1" <%=blnDisplay%>></td>
            </tr>
            </form>
          </table>
      </td>
  </tr>
</table>
</div>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

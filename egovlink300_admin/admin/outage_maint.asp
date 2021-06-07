<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  outage_maint.asp
' AUTHOR:    David Boyer
' CREATED:   01/21/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  To allow admin(s) to turn of high-level features for code promotions.
'
' MODIFICATION HISTORY
' 1.0   01/21/08    David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
  	response.redirect "../default.asp"
End If 

if request("action") = "SAVE" then
  'Update the offline status of any/all features selected.
   for e = 1 to CLng(request("total_features"))
       if request.form("feature_offline_"&e) = "Y" then
          lcl_offline_status = "Y"
       else
          lcl_offline_status = "N"
       end if

       sSQLu = "UPDATE egov_organization_features "
       sSQLu = sSQLu & " SET feature_offline = '" & lcl_offline_status & "' "
       sSQLu = sSQLu & " WHERE featureid = " & request("featureid_" & e)

  		   set rsu = Server.CreateObject("ADODB.Recordset")
       rsu.Open sSQLu, Application("DSN"),3,1
   next

  'Update the outage message
   sSQLm = "UPDATE organizations "
   sSQLm = sSQLm & " SET outage_message = '" & replace(request("outage_message"),"'","''") & "' "

   set rsm = Server.CreateObject("ADODB.Recordset")
   rsm.Open sSQLm, Application("DSN"),3,1

   response.redirect "outage_maint.asp?success=SU"
end if

lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
%>
<html>
<head>
  <title>Outage Maintenance</title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<script language="javascript" src="../scripts/modules.js"></script>

<script language="javascript">
<!--
function doCalendar() {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function storeCaret (textEl) {
  if (textEl.createTextRange) {
      textEl.caretPos = document.selection.createRange().duplicate();
  }
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
      var caretPos = textEl.caretPos;
      caretPos.text =
      caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
      text + ' ' : text;
  }else{
      textEl.value  = text;
  }
}

function doPicker(sFormField) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

//function checkFieldLength(p_value,p_limit) {
//  lcl_length = p_value.length;
//  if(lcl_length <= p_limit) {
//     document.getElementById('message_char_cnt').innerHTML = p_limit + " character limit.  Characters remaining: " + (p_limit - lcl_length);
//  } else {
//     document.getElementById("outage_message").value = document.getElementById("control_field").value.substr(0,p_limit);
//     alert("Cannot exceed " + p_limit + " characters.");
//  }
//}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="checkFieldLength(document.getElementById('outage_message').value,4000,'Y',document.getElementById('outage_message'))">

<% ShowHeader sLevel %>
<!-- #include file="../menu/menu.asp"--> 
<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<form name="outage_maint" action="outage_maint.asp?action=SAVE" method="post">
    <input type="<%=lcl_hidden%>" name="control_field" id="control_field" value="" size="20" maxlength="4001">
  <caption align="left"><font size="+1"><b>Outage Maintenance</b></font></caption>
  <tr style="font-size:10px">
      <td>
          <img src="../images/cancel.gif" align="absmiddle" onclick="location.href='../default.asp'" style="cursor: hand">&nbsp;
      		  <a href="../default.asp"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;
      		  <img src="../images/edit.gif" align="absmiddle">&nbsp;
      		  <a href="javascript:document.outage_maint.submit();">Save</a>
    		</td>
      <td align="right">
      <%
        lcl_message = ""

        if request("success") = "SU" then
           lcl_message = "<b style=""color:#FF0000"">*** Successfully Updated... ***</b>"
        else
           lcl_message = "&nbsp;"
        end if

        if lcl_message <> "" then
           response.write lcl_message
        end if
      %>
      </td>
  </tr>
  <tr valign="top">
      <td width="50%">
          <div class="shadow">
        		<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
        		  <tr>
		              <th align="left" colspan="2">Feature Name</th>
            </tr>
            <%
      				   'Retrieve all of the sub-statuses for the organization for each parent_status
              sSQL = "SELECT featureid, featurename, feature_offline "
         					sSQL = sSQL & " FROM egov_organization_features "
         					sSQL = sSQL & " WHERE parentfeatureid = 0 "
         					sSQL = sSQL & " ORDER BY admindisplayorder "

         					set rs = Server.CreateObject("ADODB.Recordset")
              rs.Open sSQL, Application("DSN"), 3, 1

         					if not rs.eof then
                 i = 0
                 lcl_bgcolor = "#eeeeee"
                 lcl_checked = ""

         					  'Loop through all of the Organization Features
         					   while not rs.eof
                    i = i + 1
                    if rs("feature_offline") = "Y" then
                       lcl_checked = "checked"
                    else
                       lcl_checked = ""
                    end if

                    response.write "            <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
             	  	   response.write "                <td width=""20"">" & vbcrlf
                    response.write "                    <input type=""" & lcl_hidden & """ name=""featureid_" & i & """ value=""" & rs("featureid") & """ size=""3"" maxlength=""10"">" & vbcrlf
                    response.write "                    <input type=""checkbox"" name=""feature_offline_" & i & """ value=""Y"" " & lcl_checked & ">" & vbcrlf
                    response.write "                </td>" & vbcrlf
                    response.write "                <td>" & rs("featurename") & "</td>" & vbcrlf
                    response.write "            </tr>" & vbcrlf

                    lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")

            						  rs.movenext
            		   wend
              else
                 response.write "            <tr><td colspan=""2"">No Features Exist</td></tr>" & vbcrlf
            		end if
            %>
        		</table>
          </div>
          <input type="<%=lcl_hidden%>" name="total_features" value="<%=i%>" size="3" maxlength="10">
      </td>
      <td>
          <div class="shadow">
          <table border="0" cellspacing="0" cellpadding="5" class="tableadmin">
            <tr><th>Outage Message</th></tr>
            <%
              'Set up the default outage message.
               lcl_default_message = "The feature(s) [FEATURE NAME(S)] will be unavailable between<br>"
               lcl_default_message = lcl_default_message & " <strong>[START DATE/TIME]</strong> to <strong>[END DATE/TIME]</strong><br>"
               lcl_default_message = lcl_default_message & " due to [scheduled maintenance outage] or [code promotion]."

              'Retrieve the outage message
               sSQL1 = "SELECT outage_message FROM organizations WHERE orgid = " & session("orgid")

          					set rs1 = Server.CreateObject("ADODB.Recordset")
               rs1.Open sSQL1, Application("DSN"), 3, 1

               if not rs1.eof then
                  if rs1("outage_message") <> ""then
                     lcl_outage_message = rs1("outage_message")
                  else
                     lcl_outage_message = lcl_default_message
                  end if
               else
                  lcl_outage_message = lcl_default_message
               end if
            %>
            <tr><td>
                    <textarea name="outage_message" cols="60" rows="15" onkeydown="document.getElementById('control_field').value=this.value;" onkeyup="javascript:checkFieldLength(this.value,4000,'Y',this)"><%=lcl_outage_message%></textarea>
                    <div align="right" id="message_char_cnt">4000 character limit.  Characters remaining: 4000</div>
                </td></tr>
          </table>
          </div>
      </td>
  </tr>
  </form>
</table>
  </div>
</div>

<!-- #include file="../admin_footer.asp"-->  

</body>
</html>

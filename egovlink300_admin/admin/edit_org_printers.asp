<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: edit_org_printers.asp
' AUTHOR:   David Boyer
' CREATED:  04/23/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This screen maintains the printers available for the membership cards
'
' MODIFICATION HISTORY
' 1.0  04/23/08  David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
'if isFeatureOffline("action line") = "Y" then
'   response.redirect "../admin/outage_feature_offline.asp"
'end if

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'If Not UserHasPermission( Session("UserId"), "action_line_substatus" ) Then
'	  response.redirect sLevel & "permissiondenied.asp"
'End If 

Dim oCmd, oRst, dDate, iDuration, sTimeZones, sLinks, bShown
dim lcl_new_printer

'-- Add Printer -----------------------------------------------
if request.form("p_action") = "new_printer" then
  'Retrieve the new printer variables
   lcl_new_printer_name    = request("p_new_printer_name")
   lcl_new_layoutid   = request("p_new_layoutid")
   lcl_new_default_printer = request("p_new_default_printer")
   lcl_new_active_flag     = request("p_new_active_flag")

   if lcl_new_layoutid = "" then
      lcl_new_layoutid = 0
   end if

   if lcl_new_default_printer = "" then
      lcl_new_default_printer = 0
   end if

   if lcl_new_active_flag = "" then
      lcl_new_active_flag = 0
   end if

  'Create the record for the new printer
   sSQLi = "INSERT INTO egov_membershipcard_printers (printer_name, layoutid, default_printer, active_flag) "
   sSQLi = sSQLi & " VALUES ("
   sSQLi = sSQLi & "'" & lcl_new_printer_name & "', "
   sSQLi = sSQLi & lcl_new_layoutid           & ", "
   sSQLi = sSQLi & lcl_new_default_printer    & ", "
   sSQLi = sSQLi & lcl_new_active_flag
   sSQLi = sSQLi & ") "
   
   set oCreateSub = Server.CreateObject("ADODB.Recordset")
   oCreateSub.Open sSQLi, Application("DSN"), 3, 1

   response.redirect "edit_org_printers.asp?success=SA"

'-- Delete Printer -----------------------------------------------
elseif request.querystring("CMD") = "delete_printer" then
   lcl_printer_id = request.querystring("PID")

  'First check to see if the PrinterID has been assigned to any Organizations
  'If it does then do not allow the delete
  'Otherwise, dislay the confirm box and delete the PrinterID if selected.
   sSQLe = "SELECT distinct (membershipcard_printer) "
   sSQLe = sSQLe & " FROM organizations "
   sSQLe = sSQLe & " WHERE membershipcard_printer = " & CLng(lcl_printer_id)

   set oExists = Server.CreateObject("ADODB.Recordset")
   oExists.Open sSQLe, Application("DSN"), 3, 1

   if not oExists.eof then
      response.redirect "edit_org_printers.asp?success=NO_DEL"
   else
     'Delete the printer
      sSQLd = "DELETE FROM egov_membershipcard_printers WHERE printerid = " & lcl_printer_id

      set oDeletePrinter = Server.CreateObject("ADODB.Recordset")
      oDeletePrinter.Open sSQLd, Application("DSN"), 3, 1

      response.redirect "edit_org_printers.asp?success=DEL"
   end if

'-- Modify Printer -----------------------------------------------
elseif request.form("p_action") = "modify_printer" then
   for e = 1 to request("total_printers")

       if request("CustomPrinterName_"&e) <> "" then
          lcl_printer_name = request("CustomPrinterName_"&e)
       else
          lcl_printer_name = ""
       end if

       if request("CustomLayoutID_"&e) <> "" then
          lcl_layoutid = request("CustomLayoutID_"&e)
       else
          lcl_layoutid = 0
       end if

       if request("CustomDefaultPrinter_"&e) <> "" then
          lcl_default_printer = request("CustomDefaultPrinter_"&e)
       else
          lcl_default_printer = 0
       end if

       if request("CustomActiveFlag_"&e) <> "" then
          lcl_active_flag = request("CustomActiveFlag_"&e)
       else
          lcl_active_flag = 0
       end if


       if lcl_printer_name    <> "" OR _
          lcl_layoutid        <> "" OR _
          lcl_default_printer <> "" OR _
          lcl_active_flag     <> "" then

          sSQLu = "UPDATE egov_membershipcard_printers SET "
          sSQLu = sSQLu & " printer_name = '"   & lcl_printer_name    & "', "
          sSQLu = sSQLu & " layoutid = "        & lcl_layoutid        & ", "
          sSQLu = sSQLu & " default_printer = " & lcl_default_printer & ", "
          sSQLu = sSQLu & " active_flag = "     & lcl_active_flag
          sSQLu = sSQLu & " WHERE printerid = " & request("CustomPrinterID_"&e)

          set rsu = Server.CreateObject("ADODB.Recordset")
          rsu.Open sSQLu, Application("DSN"),3,1
       end if
   next

   response.redirect("edit_org_printers.asp?success=SU")

end if

'----------------------------------------------------------------
'Setup query to retrieve all of the main statuses that are active
 sSql = "SELECT action_status_id, status_name, orgid, parent_status, display_order, active_flag "
 sSql = sSql & " FROM egov_actionline_requests_statuses "
 sSql = sSql & " WHERE orgid = 0 "
 sSql = sSql & " AND parent_status = 'MAIN' "
 sSql = sSql & " AND active_flag = 'Y' "
 sSql = sSql & " ORDER BY display_order, status_name "
%>
<html>
<head>
  <title>Membership Card Printers</title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript">
<!--
function fnCheckNew() {

  if ((document.new_printer.p_new_printer_name.value != '')) {
       return true;
  }else{
       return false;
  }
}

function confirm_delete(lcl_row_id) {
  lcl_printer_id   = document.getElementById("CustomPrinterID_"+lcl_row_id).value;
  lcl_printer_name = document.getElementById("CustomPrinterName_"+lcl_row_id).value;

  input_box = confirm("Are you sure you want to delete \"" + lcl_printer_name + "\"?");

  if (input_box==true) { 
      // DELETE HAS BEEN VERIFIED
      location.href='edit_org_printers.asp?cmd=delete_printer&pid='+ lcl_printer_id;
  }else{
      // CANCEL DELETE PROCESS
  }
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="document.new_printer.p_new_printer_name.focus()">
<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
  <tr>
      <td>
          <font size="+1"><b>Membership Card Printers</b></font><p>
          <input type="button" value="Close Window" onclick="window.opener.location.reload();parent.close();">
      </td>
      <td width="200">&nbsp;</td>
  </tr>
  <tr>
      <td colspan="2" valign="top">
          <!-- START: NEW PRINTER -->
          <form name="new_printer" method="post" action="edit_org_printers.asp">
            <input type="hidden" name="p_action" value="new_printer">

          <table border="0" cellpadding="2" cellspacing="0" width="100%" class="tableadmin">
            <caption align="left">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                  <tr>
                      <td style="font-size:10px; padding-bottom:5px;">
                          <img src="../images/go.gif" align="absmiddle">&nbsp;
                        		<a href="javascript:if(fnCheckNew()) {document.new_printer.submit();} else {alert('Please enter a Printer Name!');document.new_printer.p_new_printer_name.focus();}">Add Printer</a>
                      </td>
                      <td align="right">
                        <%
                          lcl_message = ""

                          if request("success") = "SU" then
                             lcl_message = "<b style=""color:#FF0000"">*** Successfully Updated... ***</b>"
                          elseif request("success") = "SA" then
                             lcl_message = "<b style=""color:#FF0000"">*** Successfully Added ***</b>"
                          elseif request("success") = "DEL" then
                             lcl_message = "<b style=""color:#FF0000"">*** Successfully Deleted ***</b>"
                          elseif request("success") = "NO_DEL" then
                             lcl_message = "<b style=""color:#FF0000"">*** This Printer has been assigned to atleast one Organization and cannot be deleted. ***</b>"
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
            </caption>
            <tr><th align="left" colspan="2">Add a Printer</th></tr>
            <tr>
                <td>Printer Name:</td>
                <td><input type="text" name="p_new_printer_name" style="width: 350px" maxlength="200"></td>
            </tr>
            <tr>
                <td>[Printer Options]</td>
                <td>
                    Card Layout:&nbsp;<select name="p_new_layoutid">
                                        <option value="1">Front ONLY (Default Layout)</option>
                                        <option value="2">Front and Back</option>
                                      </select>
                    &nbsp;&nbsp;&nbsp;
                    Default Printer:&nbsp;<input type="checkbox" name="p_new_default_printer" value="1">
                    &nbsp;&nbsp;&nbsp;
                    Active:&nbsp;<input type="checkbox" name="p_new_active_flag" value="1">
                </td>
            </tr>
          </table>
          </form>
      		  <p>
      		  <!-- END: NEW PRINTER -->

      		  <!-- START: MODIFY PRINTER -->
        		<form name="modify_printer" action="edit_org_printers.asp" method="post">
        		  <input type="hidden" name="p_action" value="modify_printer">

        		<div style="font-size:10px; padding-bottom:5px;">
         	  <img src="../images/edit.gif" align="absmiddle">&nbsp;
        		  <a href="javascript:document.modify_printer.submit();">Modify</a>
        		</div>

          <table border="0" cellpadding="5" cellspacing="0" class="tableadmin">
        		  <tr valign="bottom">
		              <th nowrap="nowrap" align="left">Printer Name</th>
		           	  <th nowrap="nowrap" align="left">Card Layout</th>
           			  <th nowrap="nowrap">Default<br>Printer</th>
		           	  <th nowrap="nowrap">Active</th>
                <th nowrap="nowrap">Actions</th>
            </tr>
          <%
     			   'Retrieve all of the sub-statuses for the organization for each parent_status
            sSQLp = "SELECT printerid, printer_name, layoutid, default_printer, active_flag "
    					   sSQLp = sSQLp & " FROM egov_membershipcard_printers "
					       sSQLp = sSQLp & " ORDER BY UPPER(printer_name), printerid "

        				set rsp = Server.CreateObject("ADODB.Recordset")
            rsp.Open sSQLp, Application("DSN"), 3, 1

            i = 0
       					if not rsp.eof then
               lcl_bgcolor    = "#ffffff"
					          lcl_line_count = 0

       					  'Loop through all of the printers
       					   while not rsp.eof
                  lcl_line_count = lcl_line_count + 1
			           		  i = i + 1

                  lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")

                  response.write "  <tr bgcolor=""" & lcl_bgcolor & """ align=""center"" valign=""top"">" & vbcrlf

                 'Printer Name
                  response.write "      <td align=""left"" width=""70%"">" & vbcrlf
                  response.write "          <input type=""" & lcl_hidden & """ name=""CustomPrinterID_" & i & """ value=""" & rsp("printerid") & """ size=""5"" maxlength=""10"">" & vbcrlf
                  response.write "          <input type=""text"" name=""CustomPrinterName_" & i & """ value=""" & rsp("printer_name") & """ style=""width: 350px"" maxlength=""200"">" & vbcrlf
                  response.write "      </td>" & vbcrlf

                 'Card Layout
                  response.write "      <td>" & vbcrlf
                  response.write "          <select name=""CustomLayoutID_" & i & """>" & vbcrlf

                  if CLng(rsp("layoutid")) = CLng(2) then
                     lcl_selected_layoutid_1 = ""
                     lcl_selected_layoutid_2 = " selected"
                  else
                     lcl_selected_layoutid_1 = " selected"
                     lcl_selected_layoutid_2 = ""
                  end if

                  response.write "            <option value=""1""" & lcl_selected_layoutid_1 & ">Front ONLY (Default Layout)</option>" & vbcrlf
                  response.write "            <option value=""2""" & lcl_selected_layoutid_2 & ">Front and Back</option>" & vbcrlf
                  response.write "          </select>" & vbcrlf
                  response.write "      </td>" & vbcrlf

                 'Default Printer
                  response.write "      <td>" & vbcrlf
                  response.write "          <select name=""CustomDefaultPrinter_" & i & """>" & vbcrlf

                  if rsp("default_printer") then
                     lcl_selected_defaultprinter_yes = " selected"
                     lcl_selected_defaultprinter_no  = ""
                  else
                     lcl_selected_defaultprinter_yes = ""
                     lcl_selected_defaultprinter_no  = " selected"
                  end if

                  response.write "            <option value=""1""" & lcl_selected_defaultprinter_yes & ">Yes</option>" & vbcrlf
                  response.write "            <option value=""0""" & lcl_selected_defaultprinter_no  & ">No</option>" & vbcrlf
                  response.write "          </select>" & vbcrlf
                  response.write "      </td>" & vbcrlf

                 'Active Flag
                  response.write "      <td>" & vbcrlf
                  response.write "          <select name=""CustomActiveFlag_" & i & """>" & vbcrlf

                  if rsp("active_flag") then
                     lcl_selected_activeflag_yes = " selected"
                     lcl_selected_activeflag_no  = ""
                  else
                     lcl_selected_activeflag_yes = ""
                     lcl_selected_activeflag_no  = " selected"
                  end if

                  response.write "            <option value=""1""" & lcl_selected_activeflag_yes & ">Yes</option>" & vbcrlf
                  response.write "            <option value=""0""" & lcl_selected_activeflag_no  & ">No</option>" & vbcrlf
                  response.write "          </select>" & vbcrlf
                  response.write "      </td>" & vbcrlf
                  response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
                  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
                  response.write "            <tr>" & vbcrlf
                  response.write "                <td bgcolor=""" & lcl_bgcolor & """ nowrap=""nowrap"">" & vbcrlf
                  response.write "                    <a href=""javascript:confirm_delete('" & i & "');""><img src=""../images/small_delete.gif"" border=""0"" align=""absmiddle"">Delete</a>" & vbcrlf
                  response.write "                </td>" & vbcrlf
                  response.write "            </tr>" & vbcrlf
                  response.write "          </table>" & vbcrlf
                  response.write "      </td>" & vbcrlf
                  response.write "  </tr>" & vbcrlf

          						  rsp.movenext
					          wend
            end if
          %>
          </table>
          <input type="<%=lcl_hidden%>" name="total_printers" value="<%=i%>" size="5" maxlength="10">
        		</form>
      </td>
  </tr>
</table>
</body>
</html>
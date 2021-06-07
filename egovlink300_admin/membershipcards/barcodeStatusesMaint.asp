<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="membership_card_functions.asp" -->
<%
  sLevel = "../" ' Override of value from common.asp

  if not userhaspermission(session("userid"),"create membership cards") then
     response.redirect sLevel & "permissiondenied.asp"
  end if

  lcl_member_id  = request("memberid")
  lcl_action     = request("action")
  lcl_poolpassid = request("poolpassid")

 'Set up the return Session variables for the page
  session("RedirectLang") = "Return to Member ID"
  session("RedirectPage") = "image_display.asp?memberid=" & lcl_member_id & "&action=" & lcl_action & "&poolpassid=" & lcl_poolpassid

 'Check for org features
  'lcl_orghasfeature_pool_attendance_view             = orghasfeature("pool_attendance_view")
  'lcl_orghasfeature_customreports_membership_scanlog = orghasfeature("customreports_membership_scanlog")

 'Check for user features
  'lcl_userhaspermission_customreports_membership_scanlog = userhaspermission(session("userid"),"customreports_membership_scanlog")

  lcl_success = ""

  if request("success") <> "" then
     lcl_success = request("success")
     lcl_success = ucase(lcl_success)
  end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Maintain Barcode Statuses}</title>
  
  <link rel="stylesheet" type="text/css" href="../global.css">
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

<style type="text/css">
#pageHeader {
   margin-bottom: 10px;
   font-size: 1.25em;
   font-weight: bold;
}

#buttonRow {
   margin-bottom: 10px;
}

#barcodeStatusTable {
   width: 600px !important;
   border: 1pt solid #c0c0c0;
     border-radius: 6px;
}

#barcodeStatusTable thead th {
   background-color: #c0c0c0;
   border-bottom: 1pt solid #808080;
   padding: 6px;
}

#barcodeStatusTable tbody td {
   border-bottom: 1pt solid #c0c0c0;
   padding: 6px;
}

#barcodeStatusTable .statusName {
   width: 100%;
   height: 18px;
}

input.button {
    border: 1px solid #777777;
    border-radius: 4px;
    color: #000000;
    cursor: pointer;
    padding: 4px;
}

input.button:hover {
    background-color: #666666;
    color: #ffffff;
}

input {
    font-family: Verdana,Tahoma,Arial;
    font-size: 11px;
}

#screenMsg {
   text-align:  right;
   color:       #ff0000;
   font-size:   10pt;
   font-weight: bold;
}
</style>

  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<script type="text/javascript">
   $(document).ready(function() {
      if('<%=lcl_success%>' == 'SU') {
         displayScreenMsg('Successfully Updated');
      }
      
      $('#backButton').click(function() {
         var lcl_url  = 'image_display.asp';
             lcl_url += '?memberid=<%=lcl_member_id%>';
             lcl_url += '&action=<%=lcl_action%>';
             lcl_url += '&poolpassid=<%=lcl_poolpassid%>';
         
         location.href = lcl_url;
      });

      $('#saveButton').click(function() {
         $('#barcodeStatuses').submit();
      });

      $('#addButton').click(function() {
        var lcl_total_statuses = $('#totalStatuses').val();
        var lcl_bgcolor        = $('#barcodeRow' + lcl_total_statuses).prop('bgColor');

        if(lcl_total_statuses == '') {
           lcl_total_statuses = 0;
        }

        if(lcl_bgcolor == "#eeeeee") {
           lcl_bgcolor = "#ffffff";
        } else {
           lcl_bgcolor = "#eeeeee";
        }

        //Determine what the rowid and total row count is
        var num              = new Number(lcl_total_statuses);
        var lcl_new_total    = (num + 1);
        var lcl_new_rowid    = lcl_new_total.toString();
        var lcl_row_html     = '';

        lcl_row_html  = '<tr id="barcodeRow' + lcl_new_rowid + '" class="barcodeRow" align="center" valign="top" bgcolor="' + lcl_bgcolor + '">';
        lcl_row_html += '    <td align="left">';
        lcl_row_html += '         <input type="hidden" name="statusid' + lcl_new_rowid + '" id="statusid' + lcl_new_rowid + '" value="0" />';
        lcl_row_html += '         <input type="text" name="statusname' + lcl_new_rowid + '" id="statusname' + lcl_new_rowid + '" value="" class="statusName" />';
        lcl_row_html += '     </td>';
        lcl_row_html += '     <td>';
        lcl_row_html += '         <input type="radio" name="isActive" id="isActive' + lcl_new_rowid + '" value="" />';
        lcl_row_html += '     </td>';
        lcl_row_html += '     <td>';
        lcl_row_html += '         <input type="checkbox" name="isEnabled' + lcl_new_rowid + '" id="isEnabled' + lcl_new_rowid + '" value="Y" checked="checked" />';
        lcl_row_html += '     </td>';
        lcl_row_html += '     <td>';
        lcl_row_html += '         <input type="checkbox" name="removeStatus' + lcl_new_rowid + '" id="removeStatus' + lcl_new_rowid + '" value="Y" onclick="modifyFields(\'' + lcl_new_rowid + '\');" />';
        lcl_row_html += '     </td>';
        lcl_row_html += '</tr>';
        
        //Append the new row to the table and increment the barcodes total.
        $('#barcodeStatusTable').append(lcl_row_html);
        $('#totalStatuses').val(lcl_new_rowid);
      });
   });

function modifyFields(iRowID) {
  var lcl_removeStatus  = $('#removeStatus' + iRowID);
  var lcl_fieldsDisabled = false;
  var lcl_fieldsBGColor  = '';

  if(lcl_removeStatus.prop('checked')) {
     lcl_fieldsDisabled = true;
     lcl_fieldsBGColor  = '#eeeeee';
  }

  $('#statusname' + iRowID).css('background-color', lcl_fieldsBGColor);

  $('#statusname'     + iRowID).prop('disabled', lcl_fieldsDisabled);
  $('#isActiveStatus' + iRowID).prop('disabled', lcl_fieldsDisabled);
  $('#isEnabled'      + iRowID).prop('disabled', lcl_fieldsDisabled);

}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     $('#screenMsg').html('*** ' + iMsg + ' ***&nbsp;&nbsp;&nbsp;');
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  $('#screenMsg').html('&nbsp;');
}
</script>

</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""barcodeStatuses"" id=""barcodeStatuses"" method=""post"" action=""barcodeStatusesAction.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""memberid"" id=""memberid"" value="""     & lcl_member_id    & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""action"" id=""action"" value="""         & lcl_action       & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""poolpassid"" id=""poolpassid"" value=""" & lcl_poolpassid   & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value="""           & session("orgid") & """ />" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "    <div id=""pageHeader"">Maintain Barcode Statuses</div>" & vbcrlf
  response.write "    <div id=""buttonRow"">" & vbcrlf
  response.write "      <input type=""button"" name=""backButton"" id=""backButton"" value=""Back to Membership Card"" class=""button"" />" & vbcrlf
  response.write "      <input type=""button"" name=""addButton"" id=""addButton"" value=""Add Status"" class=""button"" />" & vbcrlf
  response.write "      <input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "    </div>" & vbcrlf
                      displayBarcodeStatues session("orgid")
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayBarcodeStatues(iOrgID)

  sLineCount = 0
  sBGColor   = "#eeeeee"

  response.write "<table cellspacing=""0"" cellpadding=""0"" id=""barcodeStatusTable"">" & vbcrlf
  response.write "  <caption id=""screenMsg"">&nbsp;</caption>" & vbcrlf
  response.write "  <thead>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <th align=""left"">Status Name</th>" & vbcrlf
  response.write "      <th>Is ""Active""<br />Status</th>" & vbcrlf
  response.write "      <th>Enabled</th>" & vbcrlf
  response.write "      <th>Delete</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  </thead>" & vbcrlf
  response.write "  <tbody>" & vbcrlf

  sSQL = "SELECT "
  sSQL = sSQL & "statusid, "
  sSQL = sSQL & "statusname, "
  sSQL = sSQL & "isActiveStatus, "
  sSQL = sSQL & "isEnabled "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " ORDER BY statusname "

  set oBarcodeStatus = Server.CreateObject("ADODB.Recordset")
  oBarcodeStatus.Open sSQL, Application("DSN"), 3, 1
	
  if not oBarcodeStatus.eof then
     do while not oBarcodeStatus.eof
        sLineCount               = sLineCount + 1
        sStatusID                = oBarcodeStatus("statusid")
        sStatusName              = oBarcodeStatus("statusname")
        sIsEnabledChecked        = ""
        sIsActiveStatusChecked   = ""
        sDisplayRemoveCheckbox   = "&nbsp;"
        sStatusAssignedToBarcode = checkStatusAssignedToBarcode(iOrgID, _
                                                                sStatusID)

        if oBarcodeStatus("isActiveStatus") then
           sIsActiveStatusChecked = " checked=""checked"""
        end if

        if oBarcodeStatus("isEnabled") then
           sIsEnabledChecked = " checked=""checked"""
        end if

        if not sStatusAssignedToBarcode then
           sDisplayRemoveCheckbox = "<input type=""checkbox"" name=""removeStatus" & sLineCount & """ id=""removeStatus" & sLineCount & """ value=""Y"" onclick=""modifyFields('" & sLineCount & "');"" />" & vbcrlf
        end if

        response.write "  <tr id=""barcodeRow" & sLineCount & """ class=""barcodeRow"" align=""center"" valign=""top"" bgcolor=""" & sBGColor & """>" & vbcrlf
        response.write "      <td align=""left"">" & vbcrlf
        response.write "          <input type=""hidden"" name=""statusid" & sLineCount & """ id=""statusid" & sLineCount & """ value=""" & sStatusID & """ />" & vbcrlf
        response.write "          <input type=""text"" name=""statusname" & sLineCount & """ id=""statusname" & sLineCount & """ value=""" & sStatusName & """ class=""statusName"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""radio"" name=""isActiveStatus"" id=""isActiveStatus" & sLineCount & """ value=""" & sStatusID & """" & sIsActiveStatusChecked & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""isEnabled" & sLineCount & """ id=""isEnabled" & sLineCount & """ value=""Y""" & sIsEnabledChecked & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & sDisplayRemoveCheckbox & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        sBGColor = changeBGColor(sBGColor, "#eeeeee", "#ffffff")

        oBarcodeStatus.movenext
     loop
  end if

  oBarcodeStatus.close
  set oBarcodeStatus = Nothing

  response.write "  </tbody>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<input type=""hidden"" name=""totalStatuses"" id=""totalStatuses"" value=""" & sLineCount & """ />" & vbcrlf

end sub

'------------------------------------------------------------------------------
function checkStatusAssignedToBarcode(iOrgID, _
                                      iStatusID)

  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT count(memberbarcodeid) as totalBarcodes "
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes "
  sSQL = sSQL & " WHERE barcode_statusid = " & iStatusID
  sSQL = sSQL & " AND orgid = " & iOrgID

  set oStatusAssigned = Server.CreateObject("ADODB.Recordset")
  oStatusAssigned.Open sSQL, Application("DSN"), 3, 1
	
  if not oStatusAssigned.eof then
     if oStatusAssigned("totalBarcodes") > 0 then
        lcl_return = true
     end if
  end if

  oStatusAssigned.close
  set oStatusAssigned = nothing

  checkStatusAssignedToBarcode = lcl_return

end function
%>
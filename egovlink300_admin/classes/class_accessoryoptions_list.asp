<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_accessoryoptions_list.asp
' AUTHOR: David Boyer
' CREATED: 11/18/09
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Displays all accessory options (t-shirt, pants, etc) available for selection within Team Registration section
'               within a class/event sign up.
'
' MODIFICATION HISTORY
' 1.0  11/18/09	David Boyer	- Initial version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

 Dim iUserID, bIsRootAdmin, sShowDetails

 sLevel = "../"  'Override of value from common.asp

 'if not userhaspermission(session("userid"),"user permission") then
	'   response.redirect sLevel & "permissiondenied.asp"
 'end if

 'bIsRootAdmin = UserIsRootAdmin(session("userid"))

'Retrieve the paramaters
 lcl_classid       = request("classid")
 lcl_accessorytype = ""
 lcl_onload        = ""
 lcl_success       = request("success")

 if request("atype") <> "" then
    lcl_accessorytype = ucase(request("atype"))
 end if

'Check for an "edit display" for the T-shirt label
 if orgHasDisplay(session("orgid"),"class_teamregistration_tshirt_label") then
    lcl_label_tshirt = getOrgDisplay(session("orgid"),"class_teamregistration_tshirt_label")
 else
    lcl_label_tshirt = "T-Shirt"
 end if

'Setup up the display label for the accessory type
 if lcl_accessorytype = "PANTS" then
    lcl_accessorytype_label = "Pants"
 else
    lcl_accessorytype_label = lcl_label_tshirt
 end if

 'if request.ServerVariables("REQUEST_METHOD") = "POST" then
 '  'Remove all of the assignments for the org and class/event
 '   sSQL = "DELETE FROM egov_class_teamroster_accessories_to_class "
 '   sSQL = sSQL & " WHERE orgid = " & session("orgid")
 '   sSQL = sSQL & " AND classid = " & lcl_classid
 '   sSQL = sSQL & " AND accessoryid IN (select accessoryid "
 '   sSQL = sSQL &                     " from egov_class_teamroster_accessories "
 '   sSQL = sSQL &                     " where UPPER(accessorytype) = '" & UCASE(lcl_accessorytype) & "') "

 '  	set oDelAccessories = Server.CreateObject("ADODB.Recordset")
 '  	oDelAccessories.Open sSQL, Application("DSN"), 3, 1

 '   set oDelAccessories = nothing

   'Assign all values to org for the class/event
 '   lcl_total_options = 0

 '   if request("total_options") <> "" then
 '      lcl_total_options = request("total_options")
 '   end if

   'Insert all of the options that have been "checked"
 '   if lcl_total_options > 0 then
 '      for i = 1 to lcl_total_options
 '         if request.form("accessoryid_" & i) <> "" then
 '            sSQL = "INSERT INTO egov_class_teamroster_accessories_to_class ("
 '            sSQL = sSQL & "orgid, "
 '            sSQL = sSQL & "classid, "
 '            sSQL = sSQL & "accessoryid, "
 '            sSQL = sSQL & "displayorder"
 '            sSQL = sSQL & ") VALUES ("
 '            sSQL = sSQL & session("orgid")                 & ", "
 '            sSQL = sSQL & lcl_classid                      & ", "
 '            sSQL = sSQL & request.form("accessoryid_" & i) & ", "
 '            sSQL = sSQL & i
 '            sSQL = sSQL & ") "

 '           	set oInsertAccessories = Server.CreateObject("ADODB.Recordset")
 '           	oInsertAccessories.Open sSQL, Application("DSN"), 3, 1

 '            set oInsertAccessories = nothing
 '         end if
 '      next

 '      if i > 0 then
 '         lcl_success = "SU"
 '      end if
 '   end if

 'end if

'Check for a screen message
 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'setupAccessoryAssignments(session("orgid"))
%>
<html>
<head>
	 <title>E-GovLink Administration Console {Maintain Accessory Options - <%=lcl_accessorytype_label%>}</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="security.css" />

  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

<style>
  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
  }
</style>

<script language="javascript">
<!--

$(document).ready(function(){
});

checked=false;
function checkedAll () {
	 var x = document.getElementById('accessoryOptionsSave');
	 if (checked == false) {
      checked = true
  }else{
      checked = false
  }

 	for (var i=0; i < x.elements.length; i++) {
     	 x.elements[i].checked = checked;
 	}
}

function saveChanges() {
  var lcl_totaloptions = $('#total_options').val();
  var lcl_false_count  = 0;

  if(lcl_totaloptions > 0) {
     for (var i = lcl_totaloptions; i >= 1; i--) {
        if(! document.getElementById('delete_' + i).checked) {
           if($('#displayorder_' + i).val() == '') {
              inlineMsg(document.getElementById('displayorder_' + i).id,'<strong>Required Field Missing: </strong>Display Order',10,'displayorder_' + i);
              $('#displayorder_' + i).focus();
              lcl_false_count = lcl_false_count + 1;
           } else {
              if(! Number($('#displayorder_' + i).val())) {
                 if(Number($('#displayorder_' + i).val()) != '0') {
                    $('#displayorder_' + i).focus();
                    inlineMsg(document.getElementById('displayorder_' + i).id,'<strong>Invalid Value: </strong> "Display Order" must be numeric.',10,'displayorder_' + i);
                    lcl_false_count = lcl_false_count + 1;
                 }
              }
           }

           if($('#accessoryvalue_' + i).val() == '') {
              inlineMsg(document.getElementById('accessoryvalue_' + i).id,'<strong>Required Field Missing: </strong>Value',10,'accessoryvalue_' + i);
              $('#accessoryvalue_' + i).focus();
              lcl_false_count = lcl_false_count + 1;
           }

           if($('#accessoryname_' + i).val() == '') {
              inlineMsg(document.getElementById('accessoryname_' + i).id,'<strong>Required Field Missing: </strong>Name',10,'accessoryname_' + i);
              $('#accessoryname_' + i).focus();
              lcl_false_count = lcl_false_count + 1;
           }
        }
     }
  }

  if(lcl_false_count > 0) {
     return false;
  } else {
     $('#accessoryOptionsSave').submit();
  }

}

function addAccessory() {
  var lcl_total   = $('#total_options').val();
  var lcl_bgcolor = $('#accessoryRow' + lcl_total).prop('bgcolor');

  //Determine what the rowid and total row count is
  var num           = new Number(lcl_total);
  var lcl_new_total = (num + 1);
  var lcl_new_rowid = lcl_new_total.toString();
  var lcl_row_html  = "";

  if(lcl_bgcolor == "#eeeeee") {
     lcl_bgcolor = "#ffffff";
  } else {
     lcl_bgcolor = "#eeeeee";
  }

  //Build the new row
  lcl_row_html += '  <tr id="accessoryRow' + lcl_new_rowid + '" bgcolor="' + lcl_bgcolor + '">';
  lcl_row_html += '      <td align="center">';
  lcl_row_html += '          <input type="checkbox" name="assign_accessoryid_' + lcl_new_rowid + '" id="assign_accessoryid_' + lcl_new_rowid + '" value="' + lcl_new_rowid + '" size="3" maxlength="100" />';
  lcl_row_html += '          <input type="hidden" name="accessoryid_' + lcl_new_rowid + '" id="accessoryid_' + lcl_new_rowid + '" value="0" size="30" maxlength="10" />';
  lcl_row_html += '      </td>';
  lcl_row_html += '      <td>';
  lcl_row_html += '          <input type="text" name="accessoryname_' + lcl_new_rowid + '" id="accessoryname_' + lcl_new_rowid + '" value="" size="30" maxlength="50" onchange="clearMsg(\'accessoryname_' + lcl_new_rowid + '\');" />';
  lcl_row_html += '      </td>';
  lcl_row_html += '      <td>';
  lcl_row_html += '          <input type="text" name="accessoryvalue_' + lcl_new_rowid + '" id="accessoryvalue_' + lcl_new_rowid + '" value="" size="30" maxlength="50" onchange="clearMsg(\'accessoryvalue_' + lcl_new_rowid + '\');" />';
  lcl_row_html += '      </td>';
  lcl_row_html += '      <td align="center">';
  lcl_row_html += '          <input type="text" name="displayorder_' + lcl_new_rowid + '" id="displayorder_' + lcl_new_rowid + '" value="' + lcl_new_rowid + '" size="4" maxlength="10" onchange="clearMsg(\'displayorder_' + lcl_new_rowid + '\');" />';
  lcl_row_html += '      </td>';
  lcl_row_html += '      <td align="center">';
  lcl_row_html += '          <input type="checkbox" name="delete_' + lcl_new_rowid + '" id="delete_' + lcl_new_rowid + '" value="0"  onclick="deleteAccessory(' + lcl_new_rowid + ');" />';
  lcl_row_html += '      </td>';
  lcl_row_html += '  </tr>';

  //Append the new row to the table and increment the sub-catgories total.
  $('#accessories_table').append(lcl_row_html);
  $('#total_options').val(lcl_new_rowid);
  $('#accessoryname_' + lcl_new_rowid).focus();
}

function deleteAccessory(p_linenum) {
  var lcl_isDisabled = false;

  if(p_linenum != '') {
     if(document.getElementById('delete_' + p_linenum).checked) {
        lcl_isDisabled = true;
     }
  }

  $('#accessoryname_'  + p_linenum).prop('disabled',lcl_isDisabled);
  $('#accessoryvalue_' + p_linenum).prop('disabled',lcl_isDisabled);
  $('#displayorder_'   + p_linenum).prop('disabled',lcl_isDisabled);
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
<body onload="<%=lcl_onload%>">
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""1000px"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Maintain Class/Events - Accessory Options: " & lcl_accessorytype_label & "</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  displayButtons
  displayAccessoryOptions session("orgid"), lcl_classid, lcl_accessorytype
  displayButtons

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayAccessoryOptions(p_orgid, p_classid, p_accessorytype)

  dim sOrgID, sClassID, sAccessoryType

  sOrgID         = 0
  sClassID       = 0
  sAccessoryType = ""

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_classid <> "" then
     sClassID = clng(p_classid)
  end if

  if p_accessorytype <> "" then
     sAccessoryType = ucase(p_accessorytype)
  end if

  response.write "<form name=""accessoryOptionsSave"" id=""accessoryOptionsSave"" method=""post"" action=""class_accessoryoptions_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value="""     & sOrgID         & """ />" & vbcrlf
  response.write "	 <input type=""hidden"" name=""classid"" id=""classid"" value=""" & sClassID       & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""atype"" id=""atype"" value="""     & sAccessoryType & """ />" & vbcrlf
  response.write "<div class=""shadow"">" & vbcrlf
  response.write "<table id=""accessories_table"" cellspacing=""0"" cellpadding=""2"" border=""0"" class=""tableadmin"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  'response.write "      <th><input type=""checkbox"" name=""selectAll"" id=""selectAll"" value=""on"" onclick=""checkedAll();"" /></th>" & vbcrlf
  response.write "      <th>Active</th>" & vbcrlf
  response.write "      <th align=""left"">Name (value displayed to citizen)</th>" & vbcrlf
  response.write "      <th align=""left"">Value (value displayed on receipt)</th>" & vbcrlf
  response.write "      <th>Display<br />Order</th>" & vbcrlf
  response.write "      <th>Delete</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Get all of the "default" and "org specific" accessory options
  if sAccessoryType <> "" then
     sAccessoryType = dbsafe(sAccessoryType)
     sAccessoryType = "'" & sAccessoryType & "'"
  end if

 	sSQL = "SELECT "
  sSQL = sSQL & " accessoryid, "
  sSQL = sSQL & " accessorytype, "
  sSQL = sSQL & " accessoryname, "
  sSQL = sSQL & " accessoryvalue, "
  sSQL = sSQL & " displayorder "
  sSQL = sSQL & " FROM egov_class_teamroster_accessories "
  sSQL = sSQL & " WHERE UPPER(accessorytype) = " & sAccessoryType
  sSQL = sSQL & " AND orgid = " & sOrgID
  sSQL = sSQL & " ORDER BY displayorder, accessoryname "

 	set oAccessories = Server.CreateObject("ADODB.Recordset")
 	oAccessories.Open sSQL, Application("DSN"), 3, 1

  lcl_bgcolor = "#eeeeee"
  sLineCount   = 0

 	if not oAccessories.eof then
   		do while not oAccessories.eof
        sLineCount                      = sLineCount + 1
        lcl_bgcolor                     = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_checked                     = ""
        lcl_show_delete                 = true
        lcl_displayvalue_accessoryname  = ""
        lcl_displayvalue_accessoryvalue = ""
        lcl_fieldtype_accessoryname     = "text"
        lcl_fieldtype_accessoryvalue    = "text"

        if checkAssignedAccessory(sOrgID, sClassID, oAccessories("accessoryid")) then
           lcl_checked = " checked=""checked"""
        end if

        if checkAssignedtoOtherClasses(sOrgID, sClassID, oAccessories("accessoryid")) then
           lcl_show_delete = false
        end if

        if not lcl_show_delete then
           lcl_displayvalue_accessoryname  = oAccessories("accessoryname")
           lcl_displayvalue_accessoryvalue = oAccessories("accessoryvalue")
           lcl_fieldtype_accessoryname     = "hidden"
           lcl_fieldtype_accessoryvalue    = "hidden"
        end if

     			response.write "  <tr id=""accessoryRow_" & sLineCount & """ bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
    				response.write "      <td align=""center"">" & vbcrlf
        response.write "          <input type=""checkbox"" name=""assign_accessoryid_" & sLineCount & """ id=""assign_accessoryid_" & sLineCount & """ value=""" & oAccessories("accessoryid") & """" & lcl_checked & " />" & vbcrlf
        response.write "          <input type=""hidden"" name=""accessoryid_" & sLineCount & """ id=""accessoryid_" & sLineCount & """ value=""" & oAccessories("accessoryid") & """" & lcl_checked & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
		     	response.write "      <td>" & vbcrlf
        response.write "          " & lcl_displayvalue_accessoryname  & "<input type=""" & lcl_fieldtype_accessoryname & """ name=""accessoryname_" & sLineCount & """ id=""accessoryname_" & sLineCount & """ value=""" & oAccessories("accessoryname") & """ size=""30"" maxlength=""50"" onchange=""clearMsg('accessoryname_" & sLineCount & "');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
		     	response.write "      <td>" & vbcrlf
        response.write "          " & lcl_displayvalue_accessoryvalue & "<input type=""" & lcl_fieldtype_accessoryname & """ name=""accessoryvalue_" & sLineCount & """ id=""accessoryvalue_" & sLineCount & """ value=""" & oAccessories("accessoryvalue") & """ size=""30"" maxlength=""50"" onchange=""clearMsg('accessoryvalue_" & sLineCount & "');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
		     	response.write "      <td align=""center"">" & vbcrlf
        response.write "          <input type=""text"" name=""displayorder_" & sLineCount & """ id=""displayorder_" & sLineCount & """ value=""" & sLineCount & """ size=""4"" maxlength=""10"" onchange=""clearMsg('displayorder_" & sLineCount & "');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
    				response.write "      <td align=""center"">" & vbcrlf

        if lcl_show_delete then
           response.write "          <input type=""checkbox"" name=""delete_" & sLineCount & """ id=""delete_" & sLineCount & """ value=""" & oAccessories("accessoryid") & """ onclick=""deleteAccessory('" & sLineCount & "');"" />" & vbcrlf
        else
           response.write "          <input type=""hidden"" name=""delete_" & sLineCount & """ id=""delete_" & sLineCount & """ value="""" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
		     	response.write "  </tr>" & vbcrlf

   			  oAccessories.movenext
   		loop
  else
     response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
		   response.write "      <td colspan=""3"">No Accessory Options available</td>" & vbcrlf
		   response.write "  </tr>" & vbcrlf
  end if

 	oAccessories.close
	 set oAccessories = nothing 

		response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "<div align=""right"">Total Options: [" & sLineCount & "]</div>" & vbcrlf
  response.write "<input type=""hidden"" name=""total_options"" id=""total_options"" value=""" & sLineCount & """ />" & vbcrlf
  response.write "</form>" & vbcrlf
end sub

'------------------------------------------------------------------------------
function checkAssignedAccessory(iOrgID, iClassID, iAccessoryID)

  lcl_return = false

  sSQL = "SELECT classaccessoryid "
  sSQL = sSQL & " FROM egov_class_teamroster_accessories_to_class "
  sSQL = sSQL & " WHERE orgid = "     & iOrgID
  sSQL = sSQL & " AND classid = "     & iClassID
  sSQL = sSQL & " AND accessoryid = " & iAccessoryID

 	set oAssigned = Server.CreateObject("ADODB.Recordset")
 	oAssigned.Open sSQL, Application("DSN"), 3, 1

  if not oAssigned.eof then
     lcl_return = true
  end if

  oAssigned.close
  set oAssigned = nothing

  checkAssignedAccessory = lcl_return

end function

'------------------------------------------------------------------------------
function checkAssignedtoOtherClasses(iOrgID, iClassID, iAccessoryID)

  lcl_return = false

  sSQL = "SELECT count(classaccessoryid) as total_assignments "
  sSQL = sSQL & " FROM egov_class_teamroster_accessories_to_class "
  sSQL = sSQL & " WHERE orgid = "     & iOrgID
  sSQL = sSQL & " AND accessoryid = " & iAccessoryID
  sSQL = sSQL & " AND classid <> "    & iClassID

 	set oAccessoryAssigned = Server.CreateObject("ADODB.Recordset")
 	oAccessoryAssigned.Open sSQL, Application("DSN"), 3, 1

  if not oAccessoryAssigned.eof then
     if oAccessoryAssigned("total_assignments") > 0 then
        lcl_return = true
     end if
  end if

  oAccessoryAssigned.close
  set oAccessoryAssigned = nothing

  checkAssignedtoOtherClasses = lcl_return

end function

'------------------------------------------------------------------------------
sub displayButtons()

  response.write "<p>" & vbcrlf
  response.write "<input type=""button"" name=""closeButton"" id=""closeButton"" value=""Close Window"" class=""button"" onclick=""parent.close();"" />" & vbcrlf
  response.write "<input type=""button"" name=""addAccessoryButton"" id=""addAccessoryButton"" value=""Add Accessory"" class=""button"" onclick=""addAccessory()"" />" & vbcrlf
  response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""saveChanges();"" />" & vbcrlf
  response.write "</p>" & vbcrlf

end sub

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
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub setupAccessoryAssignments(iOrgID)

  sSQL = "SELECT atc.classaccessoryid, "
  sSQL = sSQL & " atc.accessoryid, "
  sSQL = sSQL & " ca.accessoryname, "
  sSQL = sSQL & " ca.accessorytype "
  sSQL = sSQL & " FROM egov_class_teamroster_accessories_to_class atc "
  sSQL = sSQL &      " INNER JOIN egov_class_teamroster_accessories ca ON atc.accessoryid = ca.accessoryid "
  sSQL = sSQL & " WHERE atc.orgid = " & iOrgID
  sSQL = sSQL & " ORDER BY accessoryid, classaccessoryid "
'dtb_debug(sSQL)
 	set oSetupAccessoryAssignments = Server.CreateObject("ADODB.Recordset")
 	oSetupAccessoryAssignments.Open sSQL, Application("DSN"), 3, 1

  if not oSetupAccessoryAssignments.eof then
     do while not oSetupAccessoryAssignments.eof

        lcl_new_accessoryid = getAccessoryIDByAccessoryName(iOrgID, oSetupAccessoryAssignments("accessorytype"), oSetupAccessoryAssignments("accessoryname"))
'dtb_debug("[" & lcl_new_accessoryid & "]")
        sSQLu = "UPDATE egov_class_teamroster_accessories_to_class SET "
        sSQLu = sSQLu & " accessoryid_new = " & lcl_new_accessoryid
        sSQLu = sSQLu & " WHERE classaccessoryid = " & oSetupAccessoryAssignments("classaccessoryid")

       	set oSetupAccessoryAssignmentsUpdate = Server.CreateObject("ADODB.Recordset")
       	oSetupAccessoryAssignmentsUpdate.Open sSQLu, Application("DSN"), 3, 1

        'set oSetupAccessoryAssignmentsUpdate = nothing


        oSetupAccessoryAssignments.movenext
     loop
  end if

  oSetupAccessoryAssignments.close
  set oSetupAccessoryAssignments = nothing


end sub

'------------------------------------------------------------------------------
function getAccessoryIDByAccessoryName(iOrgID, iAccessoryType, iAccessoryName)

  lcl_return = 0

  sSQLn = "SELECT accessoryid "
  sSQLn = sSQLn & " FROM egov_class_teamroster_accessories "
  sSQLn = sSQLn & " WHERE orgid = " & iOrgID
  sSQLn = sSQLn & " AND UPPER(accessorytype) = '" & dbsafe(ucase(iAccessoryType)) & "'"
  sSQLn = sSQLn & " AND UPPER(accessoryname) = '" & dbsafe(ucase(iAccessoryName)) & "'"

 	set oGetAccessoryID = Server.CreateObject("ADODB.Recordset")
 	oGetAccessoryID.Open sSQLn, Application("DSN"), 3, 1

  if not oGetAccessoryID.eof then
     lcl_return = oGetAccessoryID("accessoryid")
  end if

  oGetAccessoryID.close
  set oGetAccessoryID = nothing

  getAccessoryIDByAccessoryName = lcl_return

end function

%>
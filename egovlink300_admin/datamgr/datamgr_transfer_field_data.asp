<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_transfer_field_data.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the fields within each section and allows the user to
'               transfer data from one field to another.
'
' MODIFICATION HISTORY
' 1.0 07/14/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel          = "../"  'Override of value from common.asp
 lcl_isRootAdmin = False

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Determine if the parent feature is "offline"
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user has access to maintain
 lcl_feature     = "datamgr"
 lcl_featurename = getFeatureName(lcl_feature)
 lcl_pagetitle   = lcl_featurename & ": Transfer Field Data"
 lcl_success     = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

'Check for search options
 lcl_sc_dm_typeid = ""

 if request("sc_dm_typeid") <> "" then
    lcl_sc_dm_typeid = request("sc_dm_typeid")
    lcl_sc_dm_typeid = clng(lcl_sc_dm_typeid)
 end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

<script language="javascript">
<!--
function validateFields() {
  document.getElementById("transferData").submit();
}

function setupTransferData(iRowID) {
  var lcl_totalfields         = document.getElementById('totalFields').value;
  var lcl_dm_fieldid          = '';
  var lcl_transfer_field_data = '';
  var lcl_return_false        = 0;
  var lcl_indexCount          = 0;
  var lcl_dropdown_disabled   = '';

  lcl_dm_fieldid          = document.getElementById('dm_fieldid' + iRowID).value;
  lcl_transfer_field_data = document.getElementById('transfer_field_data' + iRowID).value;

  if(lcl_dm_fieldid == lcl_transfer_field_data) {
     document.getElementById('transfer_field_data' + iRowID).focus();
     inlineMsg(document.getElementById('transfer_field_data' + iRowID).id,'<strong>Invalid Value: </strong> Cannot transfer data to the same field.',10,'transfer_field_data' + iRowID);
     lcl_return_false = lcl_return_false + 1;
  } else {
     clearMsg('transfer_field_data' + iRowID);

     //Loop through and disable all fields except the current one.
     $('.transferFieldData').each(function(index) {
       lcl_indexCount = index + 1;

       if(lcl_transfer_field_data != '') {
          if(lcl_indexCount != iRowID) {
             lcl_dropdown_disabled = 'disabled';
          } else {
             lcl_dropdown_disabled = '';
          }
       }

       document.getElementById('transfer_field_data' + lcl_indexCount).disabled = lcl_dropdown_disabled;

     });
  }

  if (lcl_return_false > 0) {
      return false;
  }else{
      return true;
  }
}

function confirmDelete(p_id) {
  lcl_datamgr = document.getElementById("datamgr"+p_id).innerHTML;

 	if (confirm("Are you sure you want to delete '" + lcl_datamgr + "' ?")) { 
  				//DELETE HAS BEEN VERIFIED
		  		location.href='datamgr_action.asp<%=lcl_delete_datamgr%>&dmid='+ p_id;
		}
}

function doCalendar(ToFrom) {
  w = 350;
  h = 250;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}

function openCustomReports(p_report) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("../customreports/customreports.asp?cr=' + p_report + '&dmt=<%=lcl_dm_typeid%>", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,resizable=1,menubar=0,left=' + l + ',top=' + t + '")');
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
<%
 response.write "<div id=""content"">" & vbcrlf
 response.write " 	<div id=""centercontent"">" & vbcrlf
 response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"">" & vbcrlf
' response.write "          <div style=""margin-top:20px; margin-left:20px;"">" & vbcrlf
 response.write "          <p>" & vbcrlf
 response.write "          <form name=""transferDataSeach"" id=""transferDataSearch"" method=""post"" action=""datamgr_transfer_field_data.asp"">" & vbcrlf
 response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""1000px"">" & vbcrlf
 response.write "            <tr>" & vbcrlf
 response.write "                <td><font size=""+1""><strong>" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
 response.write "                <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;"">&nbsp;</span></td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
 response.write "          </table>" & vbcrlf
 response.write "          </p>" & vbcrlf

 response.write "          <p>" & vbcrlf
 response.write "            <fieldset name=""searchoptions"" id=""searchoptions"" class=""fieldset"">" & vbcrlf
 response.write "              <legend>Search Options</legend>" & vbcrlf
 response.write "              <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
 response.write "                <tr>" & vbcrlf
 response.write "                    <td>" & vbcrlf
 response.write "                        DM Type: " & vbcrlf
 response.write "                        <select name=""sc_dm_typeid"" id=""sc_dm_typeid"">" & vbcrlf
 response.write "                          <option value="""">&nbsp;</option>" & vbcrlf
                                           displayDMTypeOptions session("orgid"), lcl_sc_dm_typeid
 response.write "                        </select>" & vbcrlf
 response.write "                    </td>" & vbcrlf
 response.write "                </tr>" & vbcrlf
 response.write "                <tr>" & vbcrlf
 response.write "                    <td>" & vbcrlf
 response.write "                        <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
 response.write "                    </td>" & vbcrlf
 response.write "                </tr>" & vbcrlf
 response.write "              </table>" & vbcrlf
 response.write "              </form>" & vbcrlf
 response.write "            </fieldset>" & vbcrlf
 response.write "          </p>" & vbcrlf

 response.write "          <form name=""transferData"" id=""transferData"" method=""post"" action=""datamgr_transfer_field_data_action.asp"">" & vbcrlf
 response.write "          <p>" & vbcrlf
 response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
 response.write "            <tr valign=""top"">" & vbcrlf
 response.write "                <td>" & vbcrlf
 response.write "                    <input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to List"" class=""button"" onclick=""location.href='datamgr_sections_list.asp'"" />" & vbcrlf
 response.write "                    <input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""validateFields()"" />" & vbcrlf
 response.write "                </td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
 response.write "          </table>" & vbcrlf
 response.write "          </p>" & vbcrlf

                           displayDMFieldValues session("orgid"), lcl_sc_dm_typeid

 response.write "          </form>" & vbcrlf
 'response.write "          </div>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"--> 
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayDMFieldValues(p_orgid, p_sc_dm_typeid)
 	Dim iRowCount, iTotalCount

  iRowCount     = 0
  iTotalCount   = 0
  sOrgID        = 0
  sSC_DM_TypeID = ""

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_sc_dm_typeid <> "" then
     sSC_DM_TypeID = clng(p_sc_dm_typeid)
  end if

  sSQL = sSQL & "SELECT DISTINCT "
  sSQL = sSQL & " dmtf.orgid, "
  sSQL = sSQL & " dmtf.dm_typeid, "
  sSQL = sSQL & " dmt.description dm_typeid_description, "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQl = sSQL & " dms.sectionname, "
  sSQL = sSQL & " dmtf.dm_fieldid, "
  sSQL = sSQL & " dmsf.fieldname "
  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_types dmt "
  sSQL = sSQL &            " ON dmt.dm_typeid = dmtf.dm_typeid "
  sSQL = sSQL &            " AND dmt.isActive = 1 "
  sSQL = sSQL &            " AND dmt.isTemplate = 0 "
  sSQL = sSQL &            " AND dmt.orgid = " & sOrgID
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &            " ON dmts.dm_sectionid = dmtf.dm_sectionid "
  sSQL = sSQL &            " AND dmts.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
  sSQL = sSQL &            " ON dms.sectionid = dmts.sectionid "
  sSQL = sSQL &            " AND dms.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &            " ON dmsf.section_fieldid = dmtf.section_fieldid "
  sSQL = sSQL &            " AND dmsf.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_values dmv "
  sSQL = sSQL &            " ON dmv.dm_typeid = dmtf.dm_typeid "
  sSQL = sSQL &            " AND dmv.dm_sectionid = dmtf.dm_sectionid "
  sSQL = sSQL &            " AND dmv.dm_fieldid = dmtf.dm_fieldid "
  sSQL = sSQL &            " AND dmv.orgid = " & sOrgID
  sSQL = sSQL & " WHERE dmtf.orgid = " & sOrgID

  if sSC_DM_TypeID <> "" then
     sSQL = sSQL & " AND dmtf.dm_typeid = " & sSC_DM_TypeID
  end if

  sSQL = sSQL & " ORDER BY dmt.description, dms.sectionname, dmsf.fieldname "

 	set oDMTransferFields = Server.CreateObject("ADODB.Recordset")
	 oDMTransferFields.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMTransferFields.eof then
     lcl_bgcolor            = "#ffffff"
     lcl_previous_dm_typeid = 0

     do while not oDMTransferFields.eof
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount   = iRowCount   + 1
        iTotalCount = iTotalCount + 1

        if lcl_previous_dm_typeid <> oDMTransferFields("dm_typeid") then
           if iRowCount > 1 then
            		response.write "</table>" & vbcrlf
              response.write "<div align=""right""><strong>Total Fields: </strong>[" & iRowCount & "]</div>" & vbcrlf
              response.write "</p>" & vbcrlf

              iRowCount = 1
           end if

           response.write "<p>" & vbcrlf
       		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:800px"">" & vbcrlf
         		response.write "  <tr align=""left"">" & vbcrlf
           response.write "      <th nowrap=""nowrap"">DM Type</th>" & vbcrlf
           response.write "      <th nowrap=""nowrap"">Section</th>" & vbcrlf
           response.write "      <th nowrap=""nowrap"">Field</th>" & vbcrlf
           response.write "      <th>Transfer Data to Field...</th>" & vbcrlf
           response.write "  </tr>" & vbcrlf
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" align=""left"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"">" & oDMTransferFields("dm_typeid_description") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"">" & oDMTransferFields("sectionname") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"">" & oDMTransferFields("fieldname") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"">" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_sectionid" & iTotalCount & """ id=""dm_sectionid" & iTotalCount & """ value=""" & oDMTransferFields("dm_sectionid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_fieldid" & iTotalCount & """ id=""dm_fieldid" & iTotalCount & """ value=""" & oDMTransferFields("dm_fieldid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
        response.write "          <select name=""transfer_field_data" & iTotalCount & """ id=""transfer_field_data" & iTotalCount & """ class=""transferFieldData"" onchange=""setupTransferData('" & iTotalCount & "');"">" & vbcrlf
        response.write "            <option value="""">&nbsp;</option>" & vbcrlf
                                    displayTransferFieldOptions sOrgID, oDMTransferFields("dm_typeid")
        response.write "          </select>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

        lcl_previous_dm_typeid = oDMTransferFields("dm_typeid")

        oDMTransferFields.movenext
     loop

     if iRowCount > 1 then
      		response.write "</table>" & vbcrlf
        response.write "<div align=""right""><strong>Total Fields: </strong>[" & iRowCount & "]</div>" & vbcrlf
        response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""" & iTotalCount & """ size=""3"" maxlength=""10"" />" & vbcrlf
        response.write "</p>" & vbcrlf
     end if

  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Records Available.</p>" & vbcrlf
  end if

 	oDMTransferFields.close
 	set oDMTransferFields = nothing

end sub

'------------------------------------------------------------------------------
sub displayDMTypeOptions(iOrgID, iSC_DMTypeID)

  sOrgID                 = 0
  sSC_DMTypeID           = ""
  lcl_selected_dm_typeid = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iSC_DMTypeID <> "" then
     sSC_DMTypeID = clng(iSC_DMTypeID)
  end if

  sSQL = "SELECT dm_typeid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND isActive = 1 "
  sSQL = sSQL & " AND isTemplate = 0 "

'  if sSC_DMTypeID <> "" then
'     sSQL = sSQL & " AND dm_typeid = " & sSC_DMTypeID
'  end if

  sSQL = sSQL & " ORDER BY description "

 	set oDisplayDMTypeOptions = Server.CreateObject("ADODB.Recordset")
	 oDisplayDMTypeOptions.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayDMTypeOptions.eof then
     do while not oDisplayDMTypeOptions.eof

        if sSC_DMTypeID = oDisplayDMTypeOptions("dm_typeid") then
           lcl_selected_dm_typeid = " selected=""selected"""
        else
           lcl_selected_dm_typeid = ""
        end if

        response.write "  <option value=""" & oDisplayDMTypeOptions("dm_typeid") & """" & lcl_selected_dm_typeid & ">" & oDisplayDMTypeOptions("description") & "</option>" & vbcrlf

        oDisplayDMTypeOptions.movenext
     loop
  end if

  oDisplayDMTypeOptions.close
  set oDisplayDMTypeOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayTransferFieldOptions(iOrgID, iSC_DM_TypeID)
  sOrgID        = 0
  sSC_DM_TypeID = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iSC_DM_TypeID <> "" then
     sSC_DM_TypeID = clng(iSC_DM_TypeID)
  end if

  sSQL = sSQL & "SELECT DISTINCT "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQl = sSQL & " dms.sectionname, "
  sSQL = sSQL & " dmtf.dm_fieldid, "
  sSQL = sSQL & " dmsf.fieldname "
  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_types dmt "
  sSQL = sSQL &            " ON dmt.dm_typeid = dmtf.dm_typeid "
  sSQL = sSQL &            " AND dmt.isActive = 1 "
  sSQL = sSQL &            " AND dmt.isTemplate = 0 "
  sSQL = sSQL &            " AND dmt.orgid = " & sOrgID
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &            " ON dmts.dm_sectionid = dmtf.dm_sectionid "
  sSQL = sSQL &            " AND dmts.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
  sSQL = sSQL &            " ON dms.sectionid = dmts.sectionid "
  sSQL = sSQL &            " AND dms.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &            " ON dmsf.section_fieldid = dmtf.section_fieldid "
  sSQL = sSQL &            " AND dmsf.isActive = 1 "
  sSQL = sSQL & " WHERE dmtf.orgid = " & sOrgID
  sSQL = sSQL & " AND dmtf.dm_typeid = " & sSC_DM_TypeID
  sSQL = sSQL & " ORDER BY dms.sectionname, dmsf.fieldname "

 	set oDMTransferFieldsOptions = Server.CreateObject("ADODB.Recordset")
	 oDMTransferFieldsOptions.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMTransferFieldsOptions.eof then
     lcl_bgcolor            = "#ffffff"

     do while not oDMTransferFieldsOptions.eof

        response.write "  <option value=""dmsectionid" & oDMTransferFieldsOptions("dm_sectionid") & "_dmfieldid" & oDMTransferFieldsOptions("dm_fieldid") & """>" & oDMTransferFieldsOptions("sectionname") & ": " & oDMTransferFieldsOptions("fieldname") & "</option>" & vbcrlf

        oDMTransferFieldsOptions.movenext
     loop
  end if

  oDMTransferFieldsOptions.close
  set oDMTransferFieldsOptions = nothing

end sub
%>
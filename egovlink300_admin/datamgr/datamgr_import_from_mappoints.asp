<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_import_from_mappoints.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the fields for a MapPointType and allows the admin
'               to select which DataManger section/field to transfer the data into.
'
' MODIFICATION HISTORY
' 1.0  11/16/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel             = "../"  'Override of value from common.asp
 lcl_isRootAdmin    = False
 lcl_feature        = "datamgr_maint"
 lcl_url_parameters = ""

'Determine if the parent feature is "offline"
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

 if request("f") <> "" then
    lcl_feature = request("f")

   'Build return parameters
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Build page variables
 lcl_featurename = getFeatureName(lcl_feature)
 lcl_pagetitle   = lcl_featurename & ": Import Data From MapPoints"

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

'Check for import options
 lcl_mappoint_typeid = ""
 lcl_dm_typeid       = ""

 if request("mappoint_typeid") <> "" then
    lcl_mappoint_typeid = request("mappoint_typeid")
    lcl_mappoint_typeid = clng(lcl_mappoint_typeid)
 end if

 if request("dm_typeid") <> "" then
    lcl_dm_typeid = request("dm_typeid")
    lcl_dm_typeid = clng(lcl_dm_typeid)
 end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

<style type="text/css">
  .instructions {
     color:     #ff0000;
     font-size: 11pt;
  }

  .redText {
     color: #ff0000;
  }

  .importoptions_label {
     white-space: nowrap;
     text-align:  center;
  }

  .importoptions_dropdown {
     width: 100%;
  }

</style>

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

<script language="javascript">
<!--

$(document).ready(function(){
  $('#searchButton').prop('disabled',true);

  if($('#transferDataSearch_mappoint_typeid').prop('selectedIndex') > 0) {
     $('#searchButton').prop('disabled',false);
  }

  $('#transferDataSearch_mappoint_typeid').change(function(){
     if($('#transferDataSearch_mappoint_typeid').prop('selectedIndex') > 0) {
        $('#searchButton').prop('disabled',false);
     } else {
        $('#searchButton').prop('disabled',true);
     }
  });

});

function startImport() {
  var lcl_mappoint_typeid  = $('#transferDataSearch_mappoint_typeid').val();
  var lcl_dm_typeid        = $('#transferDataSearch_dm_typeid').val();
  var lcl_totalfields      = $('#totalfields').val();
  var lcl_overrideValuesID = document.getElementById('overrideValues');
  var lcl_overrideValues   = 'N';

  if(lcl_overrideValuesID.checked) {
     lcl_overrideValues = 'Y';
  }

//lcl_url  = 'datamgr_import_from_mappoints_action.asp';
//lcl_url += '?userid=<%=session("userid")%>';
//lcl_url += '&orgid=<%=session("orgid")%>';
//lcl_url += '&mappoint_typeid=' + lcl_mappoint_typeid;
//lcl_url += '&dm_typeid=' + lcl_dm_typeid;
//lcl_url += '&totalfields=' + lcl_totalfields;
//lcl_url += '&action=CREATE_DMID';
//lcl_url += '&isAjax=Y';

//alert(lcl_url);
  $('#importresults').html('Creating DataMgr Records...');
  $('#importresults_values').html('');

  $.post('datamgr_import_from_mappoints_action.asp', {
     userid:          '<%=session("userid")%>',
     orgid:           '<%=session("orgid")%>',
     mappoint_typeid: lcl_mappoint_typeid,
     dm_typeid:       lcl_dm_typeid,
     totalfields:     lcl_totalfields,
     action:          'CREATE_DMID',
     overridevalues:  lcl_overrideValues,
     isAjax:          'Y'
  }, function(result) {
//     displayScreenMsg(result);
//     $('#display_fieldvalue'+lcl_id).html(lcl_fieldvalue);
//     window.opener.location.reload();
//     $('#editHoursInfo'+lcl_id).slideUp('slow',function() {});

     $('#importresults').html(result);

     if(result.indexOf('INVALID VALUE') < 0) {

        //Cycle through DMT Field rows and determine how many have been selected
        if(lcl_totalfields > 0) {
           var i = 1;
           var lcl_mp_fieldid          = '';
           var lcl_transfer_field_data = '';
           var lcl_results             = '';

           for(i = 1; i <= lcl_totalfields; i++) {
               lcl_mp_fieldid          = $('#mp_fieldid' + i).val();
               lcl_transfer_field_data = $('#transfer_field_data' + i).val();

               if(lcl_transfer_field_data != '') {
//lcl_url  = 'datamgr_import_from_mappoints_action.asp';
//lcl_url += '?userid=<%=session("userid")%>';
//lcl_url += '&orgid=<%=session("orgid")%>';
//lcl_url += '&mappoint_typeid=' + lcl_mappoint_typeid;
//lcl_url += '&dm_typeid=' + lcl_dm_typeid;
//lcl_url += '&mp_fieldid=' + lcl_mp_fieldid;
//lcl_url += '&transfer_field_data=' + lcl_transfer_field_data;
//lcl_url += '&totalfields=' + lcl_totalfields;
//lcl_url += '&action=IMPORT_MP_VALUES';
//lcl_url += '&isAjax=Y';

//alert(lcl_url);

                  $.post('datamgr_import_from_mappoints_action.asp', {
                     userid:              '<%=session("userid")%>',
                     orgid:               '<%=session("orgid")%>',
                     mappoint_typeid:     lcl_mappoint_typeid,
                     dm_typeid:           lcl_dm_typeid,
                     mp_fieldid:          lcl_mp_fieldid,
                     transfer_field_data: lcl_transfer_field_data,
                     totalfields:         lcl_totalfields,
                     action:              'IMPORT_MP_VALUES',
                     overridevalues:      lcl_overrideValues,
                     isAjax:              'Y'
                  }, function(result) {
                     lcl_results  = $('#importresults_values').html();
                     lcl_results += result;

                     $('#importresults_values').html(lcl_results);
                  });
               }
           }
        }
     }
  });

//  $('#transferData_mappoint_typeid').val(lcl_mappoint_typeid);
//  $('#transferData_dm_typeid').val(lcl_dm_typeid);

//  document.getElementById("transferData").submit();
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
 response.write "    <p>" & vbcrlf
 response.write "    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""1000px"">" & vbcrlf
 response.write "      <tr>" & vbcrlf
 response.write "          <td><font size=""+1""><strong>" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
 response.write "          <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;"">&nbsp;</span></td>" & vbcrlf
 response.write "      </tr>" & vbcrlf
 response.write "      <tr>" & vbcrlf
 response.write "          <td colspan=""2""><input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to List"" class=""button"" onclick=""location.href='datamgr_list.asp" & lcl_url_parameters & "'"" /></td>" & vbcrlf
 response.write "      </tr>" & vbcrlf
 response.write "    </table>" & vbcrlf
 response.write "    </p>" & vbcrlf

'BEGIN: Import: Step 1 --------------------------------------------------------
 response.write "    <form name=""transferDataSeach"" id=""transferDataSearch"" method=""post"" action=""datamgr_import_from_mappoints.asp"">" & vbcrlf
 response.write "      <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
 response.write "    <p>" & vbcrlf
 response.write "    <fieldset name=""searchoptions"" id=""searchoptions"" class=""fieldset"">" & vbcrlf
 response.write "      <legend>Step 1</legend>" & vbcrlf
 response.write "      <div class=""instructions"">Select the [MapPoint Type] to import data from and then the [DataMgr Type] we are importing to.</div><br />"
 response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""importoptions_label"">" & vbcrlf
 response.write "                MapPoint Type:<br /><span class=""redText"">(import from)</span>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td class=""importoptions_dropdown"">" & vbcrlf
 response.write "                <select name=""mappoint_typeid"" id=""transferDataSearch_mappoint_typeid"">" & vbcrlf
 response.write "                  <option value="""">&nbsp;</option>" & vbcrlf
                                   displayMPTypeOptions session("orgid"), lcl_mappoint_typeid
 response.write "                </select>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""importoptions_label"">" & vbcrlf
 response.write "                DM Type:<br /><span class=""redText"">(import to)</span>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td class=""importoptions_dropdown"">" & vbcrlf
 response.write "                <select name=""dm_typeid"" id=""transferDataSearch_dm_typeid"">" & vbcrlf
 response.write "                  <option value="""">&nbsp;</option>" & vbcrlf
                                   displayDMTypeOptions session("orgid"), lcl_dm_typeid
 response.write "                </select>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "        <tr>" & vbcrlf
 response.write "            <td colspan=""2"">" & vbcrlf
 response.write "                <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Retrieve MapPoint Fields"" class=""button"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "      </table>" & vbcrlf
 response.write "    </fieldset>" & vbcrlf
 response.write "    </p>" & vbcrlf
 response.write "    </form>" & vbcrlf
'END: Import: Step 1 ----------------------------------------------------------

 if lcl_mappoint_typeid <> "" then
   'BEGIN: Import: Step 2 -----------------------------------------------------
    response.write "    <form name=""transferData"" id=""transferData"" method=""post"" action=""datamgr_import_from_mappoints_action.asp"">" & vbcrlf
    response.write "      <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
    response.write "      <input type=""hidden"" name=""mappoint_typeid"" id=""transferData_mappoint_typeid"" value="""" size=""5"" maxlength=""50"" />" & vbcrlf
    response.write "      <input type=""hidden"" name=""dm_typeid"" id=""transferData_dm_typeid"" value="""" size=""10"" maxlength=""50"" />" & vbcrlf
    response.write "    <p>" & vbcrlf
    response.write "    <fieldset name=""searchoptions"" id=""searchoptions"" class=""fieldset"">" & vbcrlf
    response.write "      <legend>Step 2</legend>" & vbcrlf
    response.write "      <div class=""instructions"">" & vbcrlf
    response.write "        For each MapPoint field, select the DataMgr field that it is to be imported into.<br />" & vbcrlf
    response.write "        <p><strong>Note: </strong>If a DataMgr field (dropdown value) is not selected for a MapPoint field (row) " & vbcrlf
    response.write "              then any/all values for that MapPoint field will NOT be imported.</p>" & vbcrlf
    response.write "      </div><br />"
                          displayMPFieldValues session("orgid"), lcl_mappoint_typeid, lcl_dm_typeid
    response.write "    </fieldset>" & vbcrlf
    response.write "    </p>" & vbcrlf
   'END: Import: Step 2 -------------------------------------------------------

   'BEGIN: Import: Step 3 -----------------------------------------------------
    response.write "    <p>" & vbcrlf
    response.write "    <fieldset name=""searchoptions"" id=""searchoptions"" class=""fieldset"">" & vbcrlf
    response.write "      <legend>Step 3</legend>" & vbcrlf
    response.write "      <div class=""instructions"">Click the [Begin Import] button when you are ready to start the import.</div><br />"
    response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
    response.write "        <tr valign=""top"">" & vbcrlf
    response.write "            <td>" & vbcrlf
    response.write "                <input type=""checkbox"" name=""overrideValues"" id=""overrideValues"" value=""Y"" checked=""checked"" /> Override existing values&nbsp;&nbsp;" & vbcrlf
    response.write "                <input type=""button"" name=""importButton"" id=""importButton"" value=""Begin Import"" class=""button"" onclick=""startImport()"" />" & vbcrlf
    response.write "                <p><div id=""importresults""></div></p>" & vbcrlf
    response.write "                <p><div id=""importresults_values""></div></p>" & vbcrlf
    response.write "            </td>" & vbcrlf
    response.write "        </tr>" & vbcrlf
    response.write "      </table>" & vbcrlf
    response.write "      </fieldset>" & vbcrlf
    response.write "    </p>" & vbcrlf
    response.write "    </form>" & vbcrlf
   'END: Import: Step 3 -------------------------------------------------------
 end if

 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"--> 
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayMPFieldValues(p_orgid, p_mappoint_typeid, p_dm_typeid)
 	Dim iRowCount, iTotalCount

  iRowCount   = 0
  iTotalCount = 0
  sOrgID      = 0
  sMPTypeID   = ""
  sDMTypeID   = ""

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_mappoint_typeid <> "" then
     sMPTypeID = clng(p_mappoint_typeid)
  end if

  if p_dm_typeid <> "" then
     sDMTypeID = clng(p_dm_typeid)
  end if

  if sMPTypeID <> "" then
     sSQL = "SELECT "
     sSQL = sSQL & " mptf.mappoint_typeid, "
     sSQL = sSQL & " mpt.description, "
     sSQL = sSQL & " mptf.mp_fieldid, "
     sSQL = sSQL & " mptf.fieldname "
     'sSQL = sSQL & " mptf.fieldtype, "
     'sSQL = sSQL & " mptf.displayInResults, "
     'sSQL = sSQL & " mptf.displayInInfoPage, "
     'sSQL = sSQL & " mptf.resultsOrder, "
     'sSQL = sSQL & " mptf.inPublicSearch "
     sSQL = sSQL & " FROM egov_mappoints_types_fields mptf "
     sSQL = sSQL &      " INNER JOIN egov_mappoints_types mpt ON mpt.mappoint_typeid = mptf.mappoint_typeid "
     sSQL = sSQL & " WHERE mptf.orgid = " & sOrgID
     sSQL = sSQL & " AND mptf.mappoint_typeid = " & sMPTypeID

    	set oMPTransferFields = Server.CreateObject("ADODB.Recordset")
   	 oMPTransferFields.Open sSQL, Application("DSN"), 3, 1

     if not oMPTransferFields.eof then
        lcl_bgcolor            = "#ffffff"
        lcl_previous_mp_typeid = 0

        do while not oMPTransferFields.eof
           lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     	   		iRowCount   = iRowCount   + 1
           iTotalCount = iTotalCount + 1

           if lcl_previous_mp_typeid <> oMPTransferFields("mappoint_typeid") then
              if iRowCount > 1 then
               		response.write "</table>" & vbcrlf
                 response.write "<div align=""right""><strong>Total Fields: </strong>[" & iRowCount & "]</div>" & vbcrlf
                 response.write "</p>" & vbcrlf

                 iRowCount = 1
              end if

              response.write "<p>" & vbcrlf
          		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:800px"">" & vbcrlf
            		response.write "  <tr align=""left"">" & vbcrlf
              response.write "      <th nowrap=""nowrap"">MapPoint Type</th>" & vbcrlf
              response.write "      <th nowrap=""nowrap"">Field</th>" & vbcrlf
              response.write "      <th>Transfer Data to DM Field...</th>" & vbcrlf
              response.write "  </tr>" & vbcrlf
           end if

           response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" align=""left"" valign=""top"">" & vbcrlf
           response.write "      <td class=""formlist"">" & oMPTransferFields("description") & "</td>" & vbcrlf
           response.write "      <td class=""formlist"">" & oMPTransferFields("fieldname") & "</td>" & vbcrlf
           response.write "      <td class=""formlist"">" & vbcrlf
           'response.write "          <input type=""hidden"" name=""dm_sectionid" & iTotalCount & """ id=""dm_sectionid" & iTotalCount & """ value=""" & oDMTransferFields("dm_sectionid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
           'response.write "          <input type=""hidden"" name=""dm_fieldid" & iTotalCount & """ id=""dm_fieldid" & iTotalCount & """ value=""" & oDMTransferFields("dm_fieldid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
           response.write "          <input type=""hidden"" name=""mp_fieldid" & iTotalCount & """ id=""mp_fieldid" & iTotalCount & """ value=""" & oMPTransferFields("mp_fieldid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
           response.write "          <select name=""transfer_field_data" & iTotalCount & """ id=""transfer_field_data" & iTotalCount & """ class=""transferFieldData"">" & vbcrlf
           response.write "            <option value="""">&nbsp;</option>" & vbcrlf
                                       displayTransferFieldOptions sOrgID, sDMTypeID
           response.write "          </select>" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "  </tr>"  & vbcrlf

           lcl_previous_mp_typeid = oMPTransferFields("mappoint_typeid")

           oMPTransferFields.movenext
        loop

        if iRowCount > 1 then
         		response.write "</table>" & vbcrlf
           response.write "<div align=""right""><strong>Total Fields: </strong>[" & iRowCount & "]</div>" & vbcrlf
           response.write "<input type=""hidden"" name=""totalfields"" id=""totalfields"" value=""" & iTotalCount & """ size=""3"" maxlength=""10"" />" & vbcrlf
           response.write "</p>" & vbcrlf
        end if

     end if

     oMPTransferFields.close
     set oMPTransferFields = nothing

  end if

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
sub displayMPTypeOptions(iOrgID, iMPTypeID)

  sOrgID                 = 0
  sMPTypeID              = ""
  lcl_selected_mp_typeid = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iMPTypeID <> "" then
     sMPTypeID = clng(iMPTypeID)
  end if

  sSQL = "SELECT mappoint_typeid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND isActive = 1 "
  sSQL = sSQL & " ORDER BY description "

 	set oDisplayMPTypeOptions = Server.CreateObject("ADODB.Recordset")
	 oDisplayMPTypeOptions.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayMPTypeOptions.eof then
     do while not oDisplayMPTypeOptions.eof

        if sMPTypeID = oDisplayMPTypeOptions("mappoint_typeid") then
           lcl_selected_mp_typeid = " selected=""selected"""
        else
           lcl_selected_mp_typeid = ""
        end if

        response.write "  <option value=""" & oDisplayMPTypeOptions("mappoint_typeid") & """" & lcl_selected_mp_typeid & ">" & oDisplayMPTypeOptions("description") & "</option>" & vbcrlf

        oDisplayMPTypeOptions.movenext
     loop
  end if

  oDisplayMPTypeOptions.close
  set oDisplayMPTypeOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayTransferFieldOptions(iOrgID, iDM_TypeID)
  sOrgID        = 0
  sDM_TypeID = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDM_TypeID <> "" then
     sDM_TypeID = clng(iDM_TypeID)
  end if

  sSQL = sSQL & "SELECT DISTINCT "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQl = sSQL & " dms.sectionname, "
  sSQL = sSQL & " dmtf.dm_fieldid, "
  sSQL = sSQL & " dmtf.section_fieldid, "
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
  sSQL = sSQL & " AND dmtf.dm_typeid = " & sDM_TypeID
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
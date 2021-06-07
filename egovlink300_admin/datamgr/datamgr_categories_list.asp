<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_categories_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the categories for a DM Type
'
' MODIFICATION HISTORY
' 1.0 04/25/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Determine if the parent feature is "offline"
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_isRootAdmin = False
 lcl_feature     = "datamgr_types_maint"
 lcl_pagetitle   = ""
 lcl_dm_typeid   = 0

 if request("f") <> "" AND request("f") <> "datamgr_types_maint" then
    lcl_feature = request("f")
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the DM_TypeID
 if request("dm_typeid") <> "" then
    lcl_dm_typeid = request("dm_typeid")
 else
    lcl_dm_typeid = getDMTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

    if lcl_dm_typeid = 0 then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 lcl_dm_typeid = clng(lcl_dm_typeid)
 lcl_pagetitle = getFeatureName(lcl_feature)
 lcl_pagetitle = lcl_pagetitle & " [Maintain Categories]"
 lcl_success   = request("success")

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f",         lcl_feature)
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dm_typeid", lcl_dm_typeid)

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
    lcl_onload = lcl_onload & "window.opener.location.reload();"
 end if

'Check for org features
 lcl_orghasfeature_feature              = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain     = orghasfeature(lcl_feature)

'Check for user permissions
 lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
 lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

'Retrieve the search options
 lcl_sc_categoryname = ""
' lcl_sc_fromcreatedate = ""
' lcl_sc_tocreatedate   = ""
' lcl_sc_title          = ""
' lcl_sc_userid         = 0
' lcl_sc_orderby        = "createdate"

 if request("sc_categoryname") <> "" then
    lcl_sc_categoryname = request("sc_categoryname")
 end if

' if request("sc_fromcreatedate") <> "" then
'    lcl_sc_fromcreatedate = request("sc_fromcreatedate")
' end if

' if request("sc_tocreatedate") <> "" then
'    lcl_sc_tocreatedate = request("sc_tocreatedate")
' end if

' if request("sc_title") <> "" then
'    lcl_sc_title = request("sc_title")
' end if

' if request("sc_userid") <> "" then
'    lcl_sc_userid = request("sc_userid")
' end if

' if request("sc_orderby") <> "" then
'    lcl_sc_orderby = request("sc_orderby")
' end if

  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_categoryname", lcl_sc_categoryname)
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--
function confirmDelete(p_id) {
  lcl_cname = document.getElementById("category"+p_id).innerHTML;

  var r = confirm("Are you sure you want to delete this category: '" + lcl_cname + "'?");
 	if (r == true) { 
      <%
        lcl_delete_params = lcl_url_parameters
        lcl_delete_params = setupUrlParameters(lcl_delete_params, "user_action", "DELETE")
      %>

		  		location.href="datamgr_categories_action.asp<%=lcl_delete_params%>&categoryid=" + p_id;
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

<style>
  .body {
     background-color:#ffffff;
  }

  #category_content {
     margin-top:  20px;
     margin-left: 20px;
  }

  #closeWindow {
     margin-bottom: 10px;
  }
</style>
</head>
<body class="body" onload="<%=lcl_onload%>">
<%
  response.write "<form name=""datamgr"" id=""datamgr"" action=""datamgr_categories_list.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" value=""" & lcl_dm_typeid & """ size=""5"" maxlength=""5"" />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <div id=""category_content"">" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td><font size=""+1""><strong>" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
  response.write "                  <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;"">&nbsp;</span></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                    <input type=""button"" name=""closeWindow"" id=""closeWindow"" class=""button"" value=""Close Window"" onclick=""parent.close();"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <fieldset class=""fieldset"">" & vbcrlf
  response.write "                        <legend>Search Options&nbsp;</legend>" & vbcrlf
  response.write "                        <p>" & vbcrlf
  response.write "                        <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                          <tr>" & vbcrlf
  response.write "                              <td>" & vbcrlf
  response.write "                                  Category:" & vbcrlf
  response.write "                                  <input type=""text"" name=""sc_categoryname"" id=""sc_categoryname"" size=""40"" maxlength=""100"" value=""" & lcl_sc_categoryname & """ />" & vbcrlf
  response.write "                              </td>" & vbcrlf
  response.write "                          </tr>" & vbcrlf
  response.write "                        </table>" & vbcrlf
  response.write "                        </p>" & vbcrlf
  response.write "                        <p><input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" /></p>" & vbcrlf
  response.write "                      </fieldset>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf

 'If a DM Type has NOT been created then do NOT allow a DM Data record to be added.
  if lcl_dm_typeid > 0 then
     response.write "            <p>" & vbcrlf
     response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "              <tr valign=""top"">" & vbcrlf
     response.write "                  <td><input type=""button"" name=""addButton"" id=""addButton"" value=""Add Category"" class=""button"" onclick=""location.href='datamgr_categories_maint.asp" & lcl_url_parameters & "'"" /></td>" & vbcrlf
     response.write "              </tr>" & vbcrlf
     response.write "            </table>" & vbcrlf
     response.write "            </p>" & vbcrlf
  end if

  displayDMCategories lcl_isRootAdmin, session("orgid"), lcl_feature, lcl_dm_typeid, lcl_sc_categoryname, lcl_url_parameters

  response.write "            </p>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
	%>
<!--#Include file="../admin_footer.asp"--> 
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayDMCategories(p_isRootAdmin, p_orgid, p_feature, p_dm_typeid, p_sc_categoryname, p_url_parameters)
 	Dim iRowCount

  lcl_sc_cname = ""

  if trim(p_sc_categoryname) <> "" then
     lcl_sc_cname = trim(p_sc_categoryname)
     lcl_sc_cname = ucase(lcl_sc_cname)
     lcl_sc_cname = "'%" & dbsafe(lcl_sc_cname) & "%'"
  end if

  sSQL = "SELECT dmc.categoryid, "
  sSQL = sSQL & " dmc.categoryname, "
  sSQL = sSQL & " dmc.dm_typeid, "
  sSQL = sSQL & " dmc.isActive, "
  sSQL = sSQL & " dmc.createdbyid, "
  sSQL = sSQL & " dmc.createdbydate, "
  sSQL = ssQL & " dmc.lastmodifiedbyid, "
  sSQL = sSQL & " dmc.lastmodifiedbydate, "
  sSQL = sSQL & " dmc.parent_categoryid, "
  sSQL = sSQL & " dmc.mappointcolor, "
  sSQL = sSQL & " (select count(dmc2.categoryid) "
  sSQL = sSQL &  " from egov_dm_categories dmc2 "
  sSQL = sSQL &  " where dmc2.parent_categoryid = dmc.categoryid) as total_sub_categories "
  sSQL = sSQL & " FROM egov_dm_categories dmc "
  sSQL = sSQL & " WHERE dmc.orgid = " & p_orgid
  sSQL = sSQL & " AND dmc.dm_typeid = " & p_dm_typeid
  sSQL = ssQL & " AND dmc.parent_categoryid = 0 "

 'Setup the WHERE clause with the search option values.
  if lcl_sc_cname <> "" then
     sSQL = sSQL & " AND upper(dmc.categoryname) like (" & lcl_sc_cname & ") "
  end if

  sSQL = sSQL & " ORDER BY dmc.categoryname "

 	set oDMCategories = Server.CreateObject("ADODB.Recordset")
	 oDMCategories.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMCategories.eof then
   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" width=""100%"">" & vbcrlf
   		response.write "  <tr valign=""bottom"">" & vbcrlf
     response.write "      <th align=""left"">Category</th>" & vbcrlf
     response.write "      <th>Active</th>" & vbcrlf
     response.write "      <th>Total<br />Sub-Categories</th>" & vbcrlf
     response.write "      <th align=""left"">Map Point Color</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     lcl_bgcolor             = "#ffffff"
     lcl_original_categoryid = 0

     do while not oDMCategories.eof
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount   = iRowCount + 1

       'Setup the onclick
        lcl_row_onclick = setupUrlParameters(p_url_parameters, "categoryid", oDMCategories("categoryid"))
        lcl_row_onclick = "location.href='datamgr_categories_maint.asp" & lcl_row_onclick & "';"

       'Check to see if this category has been associated to a DM Type to determine if the category can been deleted.
        lcl_categoryExistsOnDMType = checkForDefaultCategoryOnDMTypes(oDMCategories("categoryid"))

        if lcl_categoryExistsOnDMType then
           lcl_canDelete = False
        else
           lcl_canDelete = True
        end if

       'Set up the display fields
        lcl_display_active = "&nbsp;"

        if oDMCategories("isActive") then
           lcl_display_active = "Y"
        end if

        lcl_display_totalSubCategories = oDMCategories("total_sub_categories")

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """><span id=""category" & oDMCategories("categoryid") & """>" & oDMCategories("categoryname") & "</span></td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_active             & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_totalSubCategories & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & vbcrlf
        response.write "          <img src=""mappoint_colors/bg_" & oDMCategories("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />" & vbcrlf
        response.write            oDMCategories("mappointcolor") & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td class=""formlist"" align=""center"">" & vbcrlf

        if lcl_canDelete then
           response.write "<input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirmDelete('" & oDMCategories("categoryid") & "');"" />" & vbcrlf
        else
           response.write "&nbsp;" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

        oDMCategories.movenext
     loop

   		response.write "</table>" & vbcrlf
	    response.write "</div>" & vbcrlf

  else
   		response.write "<p class=""norecords"">No DM Categories Available.</p>" & vbcrlf
  end if

 	oDMCategories.close
 	set oDMCategories = nothing

end sub
%>
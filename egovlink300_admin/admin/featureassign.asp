<!-- #include file="../includes/common.asp" //-->
<%
 sLevel = "../"  'Override of value from common.asp

 if not UserIsRootAdmin(session("userid")) then
   	response.redirect sLevel & "default.asp"
 end if

 dim lcl_bgcolor, lcl_total_orgs, lcl_total_admins

 lcl_failed           = 0
 lcl_success          = ""
 lcl_display_msg      = ""
 lcl_newfeatureid     = CLng(request("featureid"))
 lcl_parentfeatureid1 = ""
 lcl_parentfeatureid2 = ""
 lcl_total_orgs       = 0
 lcl_total_admins     = 0

'Determine what type of assignments we are doing.
'1. ORG = Will assign feature selected to ANY/ALL orgs that do NOT have the feature already assigned.
'         It will also assign the feature select to the root admin users for ONLY orgs GETTING the new assignment.
'2. USER = Will assign the feature to all users in EVERY org that currently do NOT have the feature assigned.
 if trim(request("assigntype")) <> "" then
    lcl_assign_type = UCASE(trim(request("assigntype")))
 else
    lcl_assign_type = "ORG"
 end if

 if lcl_assign_type = "ORG" then
    lcl_pagetitle        = "Orgs"
    lcl_query_limitation = "NOT"
    lcl_assignToOrgs     = "Y"
    lcl_rootadmin        = "Y"
    lcl_userlabel        = "Root Admins"
 else
    lcl_pagetitle        = "Users"
    lcl_query_limitation = ""
    lcl_assignToOrgs     = "N"
    lcl_rootadmin        = "N"
    lcl_userlabel        = "Users"
 end if

'Determine where the user is coming from
 if trim(request("loc")) <> "" then
    lcl_return_location = UCASE(trim(request("loc")))
 else
    lcl_return_location = "HOME"
 end if

 if request("parentfeatureid1") <> "" then
    lcl_parentfeatureid1 = CLng(request("parentfeatureid1"))
 end if

 if request("parentfeatureid2") <> "" then
    lcl_parentfeatureid2 = CLng(request("parentfeatureid2"))
 end if

'Check for a screen message
 lcl_onload  = ""
 'lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Administration Consule {Assign Feature to <%=lcl_pagetitle%>}</title>
  <link rel="stylesheet" type="text/css" href="<%=sLevel%>menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="<%=sLevel%>global.css" />

  <script language="javascript" src="<%=sLevel%>scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
function validateFields() {

  var lcl_return_false = 0;

  if (document.getElementById("featureid").value == "") {
      document.getElementById("featureid").focus();
      inlineMsg(document.getElementById("featureid").id,'<strong>Required Field Missing: </strong> Feature to be assigned',10,'featureid');
      lcl_return_false = lcl_return_false + 1;
  }else{
      clearMsg("featureid");
  }

  if (lcl_return_false > 0) {
      return false;
  }else{
      document.getElementById("assignNewFeature").submit();
  }
}

function returnToPage(iFeatureID) {
<%
  if lcl_return_location = "HOME" then
     response.write "lcl_url = '../default.asp';"
  else
%>
     if(iFeatureID == "" || iFeatureID == "0") {
        lcl_url = 'featureselection.asp';
     }else{
        lcl_url = 'featureedit.asp?featureid=' + iFeatureID;
     }
<% end if %>

  location.href = lcl_url;
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
</script>
</head>
<body onload="<%=lcl_onload%>">
<%
    ShowHeader sLevel
%>
<!-- #include file="../menu/menu.asp" //--> 
<p>
<table border="0" cellspacing="0" cellpadding="2" style="margin-left:10px">
  <tr valign="top">
      <td width="60%">
          <font size="+1"><strong>Assign New Feature to <%=lcl_pagetitle%></strong></font><br />
          <input type="button" name="returnButton" id="returnButton" value="Return" class="button" onclick="returnToPage('<%=lcl_newfeatureid%>');" />
      </td>
      <td width="40%" align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
</table>
</p>
<p>
<table border="0" cellspacing="0" cellpadding="2" style="margin-left:10px">
  <form name="assignNewFeature" id="assignNewFeature" action="featureassign.asp" method="post">
    <input type="hidden" name="assigntype" id="assigntype" value="<%=lcl_assign_type%>" />
  <tr>
      <td>Feature to be assigned:</td>
      <td>
          <select name="featureid" id="featureid" onchange="clearMsg('featureid');">
            <% showFeatureOptions 0, lcl_newfeatureid, "Y" %>
          </select>
      </td>
  </tr>
  <tr><td colspan="2">&nbsp;</td></tr>
  <tr>
      <td colspan="2">
          <fieldset>
            <legend>Limit the Assignment to only those already assigned the following feature(s)&nbsp;</legend>
            <p>
            <table border="0" cellspacing="0" cellpadding="2">
              <tr>
                  <td>Feature:</td>
                  <td>
                      <select name="parentfeatureid1" id="parentfeatureid1">
                        <% showFeatureOptions 0, lcl_parentfeatureid1, "Y" %>
                      </select>
                  </td>
              </tr>
              <tr>
                  <td>Feature:</td>
                  <td>
                      <select name="parentfeatureid2" id="parentfeatureid2">
                        <% showFeatureOptions 0, lcl_parentfeatureid2, "Y" %>
                      </select>
                  </td>
              </tr>
            </table>
            </p>
          </fieldset>
      </td>
  </tr>
  <tr>
      <td colspan="2"><input type="button" name="submitButton" id="submitButton" value="Assign Feature" onclick="validateFields();" class="button" /></td>
  </tr>
  </form>
</table>
</p>
<%
'Assign the feature if the admin user has submitted the form.
 if request.ServerVariables("REQUEST_METHOD") = "POST" then
    lcl_failed = assignNewFeatures(lcl_parentfeatureid1, lcl_parentfeatureid2, lcl_newfeatureid, lcl_failed)

    if clng(lcl_failed) < 1 then
       lcl_success = "NC_ERROR"
    else
       lcl_success = "ASSIGN_SUCCESS"
    end if

    lcl_msg = setupScreenMsg(lcl_success)

    response.write "<script language=""javascript"">" & vbcrlf
    response.write "  displayScreenMsg('" & lcl_msg & "');" & vbcrlf
    response.write "</script>" & vbcrlf

 end if
%>
</body>
</html>
<%
'------------------------------------------------------------------------------
sub showFeatureOptions(iParentFeatureID, iFeatureID, iShowBlankOption)

 'Setup the orderby and option spacing
  if iParentFeatureID = 0 then
     lcl_orderby = "admindisplayorder"
     lcl_spacing = ""
  else
     lcl_orderby = "securitydisplayorder"
     lcl_spacing = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  end if

  if iShowBlankOption = "Y" then
     response.write "  <option value=""""></option>" & vbcrlf
  end if

  sSQL = "SELECT featureid, feature, featurename "
  sSQL = sSQL & " FROM egov_organization_features "
  sSQL = sSQL & " WHERE parentfeatureid = " & iParentFeatureID
  sSQL = sSQL & " ORDER BY " & lcl_orderby & ", featurename "

  set oFeatureOptions = Server.CreateObject("ADODB.Recordset")
  oFeatureOptions.Open sSQL, Application("DSN"), 3, 1

  if not oFeatureOptions.eof then
     do while not oFeatureOptions.eof
        if iFeatureID = CLng(oFeatureOptions("featureid")) then
           lcl_selected_feature = " selected=""selected"""
        else
           lcl_selected_feature = ""
        end if

        response.write "  <option value=""" & oFeatureOptions("featureid") & """" & lcl_selected_feature & ">" & lcl_spacing & oFeatureOptions("featurename") & " [" & oFeatureOptions("featureid") & " - " & oFeatureOptions("feature") & "]</option>" & vbcrlf

       'Check for any sub-level features
        showFeatureOptions oFeatureOptions("featureid"), iFeatureID, "N"

        oFeatureOptions.movenext
     loop
  end if

  set oFeatureOptions = nothing

end sub

'------------------------------------------------------------------------------
function assignNewFeatures(iParentFeatureID1, iParentFeatureID2, iNewFeatureID, iFailed)
  lcl_return           = 0
  lcl_parentfeatureid1 = iParentFeatureID1
  lcl_parentfeatureid2 = iParentFeatureID2
  lcl_newfeatureid     = iNewFeatureID
  lcl_orgids           = getOrgIDs(lcl_parentfeatureid1, lcl_parentfeatureid2, lcl_newfeatureid)

  sSQL = "SELECT distinct orgid, orgname "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid " & lcl_query_limitation & " IN (select distinct orgid "
  sSQL = sSQL &                     " from egov_organizations_to_features "
  sSQL = sSQL &                     " where featureid = " & iNewFeatureID & ") "
  sSQL = sSQL & " AND orgid IN (" & lcl_orgids & ") "
  sSQL = sSQL & " ORDER BY orgid "

  set oOrgs = Server.CreateObject("ADODB.Recordset")
  oOrgs.Open sSQL, Application("DSN"), 3, 1

  if not oOrgs.eof then

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width:600px; margin-left:10px"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <th>Organizations</th>" & vbcrlf
     response.write "      <th>" & lcl_userlabel & "</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     lcl_bgcolor    = "#ffffff"

     do while not oOrgs.eof
        lcl_bgcolor    = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_total_orgs = lcl_total_orgs + 1

        if lcl_assignToOrgs = "Y" then
           sSQL = "INSERT INTO egov_organizations_to_features (featureid, orgid, value, featurename, featuredescription, "
           sSQL = sSQL & " publicurl, publicdisplayorder, publicimageurl, publiccanview, hasadminview, admindisplayorder) VALUES ("
           sSQL = sSQL & lcl_newfeatureid & ", "
           sSQL = sSQL & oOrgs("orgid")   & ", "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "0, "
           sSQL = sSQL & "NULL, "
           sSQL = sSQL & "NULL "
           sSQL = sSQL & ")"

           set oFeatureInsert = Server.CreateObject("ADODB.Recordset")
           oFeatureInsert.Open sSQL, Application("DSN") , 3, 1

           set oFeatureInsert = nothing
        end if

       'Display the row for the org just inserted.
        response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "      <td>" & oOrgs("orgname") & " [" & oOrgs("orgid")& "]</td>" & vbcrlf
        response.write "      <td>&nbsp;</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        assignUsers oOrgs("orgid"), lcl_parentfeatureid1, lcl_parentfeatureid2, lcl_newfeatureid, lcl_rootadmin

        oOrgs.movenext
     loop

     lcl_return  = clng(iFailed) + 1
     lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

     response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
     response.write "      <td><strong>Total Orgs Assigned:</strong> ["                  & lcl_total_orgs   & "]</td>" & vbcrlf
     response.write "      <td><strong>Total " & lcl_userlabel & " Assigned:</strong> [" & lcl_total_admins & "]</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

  end if

  assignNewFeatures = lcl_return

end function

'------------------------------------------------------------------------------
function getOrgIDs(iParentFeatureID1, iParentFeatureID2, iNewFeatureID)

  lcl_return                   = ""
  lcl_parentorgids             = ""
  lcl_neworgids                = ""
  lcl_parentfeature_limitation = ""

 'Setup the Parent Features ID(s) to search on

 'Get all of the features with these parent feature ids assigned
  if iParentFeatureID1 <> "" OR iParentFeatureID2 <> "" then

    'This causes the script to NOT assign this feature to ALL orgids in case it cannot find any left to assign it to.
     lcl_parentorgids = "0"

    'Get all of the orgs for the first feature if one has been selected.
     if iParentFeatureID1 <> "" then
        sSQL = " SELECT distinct orgid "
        sSQL = sSQL & " FROM egov_organizations_to_features "
        sSQL = sSQL & " WHERE featureid = " & iParentFeatureID1
        sSQL = sSQL & " ORDER BY orgid "

        set oOrgsAssigned = Server.CreateObject("ADODB.Recordset")
        oOrgsAssigned.Open sSQL, Application("DSN"), 3, 1

        if not oOrgsAssigned.eof then
           lcl_parentorgids = ""

           do while not oOrgsAssigned.eof
              if lcl_parentorgids = "" then
                 lcl_parentorgids = oOrgsAssigned("orgid")
              else
                 lcl_parentorgids = lcl_parentorgids & ", " & oOrgsAssigned("orgid")
              end if

              oOrgsAssigned.movenext
           loop

           oOrgsAssigned.close
           set oOrgsAssigned = nothing
        end if
     end if

    'Get all of the orgs for the second feature if one has been selected.
    'Also, limit THIS list further by all of the orgs found in the first feature selection just done, if any exist.
     if iParentFeatureID2 <> "" then
        sSQL = " SELECT distinct orgid "
        sSQL = sSQL & " FROM egov_organizations_to_features "
        sSQL = sSQL & " WHERE featureid = " & iParentFeatureID2

        if lcl_parentorgids <> "" then
           sSQL = sSQL & " AND orgid IN (" & lcl_parentorgids & ") "
        end if

        sSQL = sSQL & " ORDER BY orgid "

        set oOrgsAssigned = Server.CreateObject("ADODB.Recordset")
        oOrgsAssigned.Open sSQL, Application("DSN"), 3, 1

        if not oOrgsAssigned.eof then
           lcl_parentorgids = ""

           do while not oOrgsAssigned.eof
              if lcl_parentorgids = "" then
                 lcl_parentorgids = oOrgsAssigned("orgid")
              else
                 lcl_parentorgids = lcl_parentorgids & ", " & oOrgsAssigned("orgid")
              end if

              oOrgsAssigned.movenext
           loop

           oOrgsAssigned.close
           set oOrgsAssigned = nothing
        end if
     end if
  end if

 'Now get all of the orgs that are:
 '1. IN the parent org id list
 '2. Do NOT have the new feature assigned
  sSQL = "SELECT distinct orgid "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid " & lcl_query_limitation & " IN (SELECT distinct orgid "
  sSQL = sSQL &                       "FROM egov_organizations_to_features "
  sSQL = sSQL &                       "WHERE featureid = " & iNewFeatureID & ") "

  if lcl_parentorgids <> "" then
     sSQL = sSQL & " AND orgid IN (" & lcl_parentorgids & ") "
  end if

  sSQL = sSQL & " ORDER BY orgid "

  set oOrgsNOTAssigned = Server.CreateObject("ADODB.Recordset")
  oOrgsNOTAssigned.Open sSQL, Application("DSN"), 3, 1

  if not oOrgsNOTAssigned.eof then
     do while not oOrgsNOTAssigned.eof

        if lcl_neworgids = "" then
           lcl_neworgids = oOrgsNOTAssigned("orgid")
        else
           lcl_neworgids = lcl_neworgids & ", " & oOrgsNOTAssigned("orgid")
        end if

        oOrgsNOTAssigned.movenext
     loop
  end if

  oOrgsNOTAssigned.close
  set oOrgsNOTAssigned = nothing

  if lcl_neworgids <> "" then
     lcl_return = lcl_neworgids
  else
     lcl_return = lcl_parentorgids
  end if

  getOrgIDs = lcl_return

end function

'------------------------------------------------------------------------------
sub assignUsers(iOrgID, iParentFeatureID1, iParentFeatureID2, iFeatureID, iIsRootAdmin)

 '1. Get all of the root admin users for the org
 '2. That do NOT have the feature already assigned
  if lcl_assignToOrgs = "Y" then
     sSQL = "SELECT userid, (firstname + ' ' + lastname) AS username  "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE orgid = " & iOrgID

     if iIsRootAdmin = "Y" then
        sSQL = sSQL & " AND isRootAdmin = 1 "
     end if

     sSQL = sSQL & " AND userid NOT IN (select distinct userid "
     sSQL = sSQL &                    " from egov_users_to_features "
     sSQL = sSQL &                    " where featureid = " & iFeatureID & ") "
     sSQL = sSQL & " ORDER BY userid "
  else
     lcl_parentfeatureids1  = ""
     lcl_parentfeatureids2  = ""
     lcl_assignable_userids = ""

     sSQL = "SELECT userid, username "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE orgid = " & iOrgID

     if iParentFeatureID1 <> "" AND iParentFeatureID2 <> "" then
        sSQL = sSQL & " AND userid IN (select userid "
        sSQL = sSQL &                " from egov_users_to_features "
        sSQL = sSQL &                " where featureid = " & iParentFeatureID2
        sSQL = sSQL &                " and userid IN (select userid "
        sSQL = sSQL &                               " from egov_users_to_features "
        sSQL = sSQL &                               " where featureid IN (" & iParentFeatureID1 & "))) "
     else
        if iParentFeatureID1 <> "" then
           sSQL = sSQL & " AND userid IN (select userid "
           sSQL = sSQL &                " from egov_users_to_features "
           sSQL = sSQL &                " where featureid = " & iParentFeatureID1 & ") "
        end if
     end if

     sSQL = sSQL & " AND userid NOT IN (select userid "
     sSQL = sSQL &                    " from egov_users_to_features "
     sSQL = sSQL &                    " where featureid = " & iFeatureID & ") "
dtb_debug(sSQL)
     set oGetUsers = Server.CreateObject("ADODB.Recordset")
     oGetUsers.Open sSQL, Application("DSN"), 3, 1

     if not oGetUsers.eof then
        do while not oGetUsers.eof
           if lcl_assignable_userids = "" then
              lcl_assignable_userids = oGetUsers("userid")
           else
              lcl_assignable_userids = lcl_assignable_userids & ", " & oGetUsers("userid")
           end if

           oGetUsers.movenext
        loop

        oGetUsers.close
        set oGetUsers = nothing
     end if     
  end if

  set oGetRootUsers = Server.CreateObject("ADODB.Recordset")
  oGetRootUsers.Open sSQL, Application("DSN"), 3, 1

  if not oGetRootUsers.eof then
     do while not oGetRootUsers.eof

        lcl_total_admins = lcl_total_admins + 1

        sSQL = "INSERT INTO egov_users_to_features (userid, featureid, permissionid, permissionlevelid) VALUES ("
        sSQL = sSQL & oGetRootUsers("userid") & ", "
        sSQL = sSQL & iFeatureID              & ", "
        sSQL = sSQL & "1, NULL"               & ") "

        set oAssignRootUser = Server.CreateObject("ADODB.Recordset")
        oAssignRootUser.Open sSQL, Application("DSN"), 3, 1

        set oAssignRootUser = nothing

       'Display the row for the root admin (userid) just inserted.
        response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "      <td>&nbsp;</td>" & vbcrlf
        response.write "      <td>" & trim(oGetRootUsers("username")) & " [" & oGetRootUsers("userid")& "]</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oGetRootUsers.movenext
     loop
  end if

  oGetRootUsers.close
  set oGetRootUsers = nothing

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
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "ASSIGN_SUCCESS" then
        lcl_return = "Feature Successfully Assigned..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     elseif iSuccess = "NC_ERROR" then
        lcl_return = "ERROR: Already assigned to all organizations..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>

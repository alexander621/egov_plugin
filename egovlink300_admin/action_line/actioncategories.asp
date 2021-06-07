<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: actioncategories.asp
' AUTHOR: John Stullenberger
' CREATED: 03/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  The code performs all actions on the form categories.
'
' MODIFICATION HISTORY
' 1.0 03/24/2006 John Stullenberger - Initial Version
' 1.1 10/18/2007	Steve Loar - Added code to sequence the categories after a delete
' 2.0 03/23/2012 David Boyer - Modified code to bring up to current standard and to work properly with non-IE browsers
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "form categories") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 dim oCmd, oRst, dDate, iDuration, sLinks, bShown

'Check for screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Action Line: Form Categories}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<style>
#screenMsg {
   color:       #ff0000;
   font-size:   10pt;
   font-weight: bold;
}

.categoryDiv {
   padding-bottom: 5px;
}

.cellNoWrap {
   white-space: nowrap;
}
</style>

  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
<!--
  $(document).ready(function(){
    $('#addButton').click(function() {
      if($('#newCategoryName').val() == '') {
         $('#newCategoryName').focus();
         inlineMsg(document.getElementById('addButton').id,'<strong>Required field missing: </strong> Category Name.',10,'addButton');
         return false;
      } else {
         $('#action').val('ADD');
         $('#formCategories').submit();
      }
    });

    $('input[name*="editButton"]').click(function() {
      $('#action').val('EDIT');
      $('#formCategories').submit();
    });
  });

function moveCategory(iDirection,iCategoryID) {
  var lcl_direction  = 'UP';
  var lcl_categoryid = '';
  var lcl_url        = '';

  if(iDirection != '' || iDirection != undefined) {
     lcl_direction = iDirection.toUpperCase();
  }

  if(iCategoryID != '' || iCategoryID != undefined) {
     lcl_categoryid = iCategoryID;
  }

  if(lcl_categoryid != '') {
     lcl_url  = 'order_categories.asp';
     lcl_url += '?direction=' + lcl_direction;
     lcl_url += '&iCatId='    + lcl_categoryid;
     lcl_url += '&iorgid=<%=session("orgid")%>';

     location.href = lcl_url;
  }
}

function deleteCatgory(iCategoryID) {
  var lcl_categoryid = '';
  var lcl_url        = '';

  if(iCategoryID != '' || iCategoryID != undefined) {
     lcl_categoryid = iCategoryID;
  }

  if(lcl_categoryid != '') {
     var lcl_categoryname = $('#editcategory' + lcl_categoryid).val();

    	if (confirm('Are you sure you want to delete this category: "' + lcl_categoryname + '"')) { 
        lcl_url  = 'actioncategories_action.asp';
        lcl_url += '?action=DELETE';
        lcl_url += '&categoryid=' + lcl_categoryid;
        lcl_url += '&orgid=<%=session("orgid")%>';

        location.href = lcl_url;
   		}
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
<body onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""formCategories"" id=""formCategories"" method=""post"" action=""actioncategories_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""5"" maxlength=""20"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""action"" id=""action"" value="""" size=""5"" maxlength=""20"" />" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Action Line Form Categories</strong></font></td>" & vbcrlf
  response.write "      <td><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf

 'BEGIN: New Category ---------------------------------------------------------
  response.write "		        <div class=""shadow"">" & vbcrlf
  response.write "            <table width=""100%"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <th align=""left"">Create a Form Category</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      Category Name: " & vbcrlf
  response.write "                      <input type=""text"" name=""newCategoryName"" id=""newCategoryName"" value="""" size=""80"" maxlength=""50"" onchange=""clearMsg('addButton')"" />" & vbcrlf
  response.write "                      <input type=""button"" name=""addButton"" id=""addButton"" value=""Add Category"" class=""button"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "		        </div>" & vbcrlf
 'END: New Category -----------------------------------------------------------

  response.write "<p>&nbsp;</p>" & vbcrlf

 'BEGIN: Edit Category --------------------------------------------------------
  response.write "          <div class=""categoryDiv"""">" & vbcrlf
  response.write "            <input type=""button"" name=""editButton"" id=""editButton"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "          <table width=""100%"" cellpadding=""4"" cellspacing=""0"" border=""0"" class=""tableadmin"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""left"">Modify a Form Category</th>" & vbcrlf
  response.write "                <th>&nbsp;</th>" & vbcrlf
  response.write "                <th>Display<br />Order</th>" & vbcrlf
  response.write "                <th>&nbsp;</th>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  sSQL = "SELECT form_category_id, "
  sSQL = sSQL & " form_category_name, "
  sSQL = sSQL & " form_category_sequence, "
  sSQL = sSQL & " orgid "
  sSQL = sSQL & " FROM egov_form_categories "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY form_category_sequence "

 	set oGetFormCategories = Server.CreateObject("ADODB.Recordset")
	 oGetFormCategories.Open sSQL, Application("DSN"), 3, 1

  lcl_bgcolor = "#eeeeee"
				  
  do while not oGetFormCategories.eof
     lcl_bgcolor                 = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     lcl_categoryExistsOnRequest = isCategoryUsedOnRequest(session("orgid"), oGetFormCategories("form_category_id"))

     response.write "            <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <input type=""text"" name=""editcategory" & oGetFormCategories("form_category_id") & """ id=""editcategory" & oGetFormCategories("form_category_id") & """ value=""" & oGetFormCategories("form_category_name") & """ size=""80"" maxlength=""50"" />" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "                <td>" & vbcrlf

     if lcl_categoryExistsOnRequest then
        response.write "&nbsp;"
     else
        response.write "                    <input type=""button"" name=""deleteButton" & oGetFormCategories("form_category_id") & """ id=""deleteButton" & oGetFormCategories("form_category_id") & """ value=""Delete"" class=""button"" onclick=""deleteCatgory('" & oGetFormCategories("form_category_id") & "');"" />" & vbcrlf
     end if

     response.write "                </td>" & vbcrlf
     response.write "                <td align=""center"">(" & oGetFormCategories("form_category_sequence") & ")</td>" & vbcrlf
     response.write "                <td class=""cellNoWrap"" width=""100%"">" & vbcrlf
     response.write "                    <input type=""button"" name=""moveUpButton"   & oGetFormCategories("form_category_id") & """ id=""moveUpButton"   & oGetFormCategories("form_category_id") & """ value=""Move Up"" class=""button"" onclick=""moveCategory('UP','" & oGetFormCategories("form_category_id") & "');"" />" & vbcrlf
     response.write "                    <input type=""button"" name=""moveDownButton" & oGetFormCategories("form_category_id") & """ id=""moveDownButton" & oGetFormCategories("form_category_id") & """ value=""Move Down"" class=""button"" onclick=""moveCategory('DOWN','" & oGetFormCategories("form_category_id") & "');"" />" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf

			  oGetFormCategories.movenext
		loop

  oGetFormCategories.close
  set oGetFormCategories = nothing

  response.write "          </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <div class=""categoryDiv"""">" & vbcrlf
  response.write "            <input type=""button"" name=""editButton2"" id=""editButton2"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
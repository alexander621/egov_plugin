<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="faq_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: faq_categories.asp
' AUTHOR: Steve Loar
' CREATED: 09/11/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module manages FAQ categories
'
' MODIFICATION HISTORY
' 1.0 09/11/06 Steve Loar - Original code
' 1.2 10/09/06 Steve Loar - Security, Header and Nav changed
' 1.3 07/21/09 David Boyer - Modified "move up" and "move down" to no longer call faq_category_move.asp.
'                          * This change fixes the bug with the ADO when trying to update the recordset (displayorder) in the loop.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iRowCount, oFAQCats, iMaxOrder, iInternalOnly

 sLevel             = "../"  'Override of value from common.asp
 lcl_faqtype        = "FAQ"
 lcl_userpermission = "manage faq"
 lcl_orgfeature     = "internal faq"
 lcl_pagetitle      = "FAQ"
 iMaxOrder          = GetMaxDisplayOrder()
 lcl_success        = request("success")
 lcl_onload         = ""

'Check for the faqtype
 if request("faqtype") <> "" then
    lcl_faqtype = UCASE(request("faqtype"))
 end if

'Based on the faqtype check for the proper permission
 if lcl_faqtype = "RUMORMILL" then
    lcl_userpermission = "rumormill_manage"
    lcl_orgfeature     = "rumormill_internal"
    lcl_pagetitle      = "Rumor Mill"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for a screen message
 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check to see if the Save Changes button should be disabled.
 if iMaxOrder < 1 then
    lcl_savechanges_top    = "document.getElementById('button_save_TOP').disabled=true;"
    lcl_savechanges_bottom = "document.getElementById('button_save_BOTTOM').disabled=true;"

    if lcl_onload <> "" then
       lcl_onload = lcl_onload & lcl_savechanges_top
       lcl_onload = lcl_onload & lcl_savechanges_bottom
    else
       lcl_onload = lcl_savechanges_top
       lcl_onload = lcl_onload & lcl_savechanges_bottom
    end if
 end if

'Check for org features
 lcl_orghasfeature_internal_faq = orghasfeature(lcl_orgfeature)
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%> Categories}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <link rel="stylesheet" type="text/css" href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css" />

<style type="text/css">
  #screenMsg {
     text-align:  right;
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
  }

  .placeHolderHighlight {
     background-color: #ff0000;
     height:           1.5em;
     line-height:       1.2em;
  }

  .dragDropArrows {
     cursor: pointer;
  }
</style>

  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>
  <script type="text/javascript" src="../scripts/jquery-ui-1.8.4.custom.min.js"></script>

<script type="text/javascript">
<!--
$(document).ready(function(){

  //Return a helper with preserved width of cells
  var fixHelper = function(e, ui) {
    	ui.children().each(function() {
		      $(this).width($(this).width());
  	  });

  	  return ui;
  };

  $('#tablefaqcatSAVE tbody').sortable({
     connectWith: '#tablefaqcatSAVE tbody',
     helper:      fixHelper,
     revert:      true,
     cursor:      'move',
     stop: function(event, ui) {
        $el = $(ui.item);
        $el.find('tr').click();
        $el.effect('highlight',{},2000);

        //update the Display Order
	      	$('#tablefaqcatSAVE tbody').each(function(){
          var itemorder       = $(this).sortable('toArray');
          var lcl_total_items = itemorder.length;
          var lcl_rowID       = '';

          for(var i = 0; i < lcl_total_items; i++) {
             lcl_rowID = itemorder[i];
             lcl_rowID = lcl_rowID.replace('addFieldRow','');

             $('#displayorder' + lcl_rowID).val(i+1);
          }
        });
     }
  })
  .disableSelection();

});

function SaveCategory(p_action) {
  if(p_action=="ADD") {
     lcl_form             = document.getElementById("category_add");
     lcl_total_categories = 0;
     lcl_i_start          = 0;
  }else{
    lcl_form             = document.getElementById("category_save");
    lcl_total_categories = document.getElementById("total_faq_categories").value;
    lcl_i_start          = 1;
  }

  lcl_false_cnt = 0;

  for (i=lcl_i_start; i<=lcl_total_categories; ++ i) {
     if(document.getElementById("FAQCategoryName"+i).value == "") {
        inlineMsg(document.getElementById("FAQCategoryName"+i).id,'<strong>Required Field Missing: </strong> Category Name',10,'FAQCategoryName'+i);
        lcl_false_cnt = lcl_false_cnt + 1;

        if(lcl_false_cnt == 1) {
           lcl_focus = document.getElementById("FAQCategoryName"+i);
        }
     }else{
        clearMsg('FAQCategoryName'+i);
     }
  }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
  if(lcl_false_cnt > 0) {
     lcl_focus.focus();
     return false;
  }else{
     lcl_focus_cnt = 0;
  }

  lcl_form.submit();
}

function ConfirmDelete(iRowID,iFAQCategoryId) {
  lcl_categoryname = document.getElementById("FAQCategoryName"+iRowID).value;

  var msg = "Do you wish to delete the category '" + lcl_categoryname + "'?"
		if (confirm(msg)) {
   			location.href='faq_category_delete.asp?faqtype=<%=lcl_faqtype%>&FAQCategoryId=' + iFAQCategoryId;
		}
}

//function ChangeOrder( iDisplayOrder,iDirection ) {
//  location.href='faq_category_move.asp?faqtype=<%'lcl_faqtype%>&iDisplayOrder='+ iDisplayOrder + '&iDirection=' + iDirection;
//}

//function ChangeOrder(p_ID, p_Direction) {
  //1. Determine which way we are moving this record.
  //2. Get the total number of rows.
  //2. Get the display order of the current row and the one we are moving to.
  //3. Swap the display order values
  //4. Submit (save) all of the rows.
//  if(p_Direction == "UP") {
//     iDirection = -1;
//  }else{
//     iDirection = 1;
//  }

//  iSwapRowID = p_ID + iDirection;

//  iTotalRows           = document.getElementById("total_faq_categories").value;
//  iCurrentDisplayOrder = document.getElementById("displayorder" + p_ID).value;
//  iNewDisplayOrder     = document.getElementById("displayorder" + iSwapRowID).value;

  //Swap the displayorder values for the 2 rows.
//  document.getElementById("displayorder" + iSwapRowID).value = iCurrentDisplayOrder;
//  document.getElementById("displayorder" + p_ID).value       = iNewDisplayOrder;

  //Save the changes
//  document.getElementById("category_save").submit();
//}

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
  response.write "<div id=""content"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width:550px"">" & vbcrlf
  response.write "  <tr valign=""bottom"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_pagetitle & ": Category Management</strong><br /></font>" & vbcrlf
  response.write "          <input type=""button"" name=""returnButton"" id=""returnButton"" class=""button"" value=""<< Back to " & lcl_pagetitle & "s"" onclick=""location.href='list_faq.asp?faqtype=" & lcl_faqtype & "';"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td id=""screenMsg""></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

 'Setup up variables
  lcl_rowcount        = 0
  lcl_bgcolor         = "#ffffff"
  lcl_faqcategoryid   = 0
  lcl_displayorder    = iMaxOrder+1
  lcl_faqcategoryname = ""
  lcl_internalonly    = False

 'BEGIN: Add FAQ Category -----------------------------------------------------

  displayCategoryRowHeaders "category_add"

  displayCategoryRow lcl_rowcount, _
                     lcl_bgcolor, _
                     lcl_faqtype, _
                     lcl_faqcategoryid, _
                     lcl_displayorder, _
                     lcl_faqcategoryname, _
                     lcl_internalonly

 'Close the FORM, TABLE, and DIV tags as they are opened in "displayRateRowHeaders"
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "</form>" & vbcrlf
 'END: Add FAQ Category -------------------------------------------------------

 'BEGIN: Edit FAQ Category ----------------------------------------------------
  displayCategoryRowHeaders "category_save"

	'Get the rows of existing faq categories
  sSQL = "SELECT FAQCategoryId, FAQCategoryName, displayorder, internalonly "
  sSQL = sSQL & " FROM faq_categories "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND UPPER(faqtype) = '" & lcl_faqtype & "' "
  sSQL = sSQL & " ORDER BY displayorder"

		set oFAQCats = Server.CreateObject("ADODB.Recordset")
		oFAQCats.Open sSQL, Application("DSN"), 3, 1
		
		if not oFAQCats.eof then

     response.write "<tbody>" & vbcrlf

   		do while not oFAQCats.eof
        lcl_rowcount = lcl_rowcount + 1
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

        displayCategoryRow lcl_rowcount, _
                           lcl_bgcolor, _
                           lcl_faqtype, _
                           oFAQCats("FAQCategoryID"), _
                           oFAQCats("displayorder"), _
                           oFAQCats("FAQCategoryName"), _
                           oFAQCats("internalonly")

    				oFAQCats.movenext
     loop

     response.write "</tbody>" & vbcrlf

  end if

		oFAQCats.close
		set oFAQCats = nothing

  response.write "  <tr><td colspan=""100""><input type=""hidden"" name=""total_faq_categories"" id=""total_faq_categories"" value=""" & lcl_rowcount & """ size=""5"" /></td></tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

  displayButtons "SAVE","BOTTOM"

  response.write "</p>" & vbcrlf
  response.write "</form>" & vbcrlf
 'END: Edit FAQ Category ------------------------------------------------------

  response.write "</div>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
function GetMaxDisplayOrder()
	 Dim sSql, oMax
  lcl_return = 0

 	sSQL = "SELECT MAX(displayorder) as MaxOrder "
  sSQL = sSQL & " FROM faq_categories "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND UPPER(faqtype) = '" & lcl_faqtype & "'"

	 set oMax = Server.CreateObject("ADODB.Recordset")
  oMax.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

  if isnull(oMax("MaxOrder")) then
   		lcl_return = 0
  else
   		lcl_return = oMax("MaxOrder")
  end if

  oMax.close
  set oMax = nothing

  getMaxDisplayOrder = lcl_return

end function

'------------------------------------------------------------------------------
sub displayCategoryRowHeaders(iFormName)

  if UCASE(iFormName) = "CATEGORY_ADD" then
     lcl_tableid = "category_add"
     lcl_button  = "ADD"
  else
     lcl_tableid = "category_save"
     lcl_button  = "SAVE"
  end if

 'Setup the add/update form
  response.write "<form name=""" & iFormName & """ id=""" & iFormName & """ method=""post"" action=""faq_categories_save.asp"">" & vbcrlf
  response.write "<p>" & vbcrlf

 'Display ADD/SAVE button
  if lcl_button = "SAVE" then
     displayButtons lcl_button,"TOP"
  end if

 'Open Table and display Column Headers
  response.write "<div id=""faqshadow"">" & vbcrlf
  response.write "<table cellpadding=""5"" cellspacing=""0"" border=""0"" id=""tablefaqcat" & lcl_button & """>" & vbcrlf
  response.write "  <thead>" & vbcrlf
  response.write "  <tr align=""center"">" & vbcrlf
  response.write "      <th align=""left"">Category Name</th>" & vbcrlf

  if lcl_orghasfeature_internal_faq then
     response.write "      <th>Internal<br />Only</th>" & vbcrlf
  end if

  response.write "      <th colspan=""2"">&nbsp;</th>" & vbcrlf
  response.write "		</tr>" & vbcrlf
  response.write "  </thead>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub displayCategoryRow(iRowCount, iBGColor, iFAQType, iFAQCategoryID, iDisplayOrder, iFAQCategoryName, iInternalOnly)

 'Check if we need to show the Move Down button
  'if iDisplayOrder > 1 AND iDisplayOrder <= iMaxOrder then
     'lcl_move_up = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move UP');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder(" & iDisplayOrder & ", -1);"" /><br />" & vbcrlf
  '   lcl_move_up = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move UP');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder(" & iRowCount & ",'UP');"" /><br />" & vbcrlf
  'else
  '   lcl_move_up = ""
  'end if

 'Check if we need to show the Move Up button
  'if iDisplayOrder < iMaxOrder then
     'lcl_move_down = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move DOWN');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder(" & iDisplayOrder & ", 1);"" />" & vbcrlf
  '   lcl_move_down = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move DOWN');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder(" & iRowCount & ",'DOWN');"" />" & vbcrlf
  'else
  '   lcl_move_down = ""
  'end if

 'If neither button is shown we need to put in a "space" so that the border works properly if this record just happens to be
 'the last record for the category.
  'if lcl_move_down = "" AND lcl_move_up = "" then
  '   lcl_move_down = "&nbsp;"
  'end if

  response.write "  <tr id=""addFieldRow" & iRowCount & """ align=""center"" bgcolor=""" & iBGColor & """>" & vbcrlf
  response.write "      <td align=""left"">" & vbcrlf
  response.write "          <input type=""hidden"" name=""orgid"         & iRowCount & """ id=""orgid"           & iRowCount & """ value=""" & session("orgid") & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""faqtype"       & iRowCount & """ id=""faqtype"         & iRowCount & """ value=""" & iFAQType         & """ />" & vbcrlf
  response.write "          <input type=""hidden"" name=""FAQCategoryId" & iRowCount & """ id=""FAQCategoryId"   & iRowCount & """ value=""" & iFAQCategoryID   & """ size=""5"" />" & vbcrlf
  response.write "          <input type=""hidden"" name=""displayorder"  & iRowCount & """ id=""displayorder"    & iRowCount & """ value=""" & iDisplayOrder    & """ size=""5"" />" & vbcrlf

  if iRowCount > 0 then
     response.write "          <img src=""arrow.png"" width=""12"" height=""12"" title=""click to drag and reorder"" class=""dragDropArrows"" />" & vbcrlf
  end if

  response.write "          <input type=""text"" name=""FAQCategoryName" & iRowCount & """ id=""FAQCategoryName" & iRowCount & """ value=""" & iFAQCategoryName & """ size=""50"" maxlength=""50"" onchange=""clearMsg('FAQCategoryName" & iRowCount & "')"" />" & vbcrlf
  response.write "      </td>" & vbcrlf

  if lcl_orghasfeature_internal_faq then
     if iInternalOnly then
        lcl_checked_internalonly = " checked=""checked"""
     else
        lcl_checked_internalonly = ""
     end if

     response.write "      <td>" & vbcrlf
     response.write "          <input type=""checkbox"" name=""internalonly" & iRowCount & """ id=""internalonly" & iRowCount & """" & lcl_checked_internalonly & " value=""on"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
  end if

  'response.write "      <td nowrap=""nowrap"">" & lcl_move_up & lcl_move_down & "</td>" & vbcrlf
  response.write "      <td class=""action"" nowrap=""nowrap"">" & vbcrlf

  if iRowCount > 0 then
     response.write "          <input type=""button"" name=""deleteCategory" & iRowCount & """ id=""deleteCategory" & iRowCount & """ value=""Delete"" class=""button"" onclick=""ConfirmDelete(" & iRowCount & "," & iFAQCategoryID & ");"" />" & vbcrlf
  else
     displayButtons "ADD","TOP"
  end if

  response.write "      </td>" & vbcrlf


  response.write "  </tr>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayButtons(iType,iTopBottom)
  lcl_buttonType = "SAVE"
  lcl_topBottom  = "TOP"

  if iTopBottom <> "" then
     lcl_topBottom = UCASE(iTopBottom)
  end if

  if iTopBottom = "BOTTOM" then
     lcl_padding = "padding-top: 5px"
  else
     lcl_padding = "padding-bottom: 5px"
  end if

  if iType <> "" then
     lcl_buttonType = UCASE(iType)
  end if

  if iType = "ADD" then
     lcl_buttonid    = "button_add"
     lcl_buttonvalue = "Add " & lcl_pagetitle & " Category"
  else
     lcl_buttonid    = "button_save_" & lcl_topBottom
     lcl_buttonvalue = "Save Changes"
  end if

  response.write "<div style=""" & lcl_padding & """>" & vbcrlf
  response.write "  <input type=""button"" name=""sAction"" id=""" & lcl_buttonid & """ value=""" & lcl_buttonvalue & """ class=""button"" onclick=""SaveCategory('" & iType & "');"" />" & vbcrlf
  response.write "</div>" & vbcrlf

end sub
%>
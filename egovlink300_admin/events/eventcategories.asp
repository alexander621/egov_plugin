<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: eventcategories.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the calendar
'
' MODIFICATION HISTORY
' 1.0 ???			     ???? - INITIAL VERSION
' 1.1	10/11/2006	Steve Loar - Security, Header and nav changed
' 1.2	08/10/2007	Steve Loar - Red changed from #CC0000 to #FF0000
' 1.3 08/07/2008 David Boyer - Add Custom Calendar
' 1.4 01/27/2012 David Boyer - Fix bugs in saving so screen works in Firefox.
'                              Also updated the code to our current design standards.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("calendar") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 Dim oCmd, oRst, dDate, iDuration, sTimeZones, sLinks, bShown

 lcl_sessionOrgIDisZero   = clng(session("orgId")) = clng(0)
 lcl_calendarfeatureid    = ""
 lcl_calendarfeature      = ""
 lcl_calendarfeature_url  = ""
 lcl_calendarfeature_name = ""

'Allow the user to maintain the Event Categories if any/all of the following:
 '1. The user has the "categories" permission assigned
 '2. The user has a specific Custom Calendar feature assigned: [session("calendarfeature") <> ""]
 if trim(request("cal")) <> "" then
    if not isnumeric(trim(request("cal"))) then
      	response.redirect sLevel & "permissiondenied.asp"
    else
       lcl_calendarfeatureid = CLng(trim(request("cal")))
       lcl_calendarfeature   = getFeatureByID(session("orgid"), lcl_calendarfeatureid)

       if OrgHasFeature(lcl_calendarfeature) AND UserHasPermission(session("userid"), lcl_calendarfeature) then
          lcl_calendarfeature_url  = "?cal=" & lcl_calendarfeatureid
          lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
       else
         	response.redirect sLevel & "permissiondenied.asp"
       end if
    end if
 else
    if NOT userhaspermission( session("userid"), "categories" ) then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

'New Category ---------------------------------------------------
 if request.form("_task") = "newcategory" then
   'Create a New Category for this Organization
    'newCategory session("orgid"), request.form("CustomCategory"), request.form("Color"), session("calendarfeature"), lcl_identity
    newCategory session("orgid"), request.form("CustomCategory"), request.form("Color"), lcl_calendarfeature, lcl_identity

    iCategoryID = lcl_identity

    response.redirect "../events/eventcategories.asp?success=SA&cal=" & lcl_calendarfeatureid

'Edit Categories ------------------------------------------------
 'elseif request.form("_task") = "editcategories" then

 '   for each Item IN request.form
 '      if LEFT(Item,6) = "Custom" then
 '         if LEFT(Item,15) = "CustomCategory_" then
 '            if Request.Form(Item) <> "" then
 '         	  			sUnderscore = instr(Item,"_")
 '            			sCatId      = Right(Item, Len(Item)-sUnderscore)

 '         	  			sSQLa = "UPDATE EventCategories SET CategoryName = '" & Request.Form(Item) & "' WHERE CategoryID = " & sCatId

 '               Set oUpdate = Server.CreateObject("ADODB.Recordset")
 '           				oUpdate.Open sSQLa, Application("DSN"), 3, 1
 '          	 			Set oUpdate = Nothing
 '      	  		 end if
 '         elseif LEFT(Item,12) = "CustomColor_" then
 '            if request.form(Item) <> "" then
 '               sUnderscore = instr(Item,"_")
 '               sCatId      = RIGHT(Item, Len(Item)-sUnderscore)

 '               sSQLa =  "UPDATE EventCategories SET Color = '" & Request.Form(Item) & "' WHERE CategoryID = " & sCatId

 '           				Set oUpdate = Server.CreateObject("ADODB.Recordset")
 '           				oUpdate.Open sSQLa, Application("DSN"), 3, 1
 '           				Set oUpdate = Nothing
 '            end if
 '         end if
 '      end if
 '   next

 '   response.redirect "../events/eventcategories.asp?success=SU&cal=" & lcl_calendarfeatureid

'Delete Category ----------------------------------------------
 elseif request.form("_task") = "deletecategory" then
    for each Item IN request.form
        if LEFT(UCASE(Item),11) = "CATEGORYID_" then
           if request.form(Item) <> "" then
              'lcl_categoryid = replace(request.form(UCASE(Item)),"CATEGORYID_","")
              lcl_categoryid = request.form(Item)

              delCategory(lcl_categoryid)
     	  		 end if
        end if
    next

    response.redirect "../events/eventcategories.asp?success=SD&cal=" & lcl_calendarfeatureid

'Edit Category (orgid = 0) ----------------------------------
 elseif request.form("_task") = "modifycategory" then

    sCategoryName = getCategoryName(request.form("category"))

    if NOT Request.Form("CustomCategory") = "" then
    	  sCategoryName = Request.Form("CustomCategory")
    end if

   'Update the Event Category
    'updateCategory request("Category"), sCategoryName, request("Color"), session("calendarfeature"), lcl_identity
    updateCategory request("Category"), sCategoryName, request("Color"), lcl_calendarfeature, lcl_identity

    response.redirect "../events/eventcategories.asp?success=SU&cal=" & lcl_calendarfeatureid

 else

    Set oCmd = Server.CreateObject("ADODB.Command")
    With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText      = "ListTimeZones"
      .CommandType      = adCmdStoredProc
      .Execute
    End With

    Set oRst = Server.CreateObject("ADODB.Recordset")
    With oRst
      .CursorLocation = adUseClient
      .CursorType     = adOpenStatic
      .LockType       = adLockReadOnly
      .Open oCmd
    End With
    Set oCmd = Nothing

    while NOT oRst.eof
       sTimeZones = sTimeZones & "<option "

       if oRst("TimeZoneID") = 1 then
          sTimeZones = sTimeZones & " selected=""selected"" "
       end if

       sTimeZones=sTimeZones & " value=""" & oRst("TimeZoneID") & """>" & oRst("TZName") & "</option>" & vbcrlf

       oRst.movenext
    wend

    if oRst.State=1 then oRst.Close
    set oRst = Nothing 

 end if

'Get the number of items displayed per day
 lcl_itemsPerDay = getItemsPerDay(session("orgid"))

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
  <title><%=langBSEVents%><%=lcl_calendar_name%></title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="javascript" src="../scripts/selectAll.js"></script>
	<script language="javascript" src="../scripts/ajaxLib.js"></script>
 <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

 <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

	<script language="javascript">
	<!--

		function storeCaret (textEl) 
		{
			if (textEl.createTextRange)
				textEl.caretPos = document.selection.createRange().duplicate();
		}

		function insertAtCaret (textEl, text) 
		{
			if (textEl.createTextRange && textEl.caretPos) 
			{
				var caretPos = textEl.caretPos;
				caretPos.text =
					caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				text + ' ' : text;
			}
			else
				textEl.value  = text;
		}

		function doPicker(sFormField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		//function fnCheckCategory()
		//{
		//	if (document.all.DeleteEventCategory.Category.value != '0') 
		//	{
		//		return true;
		//	}
		//	else
		//	{
		//		return false;
		//	}
		//}

function validateItemsPerDay() {
  var lcl_return_false  = 0;
  var lcl_itemsPerDay   = document.getElementById("itemsPerDay");
  var lcl_numberofitems = "4";


  if (lcl_itemsPerDay.value != "") {
      var rege = /^\d+$/;
      var Ok   = rege.exec(lcl_itemsPerDay.value);

 		   if ( ! Ok ) {
          if (lcl_itemsPerDay.value < 1) {
              lcl_message = "<strong>Invalid Value (must be greater than zero (0)): </strong>Items Per Day";
          } else {
              lcl_message = "<strong>Invalid Value (must be numeric): </strong>Items Per Day";
          }

          lcl_itemsPerDay.focus();
          inlineMsg(document.getElementById("saveItemsPerDay").id,lcl_message,10,'saveItemsPerDay');
          lcl_return_false = lcl_return_false + 1;
      } else {
        clearMsg("saveItemsPerDay");
        lcl_numberofitems = lcl_itemsPerDay.value;
      }
  } else {
      lcl_itemsPerDay.value = "4";
  }

  if (lcl_numberofitems < 1) {
      lcl_itemsPerDay.focus();
      inlineMsg(document.getElementById("saveItemsPerDay").id,'<strong>Invalid Value (must be greater than zero (0)): </strong>Items Per Day',10,'saveItemsPerDay');
      lcl_return_false = lcl_return_false + 1;
  }

  if(lcl_return_false > 0) {
     return false;
  }else{
   		var sParameter  = 'isAjaxRoutine=Y';
     sParameter     += '&orgid='       + encodeURIComponent('<%=session("orgid")%>');
     sParameter     += '&itemsPerDay=' + encodeURIComponent(lcl_numberofitems);

     doAjax('saveItemsPerDay.asp', sParameter, 'displayScreenMsg', 'post', '0');
  }
}

function returnToEvents() {
  location.href='default.asp<%=lcl_calendarfeature_url%>';
}

function createCategory() {
  var lcl_newEventCategoryForm = document.getElementById('newEventCategory');
  var lcl_new_category         = document.getElementById('newCustomCategory');

  if (lcl_new_category.value != '') {
      clearMsg('newCustomCategory');
      lcl_newEventCategoryForm.submit();
		} else {
      lcl_new_category.focus();
      inlineMsg(document.getElementById("newCustomCategory").id,'<strong>Required Field Missing: </strong>Category Name',10,'newCustomCategory');
		}
}

function saveChanges() {
  document.getElementById('editEventCategory').submit();
}

function deleteCategory(iLineNum) {
  var lcl_disable_category = true;
  var lcl_disable_color    = true;

  if(! document.getElementById('deleteCategory_' + iLineNum).checked) {
     lcl_disable_category = false;
     lcl_disable_color    = false;
  }

  $('#CustomCategory_' + iLineNum).prop('disabled',lcl_disable_category);
  $('#CustomColor_'    + iLineNum).prop('disabled',lcl_disable_color);
}

//function saveChanges2() {
//  var lcl_modifyEventCategoryForm = document.getElementById('modifyEventCategory');
//  if (((document.all.modifyEventCategory.Category.value != '0') || (document.modifyEventCategory.CustomCategory.value != ''))) {
//     document.all.modifyEventCategory.submit();
//  } else {
//     alert('Please enter a subject!');
//  }
//}

function displayHelpTip(p_status) {

  if(p_status == "SHOW") {
     inlineMsg(document.getElementById("itemsPerDayHelp").id,'<strong>NOTE: </strong>If left blank, the value will default to four (4) items per day.  This single value is used for ALL calendars/categories.',10,'itemsPerDayHelp');
  } else {
     clearMsg("itemsPerDayHelp");
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

<style type="text/css">
  #body {
     background-color: #ffffff;
     margin:           0px;
  }

  .unused  {
     font-family: Arial,Tahoma,Verdana;
     font-size:   13px;
     color:       #cc0000;
  }

  option.color0000ff { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#0000ff; }
  option.color006600 { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#006600; }
  option.colorCC0066 { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#CC0066; }
  option.colorff9900 { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#FF9900; }
  option.colorC76309 { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#C76309; }
  option.color9933cc { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#9933CC; }
  option.colorCC0000 { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#CC0000; }
  option.color0099ff { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#0099FF; }
  option.colorff33cc { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#FF33CC; }
  option.colorff0000 { font-family:Arial,Tahoma,Verdana; font-size:13px; color:#FF0000; }

  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
  }

  .buttonRowTop {
     font-size:      10px;
     padding-bottom: 10px;
  }

  .buttonRowBottom {
     font-size:   10px;
     padding-top: 10px;
  }

  .customCategory {
     width:133px;
  }

  .noWrap {
     white-space: nowrap;
  }

</style>

</head>
<body id="body" onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!-- #include file="../menu/menu.asp" //-->
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & langEvents & ": Event Categories" & lcl_calendar_name & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return"" class=""button"" onclick=""returnToEvents()"" />" & vbcrlf
  'response.write "		        <img src=""../images/arrow_2back.gif"" align=""absmiddle"" />&nbsp;<a href=""javascript:returnToEvents()"">" & langBackToEventList & "</a>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"" width=""35%"">" & vbcrlf
  response.write "          <span id=""screenMsg""></span><br />" & vbcrlf
  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "            <legend># Items Displayed per Day&nbsp;</legend>" & vbcrlf
  response.write "            Items Per Day: <input type=""text"" name=""itemsPerDay"" id=""itemsPerDay"" value=""" & lcl_itemsPerDay & """ size=""3"" maxlength=""3"" onchange=""clearMsg('saveItemsPerDay');"" />&nbsp;" & vbcrlf
  response.write "            <input type=""button"" name=""saveItemsPerDay"" id=""saveItemsPerDay"" value=""Save"" class=""button"" onclick=""validateItemsPerDay()"" />" & vbcrlf
  response.write "            <img src=""../images/help.jpg"" name=""itemsPerDayHelp"" id=""itemsPerDayHelp"" border=""0"" style=""cursor:pointer"" onmouseover=""displayHelpTip('SHOW');"" onmouseout=""displayHelpTip('HIDE');"" />" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf

 'BEGIN: New Category ---------------------------------------------------------
  response.write "          <form name=""newEventCategory"" id=""newEventCategory"" action=""eventcategories.asp"" method=""post"">" & vbcrlf
  response.write "            <input type=""hidden"" name=""_task"" id=""_task"" value=""newcategory"" />" & vbcrlf
  response.write "            <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ />" & vbcrlf

                            'displayButtons "CREATE", "TOP"
  response.write "		        <div class=""shadow"">" & vbcrlf
  response.write "            <table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"">" & vbcrlf

  'if lcl_sessionOrgIDisZero then
  '   response.write "              <tr>" & vbcrlf
  '   response.write "                  <th align=""left"" colspan=""2"">Create an Event Category</th>" & vbcrlf
  '   response.write "              </tr>" & vbcrlf
  '   response.write "	             <tr>" & vbcrlf
  '   response.write "	                 <td>Category:</td>" & vbcrlf
  '   response.write "                  <td><input type=""text"" name=""CustomCategory"" id=""CustomCategory"" class=""customCategory"" maxlength=""50"" /></td>" & vbcrlf
  '   response.write "	             </tr>" & vbcrlf
  '   response.write "	             <tr>" & vbcrlf
  '   response.write "	                 <td>Color Flag:</td>" & vbcrlf
  '   response.write "                  <td>" & vbcrlf
  '   response.write "	                    	<select name=""Color"" id=""Color"">" & vbcrlf
  '                                           displayColorOptions
  '   response.write "                      </select>" & vbcrlf
  '   response.write "			               </td>" & vbcrlf
  '   response.write "   	          </tr>" & vbcrlf
  'else
     response.write "       	      <tr>" & vbcrlf
     response.write "	                 <th align=""left"" colspan=""5"">Create Event Category</th>" & vbcrlf
     response.write "	             </tr>" & vbcrlf
     response.write "	             <tr>" & vbcrlf
     response.write "	                 <td class=""noWrap"">Category Name:</td>" & vbcrlf
     response.write "	                 <td><input type=""text"" name=""CustomCategory"" id=""newCustomCategory"" class=""customCategory"" maxlength=""50"" /></td>" & vbcrlf
     response.write "	                 <td class=""noWrap"">Category Color:</td>" & vbcrlf
     response.write "	                 <td>" & vbcrlf
     response.write "	            	        <select name=""Color"" id=""Color"">" & vbcrlf
                                             displayColorOptions
     response.write "                      </select>" & vbcrlf
     response.write "                  </td>" & vbcrlf
     response.write "	                 <td width=""70%""><input type=""button"" value=""Create Category"" class=""button"" onclick=""createCategory();"" /></td>" & vbcrlf
     response.write "   	          </tr>" & vbcrlf
  'end if

  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf

  'if lcl_sessionOrgIDisZero then
  '   displayButtons "CREATE", "BOTTOM"
  'end if

  response.write "          </form>" & vbcrlf
 'END: New Category -----------------------------------------------------------

 'BEGIN: Edit Category --------------------------------------------------------
  response.write "          <p>" & vbcrlf

  'if lcl_sessionOrgIDisZero then
  '   response.write "          <form name=""modifyEventCategory"" id=""modifyEventCategory"" action=""eventcategories.asp"" method=""post"">" & vbcrlf
  '   response.write "       	    <input type=""hidden"" name=""_task"" id=""_task"" value=""modifycategory"" />" & vbcrlf
  '   response.write "            <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ />" & vbcrlf

  '                             displayButtons "EDIT", "TOP"
  '   response.write "       			<div class=""shadow"">" & vbcrlf
  '   response.write "            <table width=""100%"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin"">" & vbcrlf
  '   response.write "              <tr>" & vbcrlf
  '   response.write "                  <th align=""left"" colspan=""2"">Modify an Event Category</th>" & vbcrlf
  '   response.write "              </tr>" & vbcrlf
  '   response.write "              <tr>" & vbcrlf
  '   response.write "                  <td valign=""top"">Category:</td>" & vbcrlf
  '   response.write "                  <td>Choose:" & vbcrlf
  '   response.write "                      <select name=""Category"" id=""Category"" class=""time"">" & vbcrlf
  '   response.write "                        <option value=""0"">None</option>" & vbcrlf
  '                                           displayCategoryOptions session("orgid"), lcl_calendarfeature
  '   response.write "                      </select>" & vbcrlf
  '   response.write "                      Change Category Name to:" & vbcrlf
  '   response.write "                      <input type=""text"" name=""CustomCategory"" id=""CustomCategory"" class=""customCategory"" maxlength=""50"" />" & vbcrlf
  '   response.write "                  </td>" & vbcrlf
  '   response.write "              </tr>" & vbcrlf
  '   response.write "              <tr>" & vbcrlf
  '   response.write "                  <td valign=""top"">Color Flag:</td>" & vbcrlf
  '   response.write "                  <td>Change Color to:" & vbcrlf
  '   response.write "              	       <select name=""Color"" id=""Color"">" & vbcrlf
  '                                           displayColorOptions
  '   response.write "                      </select>" & vbcrlf
  '   response.write "                  </td>" & vbcrlf
  '   response.write "                  <td>&nbsp;</td>" & vbcrlf
  '   response.write "              </tr>" & vbcrlf
  '   response.write "            </table>" & vbcrlf
  '   response.write "		        </div>" & vbcrlf

  '                             displayButtons "EDIT", "BOTTOM"

  '   response.write "          </form>" & vbcrlf
  'else
     response.write "          <form name=""editEventCategory"" id=""editEventCategory"" action=""eventcategories_action.asp"" method=""post"">" & vbcrlf
     response.write "            <input type=""hidden"" name=""_task"" id=""edit_task"" value=""editcategories"" />" & vbcrlf
     response.write "            <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ />" & vbcrlf

                               displayButtons "SAVE", "TOP"
     response.write "          <div class=""shadow"">" & vbcrlf
     response.write "          <table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"">" & vbcrlf
     response.write "            <tr>" & vbcrlf
     'response.write "                <th><input class=""listCheck"" type=""checkbox"" name=""chkSelectAll"" onclick=""selectAll('ModifyEventCategory2', this.checked)"" /></th>" & vbcrlf
     response.write "                <th align=""left"" colspan=""3"">Modify Event Category</th>" & vbcrlf
     response.write "                <th>Delete</th>" & vbcrlf
     response.write "            </tr>" & vbcrlf

                                 displayCategoryRows session("orgid"), lcl_calendarfeature

     response.write "          </table>" & vbcrlf
     response.write "          </div>" & vbcrlf
     response.write "          </form>" & vbcrlf
  'end if

  response.write "          </p>" & vbcrlf

 'BEGIN: Delete Category ------------------------------------------------------
  'if lcl_sessionOrgIDisZero then
  '   response.write "          <p>" & vbcrlf
  '   response.write "          <form name=""DeleteEventCategory"" id=""DeleteEventCategory"" action=""eventcategories.asp"" method=""post"">" & vbcrlf
  '   response.write "            <input type=""hidden"" name=""_task"" id=""_task"" value=""deletecategory"">" & vbcrlf
  '   response.write "            <input type=""hidden"" name=""cal"" value=""" & lcl_calendarfeatureid & """ />" & vbcrlf

  '                             displayButtons "DELETE", "TOP"
  '   response.write "      		  <div class=""shadow"">" & vbcrlf
  '   response.write "            <table width=""100%"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin"">" & vbcrlf

  '   if lcl_sessionOrgIDisZero then
  '      response.write "              <tr>" & vbcrlf
  '      response.write "                  <th align=""left"" colspan=""2"">Delete an Event Category</th>" & vbcrlf
  '      response.write "              </tr>" & vbcrlf
  '      response.write "              <tr>" & vbcrlf
  '      response.write "                  <td valign=""top"">Category:</td>" & vbcrlf
  '      response.write "                  <td>Choose:" & vbcrlf
  '      response.write "                      <select name=""Category"" id=""Category"" class=""time"">" & vbcrlf
  '      response.write "                        <option value=""0"">None</option>" & vbcrlf
  '                                              displayCategoryOptions session("orgid"), lcl_calendarfeature
  '      response.write "                      </select>" & vbcrlf
  '      response.write "                  </td>" & vbcrlf
  '      response.write "              </tr>" & vbcrlf
  '   else
  '      response.write "              <tr>" & vbcrlf
  '      response.write "                  <th align=""left"" colspan=""3"">Delete an Event Category</th>" & vbcrlf
  '      response.write "              </tr>" & vbcrlf
  '      response.write "              <tr>" & vbcrlf
  '      response.write "                  <td nowrap>Category Name:</td>" & vbcrlf
  '      response.write "                  <td nowrap>" & vbcrlf
  '      response.write "                      <select name=""Category"" id=""Category"" class=""time"">" & vbcrlf
  '      response.write "                        <option value=""0"">None</option>" & vbcrlf
  '                                              displayCategoryOptions session("orgid"), lcl_calendarfeature
  '      response.write "                      </select>" & vbcrlf
  '      response.write "                  </td>" & vbcrlf
  '      response.write "                  <td width=""85%"">&nbsp;</td>" & vbcrlf
  '      response.write "              </tr>" & vbcrlf
  '   end if

  '   response.write "            </table>" & vbcrlf
  '   response.write "            </div>" & vbcrlf

  '   if lcl_sessionOrgIDisZero then
  '      displayButtons "DELETE", "BOTTOM"
  '   end if

  '   response.write "          </form>" & vbcrlf
  'end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!-- #include file="../admin_footer.asp" //-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayButtons(p_button_type, p_location)

  dim lcl_button_type, lcl_location, lcl_buttonClass

  lcl_button_type  = ""
  lcl_location     = "TOP"
  lcl_buttonClass  = "buttonRowTop"

  if p_button_type <> "" then
     if not containsApostrophe(p_button_type) then
        lcl_button_type = ucase(p_button_type)
     end if
  end if

  if lcl_button_type <> "" then
     if p_location <> "" then
        if not containsApostrophe(p_location) then
           lcl_location = ucase(p_location)
        end if
     end if

     if lcl_location = "BOTTOM" then
        lcl_buttonClass = "buttonRowBottow"
     end if

     response.write "<div class=""" & lcl_buttonClass & """>" & vbcrlf
     response.write "  <input type=""button"" value=""Cancel"" class=""button"" onclick=""returnToEvents();"" />" & vbcrlf

     if lcl_button_type = "CREATE" then
        response.write "  <input type=""button"" value=""Create"" class=""button"" onclick=""createCategory();"" />" & vbcrlf
     elseif lcl_button_type = "SAVE" then
        'response.write "  <input type=""button"" value=""Delete"" class=""button"" onclick=""document.getElementById('edit_task').value='deletecategory';document.ModifyEventCategory2.submit();"" />" & vbcrlf
        response.write "  <input type=""button"" value=""Save Changes"" class=""button"" onclick=""saveChanges()"">" & vbcrlf
     'elseif lcl_button_type = "DELETE" then
     '   response.write "  <input type=""button"" value=""Delete"" class=""button"" onclick=""if(fnCheckCategory()) {document.all.DeleteEventCategory.submit();} else {alert('Please enter a subject!');}"" />" & vbcrlf
     'elseif lcl_button_type = "EDIT" then
     '   response.write "  <input type=""button"" value=""Cancel"" class=""button"" onclick=""returnToEvents();"" />" & vbcrlf
     '   response.write "  <input type=""button"" value=""Save Changes"" class=""button"" onclick=""saveChanges2();"" />" & vbcrlf
     'elseif lcl_button_type = "DELETE2" then
     '   response.write "  <input type=""button"" value=""Delete"" class=""button"" onclick=""document.getElementById('edit_task').value='deletecategory';document.getElementById('category').value=" & p_categoryid & ";document.ModifyEventCategory2.submit();"" />" & vbcrlf
     end if

     response.write "</div>" & vbcrlf

  end if

end sub

'------------------------------------------------------------------------------
sub displayCategoryOptions(iOrgID, iCalendarFeature)

  dim sOrgID, sCalendarFeature

  sOrgID           = 0
  sCalendarFeature = ""

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iCalendarFeature <> "" then
     sCalendarFeature = ucase(iCalendarFeature)
     sCalendarFeature = dbsafe(sCalendarFeature)
     sCalendarFeature = "'" & sCalendarFeature & "'"
  end if

  if sOrgID > 0 then
     sSQL = "SELECT "
     sSQL = sSQL & " categoryid, "
     sSQL = sSQL & " categoryname, "
     sSQL = sSQL & " orgid, "
     sSQL = sSQL & " color, "
     sSQL = sSQL & " calendarfeature "
     sSQL = sSQL & " FROM eventcategories "
     sSQL = sSQL & " WHERE orgid = " & sOrgID

     if sCalendarFeature <> "" then
        sSQL = sSQL & " AND upper(calendarfeature) = " & sCalendarFeature
     else
        sSQL = sSQL & " AND (calendarfeature = '' OR calendarfeature IS NULL) "
     end if

     set oGetCategoryOptions = Server.CreateObject("ADODB.Recordset")
     oGetCategoryOptions.Open sSQL, Application("DSN"), 3, 1

     if not oGetCategoryOptions.eof then
        do while not oGetCategoryOptions.eof
           response.write "<option value=""" & oGetCategoryOptions("CategoryID") & """>" & oGetCategoryOptions("CategoryName") & "</option>" & vbcrlf

           oGetCategoryOptions.movenext
        loop
     end if

     oGetCategoryOptions.close
	 		 set oGetCategoryOptions = nothing

  end if

end sub

'------------------------------------------------------------------------------
 sub displayColorOptions()

	   response.write "<option value=""#000000"">Black</option>"                           & vbcrlf
	   response.write "<option value=""#0000ff"" class=""color0000ff"">Blue</option>"      & vbcrlf
	   response.write "<option value=""#006600"" class=""color006600"">Green</option>"     & vbcrlf
	   response.write "<option value=""#ff33cc"" class=""colorff33cc"">Magenta</option>"   & vbcrlf
	   response.write "<option value=""#ff9900"" class=""colorff9900"">Orange</option>"    & vbcrlf
	   response.write "<option value=""#C76309"" class=""colorC76309"">Dark Orange</option>" & vbcrlf
	   response.write "<option value=""#9933cc"" class=""color9933cc"">Purple</option>"    & vbcrlf
	   response.write "<option value=""#ff0000"" class=""colorff0000"">Red</option>"       & vbcrlf
	   response.write "<option value=""#0099ff"" class=""color0099ff"">Turquoise</option>" & vbcrlf

 end sub

'------------------------------------------------------------------------------
sub displayCategoryRows(iOrgID, iCalendarFeature)

  dim sOrgID, sCalendarFeature, lcl_bgcolor, lcl_linecount

  sOrgID           = 0
  sCalendarFeature = ""
  lcl_bgcolor      = "#ffffff"
  lcl_linecount    = 0

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iCalendarFeature <> "" then
     sCalendarFeature = ucase(iCalendarFeature)
     sCalendarFeature = dbsafe(sCalendarFeature)
     sCalendarFeature = "'" & sCalendarFeature & "'"
  end if

  if sOrgID > 0 then
     sSQL = "SELECT "
     sSQL = sSQL & " categoryid, "
     sSQL = sSQL & " categoryname, "
     sSQL = sSQL & " orgid, "
     sSQL = sSQL & " color, "
     sSQL = sSQL & " calendarfeature "
     sSQL = sSQL & " FROM eventcategories "
     sSQL = sSQL & " WHERE orgid = " & sOrgID

     if sCalendarFeature <> "" then
        sSQL = sSQL & " AND upper(calendarfeature) = " & sCalendarFeature
     else
        sSQL = sSQL & " AND (calendarfeature = '' OR calendarfeature IS NULL) "
     end if

     set oGetCategoryRows = Server.CreateObject("ADODB.Recordset")
     oGetCategoryRows.Open sSQL, Application("DSN"), 3, 1

     if not oGetCategoryRows.eof then
        do while not oGetCategoryRows.eof
           lcl_bgcolor   = changeBGColor(lcl_bgcolor,"#efefef","#ffffff")
           lcl_linecount = lcl_linecount + 1

           response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
           response.write "      <td class=""nowrap""><font color=""" & oGetCategoryRows("color") & """>" & oGetCategoryRows("categoryname") & "</td>" & vbcrlf
           response.write "      <td class=""nowrap"">" & vbcrlf
           response.write "          Change Category Name to:" & vbcrlf
           response.write "          <input type=""text"" name=""CustomCategory_" & lcl_linecount & """ id=""CustomCategory_" & lcl_linecount & """ class=""CustomCategory"" maxlength=""50"" />" & vbcrlf
           response.write "          <input type=""hidden"" name=""categoryid_" & lcl_linecount & """ id=""categoryid_" & lcl_linecount & """ value=""" & oGetCategoryRows("categoryid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "      <td class=""nowrap"">" & vbcrlf
           response.write "          Change Color to:" & vbcrlf
           response.write "          <select name=""CustomColor_" & lcl_linecount & """ id=""CustomColor_" & lcl_linecount & """>" & vbcrlf
           response.write "            <option value=""""></option>" & vbcrlf
                                       displayColorOptions
           response.write "          </select>" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "      <td align=""center"">" & vbcrlf
           response.write "          <input type=""checkbox"" name=""deleteCategory_" & lcl_linecount & """ id=""deleteCategory_" & lcl_linecount & """ value=""" & oGetCategoryRows("categoryid") & """ onclick=""deleteCategory('" & lcl_linecount & "');"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "  </tr>" & vbcrlf

           'response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
           'response.write "      <td><input type=""checkbox"" class=""listcheck"" name=""categoryid_" & oGetCategoryRows("categoryid") & """ id=""categoryid_" & oGetCategoryRows("categoryid") & """ value=""" & oGetCategoryRows("categoryid") & """ /></td>" & vbcrlf
           'response.write "      <td class=""nowrap""><font color=""" & oGetCategoryRows("color") & """>" & oGetCategoryRows("categoryname") & "</td>" & vbcrlf
           'response.write "      <td class=""nowrap"">" & vbcrlf
           'response.write "          Change Category Name to:" & vbcrlf
           'response.write "          <input type=""text"" name=""CustomCategory_" & oGetCategoryRows("categoryid") & """ id=""CustomCategory_" & oGetCategoryRows("categoryid") & """ class=""CustomCategory"" maxlength=""50"" />" & vbcrlf
           'response.write "      </td>" & vbcrlf
           'response.write "      <td class=""nowrap"">" & vbcrlf
           'response.write "          Change Color to:" & vbcrlf
           'response.write "          <select name=""CustomColor_" & oGetCategoryRows("categoryid") & """ id=""CustomColor_" & oGetCategoryRows("categoryid") & """>" & vbcrlf
           'response.write "            <option value=""""></option>" & vbcrlf
           '                            displayColorOptions
           'response.write "          </select>" & vbcrlf
           'response.write "      </td>" & vbcrlf
           'response.write "      <td align=""center"">" & vbcrlf
           'response.write "          <input type=""checkbox"" name=""deleteCategory_" & oGetCategoryRows("categoryid") & """ id=""deleteCategory_" & oGetCategoryRows("categoryid") & """ value=""Y"" />" & vbcrlf
           'response.write "      </td>" & vbcrlf
           'response.write "  </tr>" & vbcrlf

           oGetCategoryRows.movenext
			     loop
     end if

     oGetCategoryRows.close
			  set oGetCategoryRows = nothing

  end if

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""4"">" & vbcrlf
  response.write "          <input type=""hidden"" name=""totalCategories"" id=""totalCategories"" value=""" & lcl_linecount & """ size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

end sub
%>
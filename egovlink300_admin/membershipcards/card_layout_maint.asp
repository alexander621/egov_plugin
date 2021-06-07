<!-- #include file="../includes/common.asp" //-->
<!-- #include file="membership_card_functions.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("memberships") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel      = "../"     'Override of value from common.asp
 lcl_hidden  = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

 if not userhaspermission(session("userid"),"card_layout_maint") then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the action
 if request("formAction") <> "" then
    lcl_action = UCASE(request("formAction"))
 else
    lcl_action = ""
 end if

'Retrieve the CardID
 if request("cardid") <> "" then
    if isnumeric(request("cardid")) then
       lcl_cardid = request("cardid")
    else
       response.redirect "../default.asp"
    end if
 end if

 if lcl_cardid = "" then
    if lcl_action = "ADD" then
       lcl_cardid = 0
    else
       lcl_cardid = getMaxCardID()
    end if
 end if

'Set up parameters on card that require a memberid
 lcl_layout_maint     = "Y"
 lcl_layout_memberid  = "12345"
 lcl_layout_fname     = "Member"
 lcl_layout_lname     = "Name"

'Determine if the user is updating or creating a new layout
 if (lcl_action = "PREVIEW CHANGES" OR lcl_action = "SAVE CHANGES" OR lcl_action = "COPY LAYOUT") AND lcl_action <> "RESET VALUES" then
     lcl_title                = request("p_new_title")
     lcl_subtitle             = request("p_new_subtitle")
     lcl_year_text            = request("p_new_year_text")
     lcl_display_date         = request("p_display_date")
     lcl_custom_image_url     = request("p_new_custom_image_url")
     lcl_quote                = request("p_new_quote")
     lcl_color1               = request("p_new_color1")
     lcl_color2               = request("p_new_color2")
     lcl_text_color1          = request("p_new_text_color1")
     lcl_text_color2          = request("p_new_text_color2")
     lcl_back_text            = request("p_new_back_text")
     lcl_back_text_color      = request("p_new_back_text_color")
     lcl_layoutname           = request("p_new_layoutname")
     lcl_new_isdisabled       = request("p_new_isDisabled")
     lcl_new_confirmsoundfile = request("p_new_confirmsoundfile")

     if lcl_action = "COPY LAYOUT" then
        lcl_layoutname = getOriginalLayoutName(lcl_cardid,"COPY OF " & lcl_layoutname)
        lcl_cardid     = 0
     end if

    'Determine the user's actions
     if lcl_action = "SAVE CHANGES" OR lcl_action = "COPY LAYOUT" then

        'if request("cardid") <> "" AND CLng(request("cardid")) > CLng(0) then
        if lcl_cardid <> 0 then
           if checkDuplicateLayoutName(lcl_cardid,lcl_layoutname) then
              alreadyExists lcl_cardid, lcl_title, lcl_subtitle, lcl_year_text, lcl_display_date, lcl_custom_image_url, _
                            lcl_quote, lcl_color1, lcl_color2, lcl_text_color1, lcl_text_color2, lcl_back_text, lcl_back_text_color, _
                            lcl_layoutname, lcl_new_isdisabled, lcl_new_confirmsoundfile
           else
              updateCardLayout lcl_cardid, lcl_title, lcl_subtitle, lcl_year_text, lcl_display_date, lcl_custom_image_url, _
                               lcl_quote, lcl_color1, lcl_color2, lcl_text_color1, lcl_text_color2, lcl_back_text, lcl_back_text_color, _
                               lcl_layoutname, lcl_new_isdisabled, lcl_new_confirmsoundfile
           end if

        else

          'Check to see if a card layout record exists for the orgid
           if checkDuplicateLayoutName(lcl_cardid,lcl_layoutname) then
              alreadyExists lcl_cardid, lcl_title, lcl_subtitle, lcl_year_text, lcl_display_date, lcl_custom_image_url, _
                            lcl_quote, lcl_color1, lcl_color2, lcl_text_color1, lcl_text_color2, lcl_back_text, lcl_back_text_color, _
                            lcl_layoutname, lcl_new_isdisabled, lcl_new_confirmsoundfile
           else
              createCardLayout lcl_title, lcl_subtitle, lcl_year_text, lcl_display_date, lcl_custom_image_url, lcl_quote, lcl_color1, _
                               lcl_color2, lcl_text_color1, lcl_text_color2, lcl_back_text, lcl_back_text_color, lcl_layoutname, _
                               lcl_action, lcl_new_isdisabled, lcl_new_confirmsoundfile
           end if

        end if
     end if

 elseif lcl_action = "DELETE LAYOUT" then
     deleteCardLayout lcl_cardid
 else

    'Set up the editable parameters.  First retrieve the original values
     sSQL = "SELECT title, subtitle, year_text, display_date, custom_image_url, quote_text, main_color, secondary_color, "
     sSQL = sSQL & " main_text_color, secondary_text_color, back_text, back_text_color, "
     sSQL = sSQL & " isnull(layoutname,'[No Layout Name Available]') AS layoutname, isDisabled, "
     sSQL = sSQL & " isnull(confirmsoundfile,'') as confirmsoundfile "
     sSQL = sSQL & " FROM egov_membershipcard_layout "
     sSQL = sSQL & " WHERE orgid = " & session("orgid")
     sSQL = sSQL & " AND cardid = " & lcl_cardid

     set oOrigVal = Server.CreateObject("ADODB.Recordset")
     oOrigVal.Open sSQL, Application("DSN"), 3, 1

     if not oOrigVal.eof then
        lcl_original_title            = oOrigVal("title")
        lcl_original_subtitle         = oOrigVal("subtitle")
        lcl_original_year_text        = oOrigVal("year_text")
        lcl_original_display_date     = oOrigVal("display_date")
        lcl_original_custom_image_url = oOrigVal("custom_image_url")
        lcl_original_quote            = oOrigVal("quote_text")
        lcl_original_color1           = oOrigVal("main_color")
        lcl_original_color2           = oOrigVal("secondary_color")
        lcl_original_text_color1      = oOrigVal("main_text_color")
        lcl_original_text_color2      = oOrigVal("secondary_text_color")
        lcl_original_back_text        = oOrigVal("back_text")
        lcl_original_back_text_color  = oOrigVal("back_text_color")
        lcl_original_layoutname       = oOrigVal("layoutname")
        lcl_original_isDisabled       = oOrigVal("isDisabled")
        lcl_original_confirmsoundfile = oOrigVal("confirmsoundfile")
    else
        lcl_original_title            = "CITY NAME"
        lcl_original_subtitle         = "Pool Pass"
        lcl_original_year_text        = "2008 Pool Member"
        lcl_original_display_date     = 1
        lcl_original_custom_image_url = ""
        lcl_original_quote            = "<strong>QUOTE!</strong>"
        lcl_original_color1           = "FFFF00"
        lcl_original_color2           = "0000C0"
        lcl_original_text_color1      = "000000"
        lcl_original_text_color2      = "FFFFFF"
        lcl_original_back_text        = ""
        lcl_original_back_text_color  = "000000"
        lcl_original_layoutname       = getOriginalLayoutName(lcl_cardid,"New Card Layout")
        lcl_original_isDisabled       = False
        lcl_original_confirmsoundfile = ""

    end if

    oOrigVal.close
    set oOrigVal = nothing

    'Set up the display fields with the values on the database
     lcl_title            = lcl_original_title
     lcl_subtitle         = lcl_original_subtitle
     lcl_year_text        = lcl_original_year_text
     lcl_display_date     = lcl_original_display_date
     lcl_custom_image_url = lcl_original_custom_image_url
     lcl_quote            = lcl_original_quote
     lcl_color1           = lcl_original_color1
     lcl_color2           = lcl_original_color2
     lcl_text_color1      = lcl_original_text_color1
     lcl_text_color2      = lcl_original_text_color2
     lcl_back_text        = lcl_original_back_text
     lcl_back_text_color  = lcl_original_back_text_color
     lcl_layoutname       = lcl_original_layoutname
     lcl_isDisabled       = lcl_original_isDisabled
     lcl_confirmsoundfile = lcl_original_confirmsoundfile

 end if

'Determine if there is a screen message to display
 lcl_success = ""
 lcl_message = ""

 if request("success") <> "" then
    lcl_success = UCASE(request("success"))
 end if

 if lcl_success = "SU" then
    lcl_message = "<span style=""font-size:12px; font-weight:bold; color:#ff0000"">*** Successfully Updated... ***</span>"
 elseif lcl_success = "SA" then
    lcl_message = "<span style=""font-size:12px; font-weight:bold; color:#ff0000"">*** Successfully Created... ***</span>"
 elseif lcl_success = "SC" then
    lcl_message = "<span style=""font-size:12px; font-weight:bold; color:#ff0000"">*** Successfully Copied... ***</span>"
 elseif lcl_success = "AE" then
    lcl_message = "<span style=""font-size:12px; font-weight:bold; color:#ff0000"">*** A Card Layout with this Layout Name already exists. ***</span>"
 elseif lcl_success = "SD" then
    lcl_message = "<span style=""font-size:12px; font-weight:bold; color:#ff0000"">*** Successfully Deleted... ***</span>"
 elseif lcl_success = "NODEL" then
    lcl_message = "<span style=""font-size:12px; font-weight:bold; color:#ff0000"">*** This card layout cannot be deleted as it exists on atleast one rate that has been purchased. ***</span>"
 end if

'Set the database character length for the card BACK_TEXT field to be used throughout code
 lcl_cardback_text_length = 1000

'Check for org features
 lcl_orghasfeature_card_layout_multiplelayouts = orghasfeature("card_layout_multiplelayouts")
%>
<html>
<head>
	<title>E-Gov Administration Console {Maintain Membership Card Layout}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="membership_card.css" />	

	<script language="javascript" src="validator.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
<script language="javascript">
function doSitePicker(sFormField) {
  lcl_width  = 600;
  lcl_height = 470;
  lcl_left   = (screen.availWidth/2) - (lcl_width/2);
  lcl_top    = (screen.availHeight/2) - (lcl_height/2);
		eval('window.open("../sitelinker/default.asp?name=' + sFormField + '", "_dositepicker", "width=' + lcl_width + ',height=' + lcl_height + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + lcl_left + ',top=' + lcl_top + '")');
}

function doPicker(iFieldID) {
  lcl_width  = 600;
  lcl_height = 400;
  lcl_left   = (screen.availWidth/2)-(lcl_width/2);
  lcl_top    = (screen.availHeight/2)-(lcl_height/2);
  eval('window.open("linkpicker/linkpicker.asp?fid=' + iFieldID + '", "_dopicker", "width=' + lcl_width + ',height=' + lcl_height + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + lcl_left + ',top=' + lcl_top + '")');
}

function storeCaret (textEl) {
  if (textEl.createTextRange)
      textEl.caretPos = document.selection.createRange().duplicate();
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
      var caretPos = textEl.caretPos;
      caretPos.text =
      caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
      text + ' ' : text;
  }
   else
      textEl.value  = text;
}

function print_card_back() {
  window.open("card_print.asp?print_side=BACK");
}
function openWin(p_page,p_field_id) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
		eval('window.open("' + p_page + '?fieldid=' + p_field_id + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function changeLayout(iValue) {
  location.href="card_layout_maint.asp?cardid=" + iValue;
}

function submitForm(iValue) {
<%
 'If the org does not have the "Multiple Card Layouts" feature then hide the delete javascript option
  if lcl_orghasfeature_card_layout_multiplelayouts then
     response.write "if(iValue==""Delete Layout"") {" & vbcrlf

     if checkCardonRatePurchased(lcl_cardid) = 0 then
        response.write "   input_box = confirm('Are you sure you want to delete this card layout?');" & vbcrlf
        response.write "   if(input_box==true) {" & vbcrlf
        response.write "      document.getElementById(""formAction"").value = iValue;" & vbcrlf
        response.write "      document.getElementById(""card_maint"").submit();" & vbcrlf
        response.write "   }" & vbcrlf
     else
        response.write "   alert(""This card layout cannot be deleted as it exists on atleast one rate that has been purchased."");" & vbcrlf
     end if

     response.write "}else{" & vbcrlf
     response.write "   document.getElementById(""formAction"").value = iValue;" & vbcrlf
     response.write "   document.getElementById(""card_maint"").submit();" & vbcrlf
     response.write "}" & vbcrlf
  else
     response.write "   document.getElementById(""formAction"").value = iValue;" & vbcrlf
     response.write "   document.getElementById(""card_maint"").submit();" & vbcrlf
  end if
%>
}
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="checkFieldLength(document.getElementById('p_new_back_text').value,<%=lcl_cardback_text_length%>,'Y',document.getElementById('p_new_back_text'))">

	<% ShowHeader sLevel %>
	<!--#include file="../menu/menu.asp"--> 

<div style="margin-left:10px">
  <p><font size="+1"><strong>Maintain Membership Card Layout</strong></font></p>
</div>

<table border="0" cellspacing="0" cellpadding="5">
  <form name="card_maint" id="card_maint" action="card_layout_maint.asp" method="post">
    <input type="hidden" name="formAction" id="formAction" value="" size="10" maxlength="50" />
    <input type="hidden" name="control_field" id="control_field" value="" size="10" maxlength="<%=lcl_cardback_text_length+1%>" />

<%
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf

 'check if the org allows multiple layouts
  if lcl_orghasfeature_card_layout_multiplelayouts then
     response.write "          <div>" & vbcrlf

     if lcl_action <> "ADD" then
        response.write "            Card Layouts: " & vbcrlf
        response.write "            <select name=""cardid"" id=""cardid"" onchange=""changeLayout(this.value)"">" & vbcrlf

        displayCardLayoutOptions lcl_cardid,"N"

        response.write "            </select>" & vbcrlf

        lcl_layoutname_label = "Change "
     else
        response.write "<input type=""hidden"" name=""cardid"" id=""cardid"" value=""" & lcl_cardid & """ size=""10"" maxlength=""50"" />" & vbcrlf
        lcl_layoutname_label = ""
     end if

     response.write "            &nbsp;&nbsp;&nbsp;" & vbcrlf
     response.write "            <span style=""color:#800000"">" & lcl_layoutname_label & "Layout Name: </span>" & vbcrlf
     response.write "            <input type=""text"" name=""p_new_layoutname"" id=""p_new_layoutname"" value=""" & lcl_layoutname & """ size=""30"" maxlength=""50"" />" & vbcrlf
     response.write "            <br /><br />" & vbcrlf
     response.write "          </div>" & vbcrlf

  else

     response.write "<input type=""hidden"" name=""cardid"" id=""cardid"" value=""" & lcl_cardid & """ size=""10"" maxlength=""50"" />" & vbcrlf
     response.write "<input type=""hidden"" name=""p_new_layoutname"" id=""p_new_layoutname"" value=""" & lcl_layoutname & """ size=""30"" maxlength=""50"" />" & vbcrlf
  end if

  displayButtons lcl_action

 'Determine if there is a screen message to display.
  if lcl_message <> "" then
      response.write "<div align=""right"">" & lcl_message & "</div>" & vbcrlf
  end if

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
   %>
  <tr valign="top">
      <td width="50%">
          <fieldset>
              <legend>Edit Card FRONT&nbsp;</legend>
              <table border="1" bordercolor="#C0C0C0" cellspacing="0" cellpadding="2">
                <tr bgcolor="#efefef">
                    <td colspan="2"><strong>Member Name</strong></td>
                    <td rowspan="2" align="center" width="40"><strong>B<br />A<br />R<br />C<br />O<br />D<br />E</strong></td>
                </tr>
                <tr valign="top">
                    <td width="45%">
                        <table border="0" cellspacing="0" cellpadding="2" width="100%">
                          <tr><td><input type="text" name="p_new_year_text" value="<%=lcl_year_text%>" /></td></tr>
                          <tr><td>&nbsp;</td></tr>
                          <tr>
                              <td>
                                  <table border="0" cellspacing="0" cellpadding="0" width="120" height="90" style="border: 1px solid #000000">
                                    <tr><td align="center" bgcolor="#efefef">Picture Here</td></tr>
                                  </table>
                              </td>
                          </tr>
                          <tr><td>&nbsp;</td></tr>
                          <tr>
                              <td><select name="p_display_date">
                                    <option value="0"></option>
                                  <%
                                    if CLng(lcl_display_date) = CLng(1) then
                                       lcl_selected_1 = " selected"
                                       lcl_selected_2 = ""
                                    elseif CLng(lcl_display_date) = CLng(2) then
                                       lcl_selected_1 = ""
                                       lcl_selected_2 = " selected"
                                    else
                                       lcl_selected_1 = ""
                                       lcl_selected_2 = ""
                                    end if
                                  %>
                                    <option value="1"<%=lcl_selected_1%>>Show Expiration Date</option>
                                    <option value="2"<%=lcl_selected_2%>>Show Issued Date</option>
                              </td></tr>
                        </table>
                    </td>
                    <td width="45%">
                        <table border="0" cellspacing="0" cellpadding="2" width="100%">
                          <tr><td>&nbsp;</td></tr>
                          <tr><td><input type="text" name="p_new_title" value="<%=lcl_title%>" /></td></tr>
                          <tr><td><input type="text" name="p_new_subtitle" value="<%=lcl_subtitle%>" /></td></tr>
                          <tr><td>
                                  <table border="0" cellspacing="0" cellpadding="0" width="100" height="50" style="border: 1px solid #000000">
                                    <tr><td align="center" bgcolor="#efefef">Custom Logo Here</td></tr>
                                  </table>
                              </td></tr>
                          <tr><td>&nbsp;</td></tr>
                          <tr><td><input type="text" name="p_new_quote" value="<%=lcl_quote%>" /></td></tr>
                        </table>
                    </td>
                </tr>
              </table>
              <table border="0" cellspacing="1" cellpadding="2">
                <tr>
                    <td nowrap="nowrap">Custom Logo:</td>
                    <td colspan="2"><input type="text" id="p_new_custom_image_url" name="p_new_custom_image_url" value="<%=lcl_custom_image_url%>" size="30" maxlength="500" /></td>
                    <td><input type="button" class="button" value="Find a Link" onClick="doSitePicker('card_maint.p_new_custom_image_url');" /></td>
                </tr>
                <tr>
                    <td nowrap="nowrap">Main Color:</td>
                    <td><input type="text" name="p_new_color1" id="p_new_color1" value="<%=lcl_color1%>" size="10" maxlength="6" /></td>
                    <td width="40%" bgcolor="<%=lcl_color1%>" style="border: 1px solid #000000;">&nbsp;</td>
                    <td nowrap="nowrap" align="center"><a href="javascript:openWin('../colorpalette.asp','p_new_color1')">[select color]</a></td>
                </tr>
                <tr>
                    <td nowrap="nowrap">Secondary Color:</td>
                    <td><input type="text" name="p_new_color2" id="p_new_color2" value="<%=lcl_color2%>" size="10" maxlength="6" /></td>
                    <td width="40%" bgcolor="<%=lcl_color2%>" style="border: 1px solid #000000;">&nbsp;</td>
                    <td nowrap="nowrap" align="center"><a href="javascript:openWin('../colorpalette.asp','p_new_color2')">[select color]</a></td>
                </tr>

                <tr>
                    <td nowrap="nowrap">Font Color (Main):</td>
                    <td><input type="text" name="p_new_text_color1" id="p_new_text_color1" value="<%=lcl_text_color1%>" size="10" maxlength="6" /></td>
                    <td width="40%" bgcolor="<%=lcl_text_color1%>" style="border: 1px solid #000000;">&nbsp;</td>
                    <td nowrap="nowrap" align="center"><a href="javascript:openWin('../colorpalette.asp','p_new_text_color1')">[select color]</a></td>
                </tr>
                <tr>
                    <td nowrap="nowrap">Font Color (Secondary):</td>
                    <td><input type="text" name="p_new_text_color2" id="p_new_text_color2" value="<%=lcl_text_color2%>" size="10" maxlength="6" /></td>
                    <td width="40%" bgcolor="<%=lcl_text_color2%>" style="border: 1px solid #000000;">&nbsp;</td>
                    <td nowrap="nowrap" align="center"><a href="javascript:openWin('../colorpalette.asp','p_new_text_color2')">[select color]</a></td>
                </tr>
              </table>
          </fieldset>
      </td>
      <td>
          <fieldset>
              <legend>Preview Card FRONT&nbsp;</legend>
              <table border="0" cellspacing="0" cellpadding="2">
                <tr valign="top">
                    <td width="350" height="180">
                        <div style="position: relative; width: 320px; height: 180px;"><!--#include file="membership_card.asp"--></div>
                    </td>
                </tr>
              </table>
          </fieldset>
          <br />
        <%
         'The "isDisabled" field works opposite from the wording.  On the screen the field is labeled: "Active Layout".
         '  Therefore, if "checked" we want to "enable" the layout but we have to set the column to "False" on the table.
         '  If "unchecked" the column must be set to "True" on the table.
         '  Column Values: "True" = disabled, "False" = enabled.
          if lcl_isDisabled then
             lcl_checked_isDisabled = ""
          else
             lcl_checked_isDisabled = " checked=""checked"""
          end if

          if lcl_orghasfeature_card_layout_multiplelayouts then
             response.write "Active Layout:&nbsp;" & vbcrlf
             response.write "<input type=""checkbox"" name=""p_new_isDisabled"" id=""p_new_isDisabled"" value=""on""" & lcl_checked_isDisabled & " />" & vbcrlf
             response.write "<br />" & vbcrlf

             response.write "Confirmation Sound File:&nbsp;" & vbcrlf
             response.write "<input type=""text"" name=""p_new_confirmsoundfile"" id=""p_new_confirmsoundfile"" size=""25"" maxlength=""500"" value=""" & lcl_confirmsoundfile & """ />" & vbcrlf
             response.write "<input type=""button"" name=""findSoundFile"" id=""findSoundFile"" value=""Find a Sound"" class=""button"" onclick=""doPicker('p_new_confirmsoundfile');"" />" & vbcrlf
             response.write "<br />" & vbcrlf
             response.write "<div align=""center"" style=""font-size:8pt; font-style:italic; color:#800000;"">" & vbcrlf
             response.write "(leave blank for default confirmation sound)" & vbcrlf
             response.write "</div>" & vbcrlf
             response.write "<br />" & vbcrlf
          else
             response.write "<input type=""hidden"" name=""p_new_isDisabled"" id=""p_new_isDisabled"" value=""on"" />" & vbcrlf
             response.write "<input type=""hidden"" name=""p_new_confirmsoundfile"" id=""p_new_confirmsoundfile"" value="""" size=""10"" maxlength=""500"" />" & vbcrlf
          end if
        %>
      </td>
  </tr>
  <tr valign="top">
      <td width="50%">
          <fieldset>
              <legend>Edit Card BACK&nbsp;</legend>
              <table border="0" cellspacing="0" cellpadding="2">
                <tr valign="top">
                    <td width="65%">
                        <textarea name="p_new_back_text" id="p_new_back_text" rows="8" cols="50" style="font-size: 8pt; width: 250px; height: 150px;" onkeydown="document.getElementById('control_field').value=this.value;" onkeyup="javascript:checkFieldLength(this.value,<%=lcl_cardback_text_length%>,'Y',this)"><%=lcl_back_text%></textarea>
                        <div id="message_char_cnt"><%=lcl_cardback_text_length%> character limit.<br />Characters remaining: <%=lcl_cardback_text_length%></div>
                    </td>
                    <td>
                        <input type="button" class="button" value="Find a Link" onclick="doSitePicker('card_maint.p_new_back_text');" /><p>
                        <table border="0" cellspacing="1" cellpadding="2" width="100%">
                          <tr>
                              <td>Font Color:</td>
                              <td><input type="text" name="p_new_back_text_color" id="p_new_back_text_color" value="<%=lcl_back_text_color%>" size="10" maxlength="6" /></td>
                          </tr>
                          <tr>
                              <td bgcolor="<%=lcl_back_text_color%>" style="border: 1px solid #000000;">&nbsp;</td>
                              <td colspan="2" nowrap="nowrap" align="center"><a href="javascript:openWin('../colorpalette.asp','p_new_back_text_color')">[select color]</a></td>
                          </tr>
                        </table>
                    </td>
                </tr>
              </table>
          </fieldset>
      </td>
      <td>
          <fieldset>
              <legend>Preview Card BACK&nbsp;</legend>
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td align="center" valign="middle">
                        <div style="position: relative; width: 300px; height: 176px; border: 1px solid #000000;">

                            <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%" style="font-size: 10pt;">
                              <tr>
                                  <td id="preview_card_back" align="center" valign="middle" style="color: #<%=lcl_back_text_color%>">
                                      <%=lcl_back_text%>
                                  </td>
                              </tr>
                            </table>

                        </div>
                    </td>
                    <td valign="top"><input type="button" name="printCardBack" id="printCardBack" value="Print Card Back" onclick="print_card_back();" /></td>
                </tr>
              </table>
          </fieldset>
      </td>
  </tr>
</table>

<% displayButtons lcl_action %>

</form>

<!--#include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displayButtons(iAction)

  if lcl_orghasfeature_card_layout_multiplelayouts then
     response.write "<input type=""button"" name=""deleteLayout"" id=""deleteLayout"" value=""Delete Layout"" onclick=""submitForm(this.value)"" />" & vbcrlf
  end if

  response.write "<input type=""button"" name=""resetValues"" id=""resetValues"" value=""Reset Values"" onclick=""submitForm(this.value)"" />" & vbcrlf
  response.write "<input type=""button"" name=""previewChanges"" id=""previewChanges"" value=""Preview Changes"" onclick=""submitForm(this.value)"" />" & vbcrlf

  if iAction <> "ADD" AND lcl_orghasfeature_card_layout_multiplelayouts then
     'response.write "<input type=""button"" name=""addLayout"" id=""addLayout"" value=""Add Layout"" onclick=""location.href='card_layout_maint.asp?formAction=add';"" />" & vbcrlf
     response.write "<input type=""button"" name=""copyLayout"" id=""copyLayout"" value=""Copy Layout"" onclick=""submitForm(this.value);"" />" & vbcrlf
  end if

  response.write "<input type=""button"" name=""saveChanges"" id=""saveChanges"" value=""Save Changes"" onclick=""submitForm(this.value)"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

  set oDTB = nothing

end sub
%>
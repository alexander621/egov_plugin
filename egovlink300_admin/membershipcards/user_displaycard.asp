<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="user_card_functions.asp" //-->
<%
 sLevel = "../"  'Override of value from common.asp

'Check to see if the feature is offline
 if isFeatureOffline("registration") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 if not userhaspermission(session("userid"),"create_user_membershipcards") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 dim lcl_userid, lcl_action, lcl_printcard_url

 lcl_userid = ""
 lcl_action = ""

 if request("userid") <> "" then
    lcl_userid = clng(request("userid"))
 end if

 if request("action") <> "" then
    if not containsApostrophe(request("action")) then
       lcl_action = request("action")
    end if
 end if

 lcl_printcard_url = lcl_printcard_url & "user_displaycard.asp"
 lcl_printcard_url = lcl_printcard_url & "?userid=" & lcl_userid
 lcl_printcard_url = lcl_printcard_url & "&action=CARD_PRINTED"
 lcl_printcard_url = lcl_printcard_url & "&card_layout=p"

'Set up Session variable for DISPLAY include file.
 session("CARD_PRINT") = "N"
' session("userid")     = lcl_userid

 if lcl_action = "CANCEL" then
    remove_image lcl_userid
   	response.redirect session("RedirectPage")
 elseif lcl_action = "REPRINT" or lcl_action = "PRINT" then
    lcl_status = lcl_action
   	lcl_action = "CARD_PRINTED"
 elseif lcl_action = "SAVE" then
    save_card session("orgid"), lcl_userid
   	response.redirect lcl_printcard_url
 elseif lcl_action = "PRINT_CARD" then
   	save_card  session("orgid"), lcl_userid
   	print_card session("orgid"), lcl_userid
    response.redirect lcl_printcard_url
 end if
%>
<html>
<head>
<title>E-Gov Administration Console {Membership Photo and ID Creation}</title>
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="user_card.css" />

  <script language="javascript" src="validator.js"></script>

<script language="javascript">
//function GoBack(ReturnToURL) {
//  if (ReturnToURL != "") {
function GoBack() {
  var lcl_return_page = '<%=session("redirectpage")%>';

  if (lcl_return_page != "") {
      location.href = lcl_return_page;
  } else {
      history.go(-1);
  }
}

function card_print() {
  var lcl_url_cardprint   = '';
  var lcl_url_carddisplay = '';

  lcl_url_cardprint += 'user_cardprint.asp';
  lcl_url_cardprint += '?userid=<%=lcl_userid%>';
  lcl_url_cardprint += '&status=PRINT';
  lcl_url_cardprint += '&card_layout=p';
  lcl_url_cardprint += '&initPrint=Y';
  lcl_url_cardprint += '&OS=XP';

  lcl_url_carddisplay += 'user_displaycard.asp';
  lcl_url_carddisplay += '?userid=<%=lcl_userid%>';
  lcl_url_carddisplay += '&action=PRINT_CARD';
  lcl_url_carddisplay += '&card_layout=p';

  window.open(lcl_url_cardprint);
  location.href = lcl_url_carddisplay;
}

function retake_picture() {
  location.href="user_takepic.asp?userid=<%=lcl_userid%>&reload_pic=Y";
}

function reload_picture() {
  if ("Y"=="<%=Session("RELOAD_PIC")%>") {
      window.location.reload();
<% session("RELOAD_PIC") = "N" %>
  }else{
     return true;
  }
}

function remove_image() {
  location.href = "user_displaycard.asp?userid=<%=lcl_userid%>&action=CANCEL";
}

function card_save() {
  location.href = "user_displaycard.asp?userid=<%=lcl_userid%>&action=SAVE&card_layout=p";
}
</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="reload_picture()">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <h3>Print Membership Card</h3>" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" class=""button"" value=""Return to List"" onclick=""GoBack();"" />" & vbcrlf
  response.write "          <p>" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "		          <tr valign=""top"">" & vbcrlf

 'BEGIN: Display Card ---------------------------------------------------------
  response.write "                <td width=""320"">" & vbcrlf
  response.write "                    <div id=""membershipcard"">" & vbcrlf
                                        displayCard session("orgid"), lcl_userid, lcl_status
  response.write "                    </div>" & vbcrlf
  response.write "                </td>" & vbcrlf
 'END: Display Card -----------------------------------------------------------

 'BEGIN: Display Instructions -------------------------------------------------
   if lcl_action <> "CARD_PRINTED" then
      lcl_print_label  = "Print Card"
   	  lcl_print_msg    = "Prints, and saves, the membership card.  Please have your color printer ready."
      lcl_cancel_label = "Cancel"
   	  lcl_cancel_msg   = "Return to the &quot;Create a Membership List&quot; without saving any of the data."
   	  lcl_cancel_url   = "remove_image()"
   else
      lcl_print_label  = "Reprint Card"
   	  lcl_print_msg    = "Prints, and saves, the membership card.  Please have your color printer ready."
      lcl_cancel_label = "Card Completed"
	     lcl_cancel_msg   = "The new membership card has been printed and the data has saved.  Click to<br />return to &quot;Create a Membership List&quot; results screen."
      lcl_cancel_url   = "GoBack()"
   end if

  response.write "                <td height=""200"">" & vbcrlf
  response.write "                    <p><strong>Button Instructions</strong></p>" & vbcrlf
  response.write "                				<ul>" & vbcrlf
  response.write "                      <li><strong>" & lcl_print_label & ": </strong>" & lcl_print_msg & "</li>" & vbcrlf

  if lcl_action <> "CARD_PRINTED" then
     response.write "                   <li><strong>Save Card: </strong>ONLY saves the membership card data so that it can be printed at another time.</li>" & vbcrlf
  end if

  response.write "                      <li><strong>Retake Picture: </strong>Click to retake the picture.</li>" & vbcrlf
  response.write "                      <li><strong>" & lcl_cancel_label & ": </strong>" & lcl_cancel_msg & "</li>" & vbcrlf
  response.write "                    </ul>" & vbcrlf
  response.write "                    <div align=""center"">" & vbcrlf
  response.write "                      <input type=""button"" value=""" & lcl_print_label & """ class=""noprint"" id=""card_print"" name=""card_print"" onclick=""javascript:card_print();"" />" & vbcrlf

  if lcl_action <> "CARD_PRINTED" then
     response.write "                   <input type=""button"" value=""Save Card"" class=""noprint"" id=""card_save"" name=""card_save"" onclick=""javascript:card_save();"" />" & vbcrlf
  end if

  response.write "                      <input type=""button"" value=""Retake Picture"" class=""noprint"" id=""retake_picture"" name=""retake_picture"" onclick=""retake_picture()"" />" & vbcrlf
  response.write "                      <input type=""button"" value=""" & lcl_cancel_label & """ class=""noprint"" id=""remove_image"" name=""remove_image"" onclick=""" & lcl_cancel_url & """ />" & vbcrlf
  response.write "                    </div>" & vbcrlf
  response.write "                </td>" & vbcrlf
 'END: Display Instructions ---------------------------------------------------

  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf

  lcl_cardprinted_count = getCardPrintedTotal(lcl_userid)

  if lcl_cardprinted_count > 0 then
     response.write "<div align=""center""># times Membership Card has been printed: " & lcl_cardprinted_count & "</div>" & vbcrlf
  end if

  response.write "          </p>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<p>" & vbcrlf
%>
	<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>

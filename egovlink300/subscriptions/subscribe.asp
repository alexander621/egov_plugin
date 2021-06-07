<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
twfpass = session("userpassword")

%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="subscribe_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: subscribe.asp
' AUTHOR: Steve Loar
' CREATED: 09/06/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Subscription registration page.
'
' MODIFICATION HISTORY
' 1.0 09/06/06 Steve Loar - Initial version
' 1.1 02/04/08 David Boyer - Added Job/Bid Postings
' 2.0 09/29/08 David Boyer - Rewrote page to no longer use ProcessRecords().
' 2.1	01/12/09	Steve Loar	-	Changed processrecords() to input the relationship id that lets these 
'                           be converted into regular registered user records.
' 2.2 02/23/10 David Boyer - Modified the new user sign up to auto-subscribe the user to subscriptions
'                            and modified the confirmation message.
' 2.3 2014-06-11 Jerry Felix - Simplified the regex for email validation (new TLDs were failing)
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Determine which list to display
if request("listtype") <> "" then
	lcl_list_type = request("listtype")
else
	lcl_list_type = ""
end if

if not orghasfeature(iorgid,"subscriptions") then
response.redirect "../"
end if

dim subscriberuserid, sError, iUserID, sMsg, oUser
dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets	

sMsg          = ""
iUserID       = 0
iUserEmail    = ""
iUserPassword = ""
iConfirmID    = ""

'Check for a userid
if request("u") <> "" then
	iUserID = request("u")
end if

'Check to see if the user is entering via link in new subscription email (unsubscribe.asp)
if request("c") <> "" then
if isnumeric(request("c")) then
iConfirmID = request("c")
end if
end if

lcl_success = request("success")

if (iUserID <> "" OR iConfirmID <> "") AND lcl_success <> "EXISTS_BAD_PWD" AND lcl_success <> "NOT_EXISTS" AND	lcl_success <> "MANAGE" AND lcl_success <> "BAD_PWD" then
	'If the user exists then attempt to bring in their email/password
	checkUserExists iUserID, "", iorgid, iConfirmID, lcl_user_id, lcl_user_email, lcl_user_password, lcl_user_address

	if lcl_user_id <> "" then
		iUserID       = lcl_user_id
		iUserEmail    = lcl_user_email
		'iUserPassword = lcl_user_password
	end if
end if

'Check for org features
lcl_orghasfeature_subscriptions = orghasfeature(iorgid,"subscriptions")
lcl_orghasfeature_bid_postings  = orghasfeature(iorgid,"bid_postings")
lcl_orghasfeature_job_postings  = orghasfeature(iorgid,"job_postings")
lcl_orghasfeature_subscriptions_distributionlist_showdesc = orghasfeature(iorgid,"subscriptions_distributionlist_showdesc")

'Check for org "edit displays"
lcl_orghasdisplay_subscriptions_intro_paragraph_public = orghasdisplay(iorgid,"subscriptions_intro_paragraph_public")

'Check for a screen message
lcl_onload  = ""

if lcl_success <> "" then
	lcl_msg    = setupScreenMsg(lcl_success)
	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
end if

'If a bad password or a record doesn't exist for the email entered then retrieve the values originally entered
if lcl_success = "EXISTS_BAD_PWD" OR lcl_success = "NOT_EXISTS" OR lcl_success = "MANAGE" OR lcl_success = "BAD_PWD" then
	iUserEmail    = session("useremail")
	'iUserPassword = session("userpassword")
else
	session("useremail")    = ""
	'session("userpassword") = ""
end If
session("userpassword") = ""


%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName%> - Subscriptions</title>

<!-- This metadata is for setting the priority and importance for CDO mail messages -->
<!--  
METADATA  
TYPE="typelib"  
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
NAME="CDO for Windows 2000 Library"  
--> 

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	<!--

		function getPassword() 
		{
			window.location.href='../forgot_password.asp';
		}

		function openWin2(url, name) 
		{
			popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		function findme() 
		{
			document.register.doaction.value = "find";
			validate();
		}

function validate() {
  var msg              = "";
  var lcl_return_false = "N";

		if(document.register.userpassword.value == "") {
     //msg+="The password cannot be blank.\n";
     inlineMsg(document.getElementById("userpassword").id,'<strong>Required Field Missing: </strong> Password',10,'userpassword');
     lcl_return_false = "Y";
		}

		if(document.getElementById("useremail").value == "") {
     inlineMsg(document.getElementById("useremail").id,'<strong>Required Field Missing: </strong> Email',10,'useremail');
     lcl_return_false = "Y";
  }else{
     var lcl_useremail = document.register.useremail.value.toLowerCase();
     var lcl_useremail = lcl_useremail.replace(' ', '');

  			//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us|info))$/;
			var rege = /.+@.+\..+/i;
		  	var Ok   = rege.test(lcl_useremail);

  			if (! Ok) {
         //msg+="The email must be in a valid format.\n";
         inlineMsg(document.getElementById("useremail").id,'<strong>Invalid Value: </strong> The email must be in a valid format.',10,'useremail');
         lcl_return_false = "Y";
  			}
		}

		if (document.register.subjecttext.value != '')	{
  				//msg+="Please remove any input from the Internal Only field at the bottom of the page.\n";
      inlineMsg(document.getElementById("problemtextinput").id,'<strong>Invalid Value: </strong> This field must be blank.  Please remove any input from this field.',10,'problemtextinput');
      lcl_return_false = "Y";
		}

  if(lcl_return_false == "Y") {
     return false;
  } else {
     document.register.submit();
  }

//			if(msg != "")
//			{
//				msg="Your form could not be submitted for the following reasons.\n\n" + msg;
//				alert(msg);
//			}
//			else 
//			{	
//				if (validateForm('register')) { document.register.submit(); }
//			}
		}

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

		function autoTab(input,len, e) 
		{
			var keyCode = (isNN) ? e.which : e.keyCode; 
			var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

			if(input.value.length >= len && !containsElement(filter,keyCode)) {
				input.value = input.value.slice(0, len);
			var addNdx = 1;

			while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
			{
				addNdx++;
				//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
			}

			input.form[(getIndex(input)+addNdx) % input.form.length].focus();
		}

		function containsElement(arr, ele) 
		{
			var found = false, index = 0;

			while(!found && index < arr.length)
				if(arr[index] == ele)
					found = true;
				else
					index++;
			return found;
		}

		function getIndex(input) 
		{
			var index = -1, i = 0, found = false;

			while (i < input.form.length && index == -1)
				if (input.form[i] == input)index = i;
				else i++;
					return index;
		}
			return true;
		}

function expand_collapse(p_value) {
//  whichEl = eval(p_value);
//  whichIm = event.srcElement;
  whichEl = document.getElementById(p_value);
  whichIm = document.getElementById('IMG_' + p_value);

  if (whichEl.style.display == "none") {
      whichEl.style.display = "block";
      whichIm.src           = "../images/collapse.jpg";
  }else{
      whichEl.style.display = "none";
      whichIm.src           = "../images/expand.jpg";
  }
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "&nbsp;";
}
	//-->
	</script>

<style type="text/css">
  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
     text-align:  right;
  }
</style>
</head>
<!--#Include file="../include_top.asp"-->
<%
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
                            RegisteredUserDisplay  "../"
  response.write "          <div id=""content"">" & vbcrlf
  response.write "            <div id=""centercontent"">" & vbcrlf
  response.write "              <div id=""screenMsg"">&nbsp;</div>" & vbcrlf

  if iConfirmID <> "" then

     response.write "<form name=""register"" action=""subscribe_action.asp"" method=""post"" autocomplete=""off"">" & vbcrlf
     response.write "  <input type=""hidden"" name=""subscription_confirmid"" id=""subscription_confirmid"" value=""" & iConfirmID & """ />" & vbcrlf
     response.write "<table id=""subscribe"" border=""0"" cellspacing=""0"">" & vbcrlf
                       displayMailLists False, iUserID

     response.write "  <tr>" & vbcrlf
     response.write "      <td colspan=""2"" align=""center"">" & vbcrlf
     response.write "          <input type=""submit"" name=""doaction"" id=""unsubscribeButton"" value=""UNSUBSCRIBE"" class=""actionbtn"" />" & vbcrlf
     response.write "          <input type=""submit"" name=""doaction"" id=""manageButton"" value=""Manage Subscriptions"" class=""actionbtn"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "</form>" & vbcrlf

  else
    'Determine if there is an "edit display" for the org.  If "yes" then use the value entered.
    'Otherwise, display the text that exists today.
     lcl_display_introtext = ""
     lcl_display_introtext = lcl_display_introtext & " Subscribe to our email lists to receive newsletters delivered directly to your inbox. "
     lcl_display_introtext = lcl_display_introtext & " To sign-up, type in your email address, password and in the Subscriptions list below check "
     lcl_display_introtext = lcl_display_introtext & " the box(es) of the emails you would like to receive and then click Subscribe. "
     lcl_display_introtext = lcl_display_introtext & " <strong>You will receive an email confirmation of your subscription(s) along with instructions on how to unsubscribe or manage your subscriptions.</strong>"

     if lcl_orghasdisplay_subscriptions_intro_paragraph_public then
        lcl_displayid_subscriptions_intro_paragraph_public = GetDisplayId("subscriptions_intro_paragraph_public")
        lcl_display_introtext = getOrgDisplayWithID(iOrgId, lcl_displayid_subscriptions_intro_paragraph_public, False)
     end if

    'Intro Text
     response.write "<p>" & vbcrlf
     response.write   lcl_display_introtext
     response.write "</p>" & vbcrlf
     'response.write "<p>" & vbcrlf
     'response.write "   Subscribe to our email lists to receive newsletters delivered directly to your inbox." & vbcrlf
     'response.write "   To sign-up, type in your email address, password and in the Subscriptions list below check" & vbcrlf
     'response.write "   the box(es) of the emails you would like to receive and then click Subscribe." & vbcrlf
     'response.write "   <strong>You will receive an email confirmation of your subscription(s) along with instructions on how to unsubscribe or manage your subscriptions.</strong>" & vbcrlf
     'response.write "   <strong>You will receive an email asking you to confirm your subscription, please follow the instructions provided.</strong>" & vbcrlf
     'response.write "</p>" & vbcrlf

     response.write "<p>" & vbcrlf
     response.write "<div class=""box_header4"">" & sOrgName & " Subscriptions</div>" & vbcrlf
     response.write "  <div class=""groupsmall2"">" & vbcrlf
     response.write "<form name=""register"" id=""register"" action=""subscribe_action.asp"" method=""post"" autocomplete=""off"">" & vbcrlf
     response.write "  <input type=""hidden"" name=""columnnameid"" id=""columnnameid"" value=""userid"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""userregistered"" id=""userregistered"" value=""1"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""residenttype"" id=""residenttype"" value=""N"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""headofhousehold"" id=""headofhousehold"" value=""1"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""doaction"" id=""doaction"" value=""subscribe"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""listtype"" id=""listtype"" value=""" & lcl_list_type & """ size=""5"" maxlength=""5"" />" & vbcrlf

     response.write "<table id=""subscribe"" border=""0"" cellspacing=""0"" class=""respTable gutterwidth"">" & vbcrlf

    	if sMsg <> "" then
        response.write "  <tr>" & vbcrlf
        response.write "      <td nowrap=""nowrap"" colspan=""2"">" & vbcrlf
        response.write "          <font style=""background-color:red; color:white;padding:.6em;"">" & sMsg & "</font>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""label"" align=""right"">" & vbcrlf
     response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><font color=""red"">*</font></span> " & vbcrlf
     response.write "          Email:" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <input type=""text"" name=""useremail"" id=""useremail"" value=""" & iUserEmail & """ style=""width:300px;"" maxlength=""100"" onchange=""clearMsg('useremail')"" placeholder=""Email Address"" autocomplete=""off"" />" & vbcrlf
     response.write "          <font color=""red""></font>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""label"" align=""right"">" & vbcrlf
     response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
     response.write "          <span class=""cot-text-emphasized""><font color=""red"">*</font></span> " & vbcrlf
     response.write "          Password:" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <input type=""password"" name=""userpassword"" id=""userpassword"" value=""" & twfpass & """ style=""width:300px;"" maxlength=""50"" onchange=""clearMsg('userpassword')"" placeholder=""Password"" autocomplete=""off"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td colspan=""2"" align=""center"">" & vbcrlf
     response.write "          <p>" & vbcrlf
     response.write "            <input class=""actionbtn"" type=""button"" name=""find"" value=""Find your subscriptions"" onclick=""findme();"" />" & vbcrlf
     response.write "            <input class=""actionbtn"" type=""button"" name=""forgot"" value=""Forgot your password"" onclick=""getPassword();"" />" & vbcrlf
     response.write "          </p>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
                       displayMailLists True, iUserID
     response.write "  <tr>" & vbcrlf
     response.write "      <td>&nbsp;</td>" & vbcrlf
     response.write "      <td align=""center"">" & vbcrlf
     response.write "          <input class=""actionbtn"" type=""button"" name=""subscribe"" value=""Subscribe"" onClick=""validate();"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

    'BEGIN: Problem field -----------------------------------------------------
     response.write "<div id=""problemtextfield1"">" & vbcrlf
     response.write "<p>" & vbcrlf
     response.write "  Internal Use Only, Leave Blank: <input type=""text"" name=""subjecttext"" id=""problemtextinput"" value="""" size=""6"" onchange=""clearMsg('problemtextinput')"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""problemorg"" value=""" & iorgid & """ /><br />" & vbcrlf
     response.write "  <strong>Please leave this field blank and remove any values that have been populated for it.</strong>" & vbcrlf
     response.write "</p>" & vbcrlf
     response.write "</div>" & vbcrlf
    'END: Problem field -------------------------------------------------------

     response.write "</form>" & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "<p>&nbsp;</p>" & vbcrlf

    'BEGIN: Inline javascripts ------------------------------------------------
     response.write "<script language=""javascript"">" & vbcrlf

    'Collapse all sections
     if lcl_orghasfeature_subscriptions AND check_for_jobbid_categories("",iorgid) = "Y" then
        response.write "document.getElementById('IMG_LIST').click();" & vbcrlf
     end if

     if lcl_orghasfeature_bid_postings AND check_for_jobbid_categories("BID",iorgid) = "Y" then
        response.write "document.getElementById('IMG_BID').click();" & vbcrlf
     end if

     if lcl_orghasfeature_job_postings AND check_for_jobbid_categories("JOB",iorgid) = "Y" then
        response.write "document.getElementById('IMG_JOB').click();" & vbcrlf
     end if

    'Expand the section that has been selected
     if lcl_list_type = "JOB" AND check_for_jobbid_categories("JOB",iorgid) = "Y" then
        response.write "document.getElementById('IMG_JOB').click();" & vbcrlf
     elseif lcl_list_type = "BID" AND check_for_jobbid_categories("BID",iorgid) = "Y" then
        response.write "document.getElementById('IMG_BID').click();" & vbcrlf
     else
        if check_for_jobbid_categories("",iorgid) = "Y" then
           response.write "document.getElementById('IMG_LIST').click();" & vbcrlf
        end if
     end if

    'Check for any "onload" scripts.
     if lcl_onload <> "" then
        response.write lcl_onload & vbcrlf
     end if

     response.write "</script>" & vbcrlf
    'END: Inline javascripts --------------------------------------------------

  end if
%>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub displayMailLists(iCanEdit, iUserID)
	Dim sSql, oList

'Retrieve all of the categories
 sSQL = "SELECT distributionlistid, "
 sSQL = sSQL & " distributionlistname, "
 sSQL = sSQL & " distributionlistdescription, "
 sSQL = sSQL & " distributionlistdisplay, "
 sSQL = sSQL & " orgid, "
 sSQL = sSQL & " parentid, "
 sSQL = sSQL & " isnull(distributionlisttype, '') as distributionlisttype "
 sSQL = sSQL & " FROM egov_class_distributionlist "
 sSQL = sSQL & " WHERE orgid = '" & iorgid & "' "
 sSQL = sSQL & " AND distributionlistdisplay = 1 "
 sSQL = sSQL & " AND parentid is null "
 sSQL = sSQL & " ORDER BY distributionlisttype, UPPER(distributionlistname) "

	set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

 lcl_listtype_prev = "LIST"
 lcl_line_count    = 0

	if not oList.eof then
  		response.write "<tr>" & vbcrlf
    response.write "    <td colspan=""2"">" & vbcrlf
		  response.write "        <fieldset>" & vbcrlf
    response.write "          <legend><strong>Subscriptions&nbsp;</strong></legend>" & vbcrlf

    if iCanEdit then
     		response.write "          <p>Check the email list to which you would like to subscribe.</p>" & vbcrlf
		     response.write "          <p>Uncheck to unsubscribe from an email list.</p>" & vbcrlf
    else
       response.write "          <p>The following is a list of all of the subscriptions you are subscribed to.</p>" & vbcrlf
       response.write "          <p>" & vbcrlf
       response.write "             Click the ""UNSUBSCRIBE"" button if you wish to unsubscribe from ALL of the subscriptions.<br />" & vbcrlf
       response.write "             Click the ""Manage Subscriptions"" button if you would like to add or change any subscriptions." & vbcrlf
       response.write "          </p>" & vbcrlf
    end if

  		while not oList.eof
       lcl_line_count = lcl_line_count + 1

       if oList("distributionlisttype") = "JOB" then
          lcl_listtype   = oList("distributionlisttype")
          lcl_list_title = getFeatureName("job_postings")
          'lcl_list_title = "JOB POSTINGS"
       elseif oList("distributionlisttype") = "BID" then
          lcl_listtype   = oList("distributionlisttype")
          lcl_list_title = getFeatureName("bid_postings")
          'lcl_list_title = "BID POSTINGS"
       else
          lcl_listtype   = "LIST"
          lcl_list_title = "DISTRIBUTION LISTS"
       end if

      'Determine which features the org has "turned-on"
       if iCanEdit then
          if lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions then
             lcl_show = "Y"
          elseif lcl_listtype = "JOB" AND lcl_orghasfeature_job_postings then
             lcl_show = "Y"
          elseif lcl_listtype = "BID" AND lcl_orghasfeature_bid_postings then
             lcl_show = "Y"
          else
             lcl_show = "N"
          end if
       else
          lcl_show = "Y"
       end if

      'If the current parent (category) is different then the previous record then reset the variables
      'if (isnull(lcl_listtype_prev) OR lcl_listtype_prev <> oList("distributionlisttype")) AND lcl_line_count > 1 then
       if lcl_line_count > 1 then
          if isnull(lcl_listtype_prev) then
             lcl_listtype_prev = "LIST"
          end if

          'if lcl_listtype_prev <> oList("distributionlisttype") then
          if lcl_listtype_prev <> lcl_listtype then
             lcl_line_count = 1
          end if
       end if

       if lcl_line_count = 1 then
          if lcl_list_title <> "" then
             if lcl_listtype <> "LIST" then
                response.write "</div>" & vbcrlf
                response.write "<hr size=""1"" width=""100%"">" & vbcrlf
             end if

             if lcl_show = "Y" then
                response.write "<p>" & vbcrlf

                if iCanEdit then
                   response.write "<img src=""../images/expand.jpg"" id=""IMG_" & lcl_listtype & """ width=""9"" height=""9"" style=""cursor: hand"" onclick=""expand_collapse('" & lcl_listtype & "')"" alt=""Click to expand/collapse"">&nbsp;" & vbcrlf
                end if

                response.write "<strong>" & UCASE(lcl_list_title) & "</strong>" & vbcrlf
                response.write "<div id=""" & lcl_listtype & """>" & vbcrlf
             end if
          end if
       end if

       lcl_listtype_prev = lcl_listtype

       if lcl_show = "Y" then

         'If this is a listtype of JOB/BID then check for a description and display it ONLY if one exists.
          lcl_desc = ""

          if lcl_listtype <> "LIST" OR (lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions_distributionlist_showdesc) then
             if oList("distributionlistdescription") <> "" then
                lcl_desc = " <i>(" & oList("distributionlistdescription") & ")</i>"
             else
                lcl_desc = ""
             end if
          else
             lcl_desc = ""
          end if

         'If the list is maintainable then show the checkbox and determine if it is checked.
          lcl_show_distributionlistname = "N"

          if iCanEdit then
             lcl_show_distributionlistname = "Y"

           		if IsMember( iUserID, oList("distributionlistid") ) then
            				lcl_checked_list1 = " checked=""checked"""
             else
                lcl_checked_list1 = ""
          			end if

   		    	   response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
             response.write "<input name=""maillist"" type=""checkbox"" value=""" &  oList("distributionlistid") & """" & lcl_checked_list1 & " />&nbsp;" & vbcrlf
          else
     	      	if IsMember( iUserID, oList("distributionlistid") ) then
                lcl_show_distributionlistname = "Y"
      		    	   response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
             end if
          end if

         'Determine if the Distribution List Name is to be displayed.
         'It was originally designed to always show, but in the case that this list is NOT maintainable then
         'depending on if the user is assigned to that subscription(s) will determine if we show/hide the value.
          if lcl_show_distributionlistname = "Y" then
             response.write "<strong>" & oList("distributionlistname") & "</strong>" & lcl_desc & "<br />" & vbcrlf
          end if

         'Check for any sub-categories
         	sSQL2 = "SELECT * "
     	    sSQL2 = sSQL2 & " FROM egov_class_distributionlist "
        	 sSQL2 = sSQL2 & " WHERE orgid = " & iorgid
        	 sSQL2 = sSQL2 & " AND distributionlistdisplay = 1 "
  	       sSQL2 = sSQL2 & " AND UPPER(distributionlisttype) = '" & UCASE(oList("distributionlisttype")) & "' "
          sSQL2 = sSQL2 & " AND parentid = " & oList("distributionlistid")
          sSQL2 = sSQL2 & " ORDER BY UPPER(distributionlistname) "

        		set rs2 = Server.CreateObject("ADODB.Recordset")
        		rs2.Open sSQL2, Application("DSN"), adOpenForwardOnly, adLockReadOnly

          if not rs2.eof then
           		while not rs2.eof
         		    	response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

               'If this is a listtype of JOB/BID then check for a description and display it ONLY if one exists.
                lcl_desc = ""

    						      if lcl_listtype <> "LIST" OR (lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions_distributionlist_showdesc) then
                   if oList("distributionlistdescription") <> "" then
                      lcl_desc = " <i>(" & oList("distributionlistdescription") & ")</i>"
                   else
                      lcl_desc = ""
                   end if
                else
                   lcl_desc = ""
                end if

               'If the list is maintainable then show the checkbox and determine if it is checked.
                lcl_show_distributionlistname = "N"

                if iCanEdit then
                   lcl_show_distributionlistname = "Y"

           	      	if IsMember( iUserID, rs2("distributionlistid") ) then
                  				lcl_checked_list2 = " checked=""checked"""
                   else
                      lcl_checked_list2 = ""
                			end if

                   response.write "<input name=""maillist"" type=""checkbox"" value=""" & rs2("distributionlistid") & """" & lcl_checked_list2 & " />&nbsp;" & vbcrlf
                else
           	      	if IsMember( iUserID, rs2("distributionlistid") ) then
                      lcl_show_distributionlistname = "Y"
                   end if
                end if

               'Determine if the Distribution List Name is to be displayed.
               'It was originally designed to always show, but in the case that this list is NOT maintainable then
               'depending on if the user is assigned to that subscription(s) will determine if we show/hide the value.
                if lcl_show_distributionlistname = "Y" then
                   response.write "<strong>" & rs2("distributionlistname") & "</strong>" & lcl_desc & "<br />" & vbcrlf
                end if

             			rs2.movenext
           		wend
          end if

          rs2.close
          set rs2 = nothing
       end if

       lcl_listtype_prev = lcl_listtype

    			oList.MoveNext
  		wend

   	response.write "          <p>" & vbcrlf
    response.write "        </fieldset>" & vbcrlf
    response.write "        <p>" & vbcrlf
  		response.write "    </td>" & vbcrlf
    response.write "</tr>" & vbcrlf
	end if

	oList.close 
	set oList = nothing 

end sub

'------------------------------------------------------------------------------
'function getCategoryName(p_dlid)
'  sSQLc = "SELECT distributionlistname, distributionlistdescription "
'  sSQLc = sSQLc & " FROM egov_class_distributionlist "
'  sSQLc = sSQLc & " WHERE distributionlistid = " & p_dlid

'  set rsc = Server.CreateObject("ADODB.Recordset")
'  rsc.Open sSQLc, Application("DSN"), 0, 1

'  if not rsc.eof then
'     lcl_return = rsc("distributionlistname")

'     if rsc("distributionlistdescription") <> "" then
'        lcl_return = lcl_return & "&nbsp;-&nbsp;" & rsc("distributionlistdescription")
'     end if
'  else
'     lcl_return = ""
'  end if

'  getCategoryName = lcl_return

'end function
%>

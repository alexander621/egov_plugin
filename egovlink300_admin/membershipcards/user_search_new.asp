<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: user_search.asp
' AUTHOR: David Boyer
' CREATED: 07/28/2011
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the user search used to create/reprint a pool pass "membership" 
'               cards (by egov_users.userid) and was initially set up for Menlo Park.
'
' MODIFICATION HISTORY
' 1.0 07/28/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Check to see if the feature is offline
 if isFeatureOffline("registration") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 if not userhaspermission(session("userid"),"create_user_membershipcards_new") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the search options
 lcl_sc_username                 = ""
 lcl_sc_userid                   = ""
 lcl_sc_familyid                 = ""
 lcl_sc_showcardsnotprinted      = ""
 lcl_showcardsnotprinted_checked = ""
 lcl_session_redirect_url        = ""

 if request("sc_username") <> "" then
    lcl_sc_username      = formatUserNameForPage(request("sc_username"))
    lcl_username_for_url = formatUserNameForURL(lcl_sc_username)
 end if

 if request("sc_userid") <> "" then
    lcl_sc_userid = request("sc_userid")
    lcl_sc_userid = clng(lcl_sc_userid)
 end if

 if request("sc_familyid") <> "" then
    lcl_sc_familyid = request("sc_familyid")
    lcl_sc_familyid = clng(lcl_sc_familyid)

   'If the user is searching on a familyid then we need to clear out the other search criteria
    lcl_sc_username = ""
    lcl_sc_userid   = ""
 end if

 if request("sc_showcardsnotprinted") = "Y" then
    lcl_sc_showcardsnotprinted      = request("sc_showcardsnotprinted")
    lcl_showcardsnotprinted_checked = " checked=""checked"""
 end if

'Set up the return Session variables for the page
 session("RedirectLang") = "Return to Member List"

 lcl_session_redirect_url = lcl_session_redirect_url & "../membershipcards/user_search_new.asp"
 lcl_session_redirect_url = lcl_session_redirect_url & "?sc_username="            & lcl_username_for_url
 lcl_session_redirect_url = lcl_session_redirect_url & "&sc_userid="              & lcl_sc_userid
 lcl_session_redirect_url = lcl_session_redirect_url & "&sc_showcardsnotprinted=" & lcl_sc_showcardsnotprinted

 if lcl_sc_familyid <> "" then
    lcl_session_redirect_url = lcl_session_redirect_url & "&sc_familyid=" & lcl_sc_familyid
 end if

 session("RedirectPage") = lcl_session_redirect_url

'Set the membershipid to the one for pools
' Set oMembership = New classMembership
' oMembership.SetMembershipId( "pool" )

' subTotals  = Request("subTotals")
' showDetail = Request("showDetail")
 'oMemberlist = Request("username")

' session("STATUS") = ""

'Check for org features
' lcl_orghasfeature_pool_attendance_view             = orghasfeature("pool_attendance_view")
' lcl_orghasfeature_customreports_membership_scanlog = orghasfeature("customreports_membership_scanlog")

'Check for user features
' lcl_userhaspermission_customreports_membership_scanlog = userhaspermission(session("userid"),"customreports_membership_scanlog")

'Determine if the user has performed a search
 lcl_init = true

 if request.ServerVariables("REQUEST_METHOD") = "POST" OR lcl_sc_familyid <> "" then
    lcl_init = false
 end if

'Setup BODY onload
 lcl_onload = "document.getElementById('sc_username').focus();"
%>
<html>
<head>
  <title>E-Gov Administration Console {Search, Create, Reprint Member IDs}</title>
  
	 <link rel="stylesheet" type="text/css" href="../global.css">
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
  <!--
//  function checkStat() {
//  if ( !(form1.statusInProgress.checked) &&  !(form1.statusPending.checked) && !(form1.statusRefund.checked) && !(form1.statusDenied.checked) &&  !(form1.statusCompleted.checked) && !(form1.statusProcessed.checked)) {
//		alert("You must select the status.");
//		form1.statusPending.focus();
//		return false;
//	}
//  }
function CheckAllStatus() {
		if (document.form1.CheckAllStat.checked) {
			document.form1.statusPending.checked   = true;
			document.form1.statusCompleted.checked = true;
			document.form1.statusDenied.checked    = true;
		} else {
			document.form1.statusPending.checked   = false;
			document.form1.statusCompleted.checked = false;
			document.form1.statusDenied.checked    = false;
		}
}

function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1&updateform=searchform&updatefield=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function Open_Profile( sUserId, sFamilyID ) {
  location.href='../dirs/manage_family_member.asp?u=' + sUserId + "&iReturn=" + sFamilyID;
}

function goToCard(sUserID,sAction) {
  var lcl_card_url = '';

  if(sAction=="CREATE") {
     lcl_card_url = 'user_takepic_new.asp';
  }else if (sAction=="REPRINT" || sAction=="PRINT") {
     lcl_card_url = 'user_displaycard_new.asp';
  }

  location.href = lcl_card_url + '?userid=' + sUserID + '&action=' + sAction;
}

function openCustomReports(p_report,p_memberid,p_rateid) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  var lcl_additional_parameters;

  if(p_memberid != "") {
     lcl_additional_parameters = "&memberid=" + p_memberid;
  }

  if(p_memberid != "") {
     if(lcl_additional_parameters != "") {
        lcl_additional_parameters = lcl_additional_parameters + "&rateid=" + p_rateid;
     }else{
        lcl_additional_parameters = "&rateid=" + p_rateid;
     }
  }

  eval('window.open("../customreports/customreports.asp?cr='+p_report+lcl_additional_parameters+'", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function validateFields() {

//		var daterege   = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
//		var dateFromOk = daterege.test(document.getElementById("sc_from_date").value);

//		if (! dateToOk ) {
//      document.getElementById("sc_to_date").focus();
//      inlineMsg(document.getElementById("sc_to_date_cal").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'sc_to_date_cal');
//      return false;
//  }else{
//      clearMsg("sc_to_date_cal");
//  }

  if(document.getElementById('sc_username').value == '' && document.getElementById('sc_userid').value == '' && '<%=lcl_sc_familyid%>' == '') {
     inlineMsg(document.getElementById('sc_userid').id,'<strong>Required Field Missing: </strong>At least one search option must be entered.',10,'sc_userid');
     inlineMsg(document.getElementById('sc_username').id,'<strong>Required Field Missing: </strong>At least one search option must be entered.',10,'sc_username');
     document.getElementById('sc_username').focus();
     return false;
  } else {
     clearMsg('sc_userid');
     clearMsg('sc_username');

     return true;
  }
}

  //-->
  </script>

<style type="text/css">
  .fieldset legend {
     margin-bottom: 5px;
  }
</style>

</head>

<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""searchform"" id=""searchform"" action=""user_search_new.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""qry"" value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  'response.write "      <td><font size=""+1""><strong>" & session("sOrgName") & "&nbsp;" & oMembership.GetMembershipName() & ":&nbsp;Search, Create, Reprint Member IDs</strong></font></td>
  response.write "      <td><font size=""+1""><strong>" & session("sOrgName") & ":&nbsp;Search, Create, Reprint Member IDs</strong></font></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'BEGIN: Search Options -------------------------------------------------------
  response.write "  <tr class=""noprint"">" & vbcrlf
  response.write "     <td class=""noprint"">" & vbcrlf
  response.write "        	<fieldset class=""fieldset"">" & vbcrlf
  response.write "         		<legend><strong>Search Options</strong>&nbsp;</legend>" & vbcrlf
  response.write "        			<table border=""0"" cellpadding=""2"" cellspacing=""0"" id=""searchtable"" style=""width:600px;"">" & vbcrlf
  response.write "             <tr>" & vbcrlf
  response.write "                 <td>Name:</td>" & vbcrlf
  response.write "                 <td><input type=""text"" name=""sc_username"" id=""sc_username"" value=""" & lcl_sc_username & """ size=""30"" maxlength=""50"" onchange=""clearMsg('sc_username');"" /></td>" & vbcrlf
  response.write "                 <td><input type=""checkbox"" name=""sc_showcardsnotprinted"" id=""sc_showcardsnotprinted"" value=""Y""" & lcl_showcardsnotprinted_checked & " />&nbsp;Show ONLY Cards Not Printed</td>" & vbcrlf
  response.write "             </tr>" & vbcrlf
  response.write "             <tr>" & vbcrlf
  response.write "                 <td>User ID:</td>" & vbcrlf
  response.write "                 <td colspan=""2""><input type=""text"" name=""sc_userid"" id=""sc_userid"" value=""" & lcl_sc_userid & """ size=""10"" maxlength=""10"" onchange=""clearMsg('sc_userid');"" /></td>" & vbcrlf
  response.write "             </tr>" & vbcrlf

  if lcl_sc_familyid <> "" then
     response.write "<tr>" & vbcrlf
     response.write "    <td colspan=""3"">Searching on Family ID: [<span style=""color: #ff0000"">" & lcl_sc_familyid & "</span>]</td>" & vbcrlf
     response.write "</tr>" & vbcrlf
  end if

  response.write "             <tr>" & vbcrlf
  response.write "                 <td colspan=""3""><input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" onclick=""return validateFields()"" /></td>" & vbcrlf
  response.write "             </tr>" & vbcrlf
  response.write "        			</table>" & vbcrlf
  response.write "		      	</fieldset>" & vbcrlf
  response.write "     </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: Search Options ---------------------------------------------------------

 'BEGIN: Member List ----------------------------------------------------------
if not lcl_init then
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
                            showUserList session("orgid"), lcl_sc_username, lcl_sc_userid, _
                                         lcl_sc_familyid, lcl_sc_showcardsnotprinted
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
end if
 'END: Member List ------------------------------------------------------------

  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'Set oMembership = Nothing

'------------------------------------------------------------------------------
sub showUserList(iOrgID, iSC_username, iSC_userid, iSC_familyid, iSC_showCardsNotPrinted)

  sSC_username            = ""
  sSC_userid              = ""
  sSC_familyid            = ""
  sSC_showCardsNotPrinted = false

  lcl_orgid               = 0
  lcl_userid              = 0
  lcl_userfname           = ""
  lcl_userlname           = ""
  lcl_userstreetnumber    = ""
  lcl_userstreetprefix    = ""
  lcl_useraddress         = ""
  lcl_userunit            = ""
  lcl_familyid            = 0
  lcl_birthdate           = ""
  lcl_card_pic_uploaded   = 0
  lcl_card_printed_count  = 0
  lcl_relationship        = ""
  lcl_display_username    = ""
  lcl_display_address     = ""
  lcl_display_age         = ""
  lcl_photo_id_text       = ""
  lcl_photo_id_url        = ""
  lcl_familyid_url        = ""
  lcl_bgcolor             = "#ffffff"
  lcl_linecount           = 0

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

  if iSC_username <> "" then
     sSC_username = iSC_username
     sSC_username = ucase(sSC_username)
     sSC_username = dbsafe(sSC_username)
     sSC_username = "'%" & sSC_username & "%'"
  end if

  if iSC_userid <> "" then
     sSC_userid = iSC_userid
     sSC_userid = clng(sSC_userid)
  end if

  if iSC_familyid <> "" then
     sSC_familyid = iSC_familyid
     sSC_familyid = clng(sSC_familyid)
  end if

  if iSC_showCardsNotPrinted = "Y" then
     sSC_showCardsNotPrinted = true
  end if

  sSQL = "SELECT u.userid, "
  sSQL = sSQL & " u.userfname, "
  sSQL = sSQL & " u.userlname, "
  sSQL = sSQL & " u.userstreetnumber, "
  sSQL = sSQL & " u.userstreetprefix, "
  sSQL = sSQL & " u.useraddress, "
  sSQL = sSQL & " u.userunit, "
  sSQL = sSQL & " u.familyid, "
  sSQL = sSQL & " u.birthdate, "
  sSQL = sSQL & " u.card_pic_uploaded, "
  sSQL = sSQL & " u.card_printed_count, "
  sSQL = sSQL & " fm.relationship "
  sSQL = sSQL & " FROM egov_users u "
  sSQL = sSQL &      " INNER JOIN egov_familymembers fm "
  sSQL = sSQL &                 " ON fm.userid = u.userid "
  sSQL = sSQL &                 " AND fm.isdeleted = 0 "
  sSQL = sSQL & " WHERE orgid = " & lcl_orgid
  sSQL = sSQL & " AND u.isdeleted = 0 "
  sSQL = sSQL & " AND ((u.userfname IS NOT NULL OR u.userfname <> '') "
  sSQL = sSQL & " OR   (u.userlname IS NOT NULL OR u.userlname <> '')) "

 'BEGIN: Check for search criteria --------------------------------------------
 'If the family id has been clicked on then query on the family id value.
 'Otherwise, query on the first/last name
  if sSC_familyid <> "" then
     sSQL = sSQL & " AND  u.familyid = " & sSC_familyid
  else
     if sSC_username <> "" then
        sSQL = sSQL & " AND ((upper(u.userlname) like (" & sSC_username & ")) "
        sSQL = sSQL & " OR   (upper(u.userfname) like (" & sSC_username & "))) "
     end if

     if sSC_userid <> "" then
        sSC_userid = "'%" & sSC_userid & "%'"

        sSQL = sSQL & " AND CAST(u.userid as varchar) LIKE (" & sSC_userid & ") "
     end if
  end if

  if sSC_showCardsNotPrinted then
     sSQL = sSQL & " AND u.card_pic_uploaded = 1 "
     sSQL = sSQL & " AND u.card_printed_count = 0 "
  end if
 'END: Check for search criteria ----------------------------------------------

  sSQL = sSQL & " ORDER BY upper(userlname), upper(userfname) "

  set oBuildUserList = Server.CreateObject("ADODB.Recordset")
  oBuildUserList.Open sSQL, Application("DSN"), 3, 1

  if not oBuildUserList.eof then
 		  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""100%"">" & vbcrlf
   		response.write "  <tr valign=""bottom"">" & vbcrlf
 		  response.write "      <th>User<br />ID</th>" & vbcrlf
 		  response.write "      <th align=""left"" nowrap=""nowrap"">Name</th>" & vbcrlf
 		  response.write "      <th>Age</th>" & vbcrlf
 		  response.write "      <th>Family<br />ID</th>" & vbcrlf
 		  response.write "      <th align=""left"" nowrap=""nowrap"">Relationship</th>" & vbcrlf
 		  response.write "      <th align=""left"" nowrap=""nowrap"">Home Address</th>" & vbcrlf
 		  response.write "      <th nowrap=""nowrap"">&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oBuildUserList.eof
        lcl_linecount          = lcl_linecount + 1
        lcl_bgcolor            = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_userid             = oBuildUserList("userid")
        lcl_userfname          = oBuildUserList("userfname")
        lcl_userlname          = oBuildUserList("userlname")
        lcl_userstreetnumber   = oBuildUserList("userstreetnumber")
        lcl_userstreetprefix   = oBuildUserList("userstreetprefix")
        lcl_useraddress        = oBuildUserList("useraddress")
        lcl_userunit           = oBuildUserList("userunit")
        lcl_familyid           = oBuildUserList("familyid")
        lcl_birthdate          = oBuildUserList("birthdate")
        lcl_card_pic_uploaded  = oBuildUserList("card_pic_uploaded")
        lcl_card_printed_count = oBuildUserList("card_printed_count")
        lcl_relationship       = oBuildUserList("relationship")
        lcl_display_username   = ""
        lcl_display_address    = ""
        lcl_display_age        = ""
        lcl_photo_id_text      = ""
        lcl_photo_id_url       = ""
        lcl_familyid_url       = ""

        if lcl_userlname <> "" then
           lcl_display_username = lcl_userlname
        end if

        if lcl_userfname <> "" then
           if lcl_display_username <> "" then
              lcl_display_username = lcl_display_username & ", " & lcl_userfname
           else
              lcl_display_username = lcl_userfname
           end if
        end if

        if lcl_birthdate <> "" then
           lcl_display_age = year(now) - year(lcl_birthdate)
        else
           lcl_display_age = "&nbsp;"
        end if

        if lcl_userstreetnumber <> "" then
           lcl_display_address = lcl_userstreetnumber
        end if

        if lcl_userstreetprefix <> "" then
           if lcl_display_address <> "" then
              lcl_display_address = lcl_display_address & " " & lcl_userstreetprefix
           else
              lcl_display_address = lcl_userstreetprefix
           end if
        end if

        if lcl_useraddress <> "" then
           if lcl_display_address <> "" then
              lcl_display_address = lcl_display_address & " " & lcl_useraddress
           else
              lcl_display_address = lcl_useraddress
           end if
        end if

        if lcl_card_pic_uploaded then
           if lcl_card_printed_count > 0 then
    	         lcl_photo_id_text = "Reprint"
           else
    	         lcl_photo_id_text = "Print"
           end if
     	  else
           lcl_photo_id_text = "Create"
     	  end if

        lcl_photo_id_url = "javascript:goToCard('" & lcl_userid & "','" & ucase(lcl_photo_id_text) & "');"

        lcl_familyid_url = lcl_familyid_url & "user_search_new.asp"
        lcl_familyid_url = lcl_familyid_url & "?sc_familyid="            & lcl_familyid
        lcl_familyid_url = lcl_familyid_url & "&sc_username="            & iSC_username
        lcl_familyid_url = lcl_familyid_url & "&sc_userid="              & iSC_userid
        lcl_familyid_url = lcl_familyid_url & "&sc_showcardsnotprinted=" & iSC_showcardsnotprinted

        response.write "  <tr bgcolor=""" &  lcl_bgcolor  & """ class=""tablelist"" valign=""top"">" & vbcrlf
     	  response.write "      <td align=""center"">" & lcl_userid & "</td>" & vbcrlf
     	  response.write "      <td>" & vbcrlf
        response.write "          <a href=""javascript:Open_Profile('" & lcl_userid & "','" & lcl_familyid & "');"" onMouseOver=""tooltip.show('Click to Edit User');"" onMouseOut=""tooltip.hide();"">" & lcl_display_username & "</a>" & vbcrlf
        response.write "      </td>" & vbcrlf
     	  response.write "      <td align=""center"">"  & lcl_display_age & "</td>" & vbcrlf
     	  response.write "      <td align=""center"">" & vbcrlf
        response.write "          <a href=""" & lcl_familyid_url & """ onMouseOver=""tooltip.show('Click to search on this Family ID');"" onMouseOut=""tooltip.hide();"">" & lcl_familyid & "</a></td>" & vbcrlf
        response.write "      </td>" & vbcrlf
     	  response.write "      <td>" & lcl_relationship    & "</td>" & vbcrlf
     	  response.write "      <td>" & lcl_display_address & "</td>" & vbcrlf
 	      response.write "      <td>" & vbcrlf
        response.write "          <input type=""button"" value=""" & lcl_photo_id_text & """ style=""cursor: hand"" onMouseOver=""tooltip.show('Click to <strong style=\'color:#ffff00\'>" & UCASE(lcl_photo_id_text) & "</strong> a Membership Card');"" onMouseOut=""tooltip.hide();"" onclick=""" & lcl_photo_id_url & """ />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oBuildUserList.movenext
	response.flush
     loop

     response.write "</table>" & vbcrlf
     response.write "<div align=""right""><strong>Total Users: </strong>[" & lcl_linecount & "]</div>" & vbcrlf

  end if

  oBuildUserList.close
  set oBuildUserList = nothing

end sub

'------------------------------------------------------------------------------
Function TranslateMember( sRelationship )
	If UCase(sRelationship) = "YOURSELF" Then
		TranslateMember = "Purchaser"
	Else 
		TranslateMember = sRelationship
	End If 
	
End Function 

'------------------------------------------------------------------------------
Function MakeProper( sString )
	If sString = "" Then
		MakeProper = ""
	Else
		MakeProper = UCase(Left(sString,1)) & LCase(Mid(sString,2))
	End If 
End Function 

'------------------------------------------------------------------------------
Function FormatPhone( Number )
	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If
End Function

'------------------------------------------------------------------------------
function formatUserNameForPage(iUserName)

  lcl_return = ""

  if iUserName <> "" then
     lcl_return = iUserName
     lcl_return = replace(lcl_return,"<<AMP>>","&")
     lcl_return = replace(lcl_return,"<<QUT>>","'")
     lcl_return = replace(lcl_return,"<<DBL>>","""")
  end if

  formatUserNameForPage = lcl_return

end function

'------------------------------------------------------------------------------
function formatUserNameForURL(iUserName)

  lcl_return = ""

  if iUserName <> "" then
     lcl_return = iUserName
     lcl_return = replace(lcl_return,"&","<<AMP>>")
     lcl_return = replace(lcl_return,"'","<<QUT>>")
     lcl_return = replace(lcl_return,"""","<<DBL>>")
  end if

  formatUserNameForURL = lcl_return

end function

'--------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = ""
  end if

  dbsafe = lcl_value

end function
%>

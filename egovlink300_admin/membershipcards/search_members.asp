<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="membership_card_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: search_members.asp
' AUTHOR: David Boyer
' CREATED: 07/17/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the Member Search used to create/reprint a pool pass membership card
'
' MODIFICATION HISTORY
' 1.0 07/17/07 David Boyer - Created code.
' 1.1 03/01/08 David Boyer - Added DEMO version
' 1.2 05/28/08 David Boyer - Added Next/Previous buttons and records returned per page
' 1.3 06/11/08 David Boyer - Removed "1 year" restriction from query and added search criteria.
' 1.4 09/18/08	David Boyer - Added Membership Renewals and now pulled StartDate and ExpirationDate from egov_poolpasspurchases
'                            instead of generating the expirationdate.
' 1.5 01/30/09 David Boyer - Now allow members that have purchased a "punchcard" type of rate to create a photo ID.
' 1.6 02/05/09 David Boyer - Added "MembershipCard_Scans_By_MemberID" custom report.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iMembershipId, oMembership, lcl_initial_search, lcl_search_name

sLevel = "../" ' Override of value from common.asp

'Retrieve the variable to determine if this is the initial search
 lcl_initial_search = request("qry")

'Determine if this is a demo
 if request("demo") = "Y" then
    lcl_demo       = request("demo")
    lcl_demo_title = " (DEMO)"
 else
    lcl_demo       = "N"
    lcl_demo_title = ""
 end if

if lcl_demo = "Y" then
   if not userhaspermission(session("userid"),"demo_create_membership_cards") then
  	   response.redirect sLevel & "permissiondenied.asp"
   end if
else
   if not userhaspermission(session("userid"),"create membership cards") then
     	response.redirect sLevel & "permissiondenied.asp"
   end if
end if

'Retrieve the variable to determine if the list should be pulled for the family id only
' if request.querystring("username") <> "" then
'    lcl_search_name = request.querystring("username")
' else
'    lcl_search_name = request.form("username")
' end if
 lcl_search_name     = ""
 lcl_family_id       = request("familyid")
 lcl_sc_show_expired = request("sc_show_expired")
 lcl_sc_from_date    = request("sc_from_date")
 lcl_sc_to_date      = request("sc_to_date")

 if request("username") <> "" then
    lcl_search_name         = formatUserNameForPage(request("username"))
    lcl_search_name_for_url = formatUserNameForURL(lcl_search_name)
 end if

 if lcl_sc_show_expired <> "Y" then
    lcl_showexpired_checked = ""
 else
    lcl_showexpired_checked = " CHECKED"
 end if

 if lcl_sc_from_date = "" OR isnull(lcl_sc_from_date) then
    lcl_sc_from_date = date()
 end if

 if lcl_sc_to_date = "" OR isnull(lcl_sc_to_date) then
    lcl_sc_to_date = DATEADD("yyyy",1,date())
 end if

'Set up the return Session variables for the page
 session("RedirectLang") = "Return to Member ID List"

 if lcl_family_id = "" then
    session("RedirectPage") = "../MembershipCards/search_members.asp?username=" & lcl_search_name_for_url & "&sc_show_expired=" & lcl_show_expired & "&sc_from_date=" & lcl_sc_from_date & "&sc_to_date=" & lcl_sc_to_date & "&demo=" & lcl_demo
 else
    session("RedirectPage") = "../MembershipCards/search_members.asp?familyid=" & lcl_family_id   & "&sc_show_expired=" & lcl_show_expired & "&sc_from_date=" & lcl_sc_from_date & "&sc_to_date=" & lcl_sc_to_date & "&demo=" & lcl_demo
 end if

'Set the membershipid to the one for pools
 Set oMembership = New classMembership
 oMembership.SetMembershipId( "pool" )

 subTotals  = Request("subTotals")
 showDetail = Request("showDetail")
 'oMemberlist = Request("username")

 session("STATUS") = ""

'Check for org features
 lcl_orghasfeature_pool_attendance_view             = orghasfeature("pool_attendance_view")
 lcl_orghasfeature_customreports_membership_scanlog = orghasfeature("customreports_membership_scanlog")
 lcl_orghasfeature_memberships_usekeycards          = orghasfeature("memberships_usekeycards")

'Check for user features
 lcl_userhaspermission_customreports_membership_scanlog = userhaspermission(session("userid"),"customreports_membership_scanlog")
%>
<html>
<head>
  <title>E-Gov Administration Console {Search, Create, Reprint Member IDs<%=lcl_demo_title%>}</title>
  
	 <link rel="stylesheet" type="text/css" href="../global.css">
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <!--	<link rel="stylesheet" type="text/css" href="poolpass.css"> -->
  <!-- <link rel="stylesheet" type="text/css" href="style_pool.css"> -->
  <!--	<link rel="stylesheet" type="text/css" media="print" href="receiptprint.css" /> -->

<style type="text/css">
.fieldset {
   border: 1pt solid #c0c0c0;
     border-radius: 5px !important; 
}

.fieldset legend {
   border: 1pt solid #c0c0c0;
     border-radius: 5px;
   padding: 4px 8px;
   font-size: 1.125em;
   font-weight: bold;
}

.barcodeTooltipTable th {
   border-bottom: 1pt solid #ffffff;
   text-align: left;
}

.isActiveBarcode {
   font-size: 1.25em;
   font-weight: bold;
   color: #ffff00;
}
</style>


  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<script type="text/javascript">
<!--
$(document).ready(function() {
   $('#username').focus();
});

function CheckAllStatus() {
		if (document.form1.CheckAllStat.checked) {
			document.form1.statusPending.checked = true;
			document.form1.statusCompleted.checked = true;
			document.form1.statusDenied.checked = true;
		} else {
			document.form1.statusPending.checked = false;
			document.form1.statusCompleted.checked = false;
			document.form1.statusDenied.checked = false;
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

function goToCard(sMemberID,sPoolPassID,sAction) {
  if(sAction=="CREATE") {
     location.href='image_takepic.asp?memberid=' + sMemberID + '&poolpassid=' + sPoolPassID + '&demo=<%=lcl_demo%>';
  }else if (sAction=="REPRINT") {
     location.href='image_display.asp?memberid=' + sMemberID + '&poolpassid=' + sPoolPassID + '&action=REPRINT&demo=<%=lcl_demo%>';
  }
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

		var daterege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;

		var dateFromOk = daterege.test(document.getElementById("sc_from_date").value);
		var dateToOk   = daterege.test(document.getElementById("sc_to_date").value);

		if (! dateFromOk ) {
      document.getElementById("sc_from_date").focus();
      inlineMsg(document.getElementById("sc_from_date_cal").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'sc_from_date_cal');
      return false;
  }else{
      clearMsg("sc_from_date_cal");
  }

		if (! dateToOk ) {
      document.getElementById("sc_to_date").focus();
      inlineMsg(document.getElementById("sc_to_date_cal").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'sc_to_date_cal');
      return false;
  }else{
      clearMsg("sc_to_date_cal");
  }

  return true;
}

  //-->
  </script>

</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>" & session("sOrgName") & "&nbsp;" & oMembership.GetMembershipName() & ":&nbsp;Search, Create, Reprint Member IDs" & lcl_demo_title & "</strong></font></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'BEGIN: Search Options -------------------------------------------------------
  response.write "  <tr class=""noprint"">" & vbcrlf
  response.write "     <td class=""noprint"">" & vbcrlf
  response.write "        	<fieldset class=""fieldset"">" & vbcrlf
  response.write "         		<legend>Search Options</legend>" & vbcrlf
  response.write "           <form action=""search_members.asp"" method=""post"" name=""searchform"">" & vbcrlf
  response.write "             <input type=""hidden"" name=""qry"" value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "             <input type=""hidden"" name=""demo"" value=""" & lcl_demo & """ size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "      			<table border=""0"" cellpadding=""2"" cellspacing=""0"" id=""searchtable"" style=""width: 600px"">" & vbcrlf
  response.write "           <tr>" & vbcrlf
  response.write "               <td>Name:</td>" & vbcrlf
  response.write "               <td><input type=""text"" name=""username"" id=""username"" value=""" & lcl_search_name & """ size=""30"" maxlength=""50"" /></td>" & vbcrlf
  response.write "               <td>Show Expired:&nbsp;<input type=""checkbox"" name=""sc_show_expired"" value=""Y""" & lcl_showexpired_checked & " /></td>" & vbcrlf
  response.write "           </tr>" & vbcrlf
  response.write "           <tr>" & vbcrlf
  response.write "               <td>Expiration Date:</td>" & vbcrlf
  response.write "               <td colspan=""2"">" & vbcrlf
  response.write "                   From: <input type=""text"" name=""sc_from_date"" id=""sc_from_date"" value=""" & lcl_sc_from_date & """ size=""10"" maxlength=""10"" onchange=""clearMsg('sc_from_date_cal');"" />" & vbcrlf
  response.write "                   <img src=""../images/calendar.gif"" id=""sc_from_date_cal"" border=""0"" style=""cursor: hand"" onclick=""clearMsg('sc_from_date_cal');doCalendar('sc_from_date')"" />" & vbcrlf
  response.write "                   &nbsp;&nbsp;" & vbcrlf
  response.write "                   To: <input type=""text"" name=""sc_to_date"" id=""sc_to_date"" value=""" & lcl_sc_to_date & """ size=""10"" maxlength=""10"" onchange=""clearMsg('sc_to_date_cal');"" />" & vbcrlf
  response.write "                   <img src=""../images/calendar.gif"" id=""sc_to_date_cal"" border=""0"" style=""cursor: hand"" onclick=""clearMsg('sc_to_date_cal');doCalendar('sc_to_date')"" />" & vbcrlf
  response.write "               </td>" & vbcrlf
  response.write "           </tr>" & vbcrlf

  if lcl_family_id <> "" then
     response.write "<tr>" & vbcrlf
     response.write "    <td colspan=""3"">Searching on Family ID: [<span style=""color: #ff0000"">" & lcl_family_id & "</span>]</td>" & vbcrlf
     response.write "</tr>" & vbcrlf
  end if

  response.write "           <tr>" & vbcrlf
  response.write "               <td colspan=""3"">&nbsp;</td>" & vbcrlf
  response.write "           </tr>" & vbcrlf
  response.write "           <tr>" & vbcrlf
  response.write "               <td colspan=""3""><input type=""submit"" class=""button"" value=""Search"" onclick=""return validateFields()"" /></td>" & vbcrlf
  response.write "           </tr>" & vbcrlf
  response.write "      			</table>" & vbcrlf
  response.write "           </form>" & vbcrlf
  response.write "		      	</fieldset>" & vbcrlf
  response.write "     </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: Search Options ---------------------------------------------------------

 'BEGIN: Members List ---------------------------------------------------------
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf

 	PoolPassMembershipList lcl_search_name, _
                         oMembership.MembershipId, _
                         lcl_family_id, _
                         lcl_demo, _
                         lcl_sc_show_expired, _
                         lcl_sc_from_date, _
                         lcl_sc_to_date, _
                         lcl_orghasfeature_memberships_usekeycards

  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
 'END: Members List -----------------------------------------------------------

  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

Set oMembership = Nothing 
'------------------------------------------------------------------------------
'
Sub PoolPassMembershipList(p_sc_username, _
                           iMembershipId, _
                           p_family_id, _
                           p_demo, _
                           p_sc_show_expired, _
                           p_sc_from_date, _
                           p_sc_to_date, _
                           iOrgHasFeatureMemberShipsUseKeycards)

	dim sSQL, lcl_current_date, lcl_age, lcl_display_expire, lcl_photo_id_text, lcl_photo_id_url

	If oUserName = "" Then
  		oUserName = "%"
	End If
	
'Set the page size
 pagesize = GetUserPageSize( Session("UserId") )

'Set the current page
 if not isempty(request.querystring("pagenum")) then
	   currentpage = clng(request.querystring("pagenum"))
 else
	   currentpage = 1
 end if

	sSQL = sSQL & "SELECT "
 sSQL = sSQL & " U.userid, "
 sSQL = sSQL & " U.userfname, "
 sSQL = sSQL & " U.userlname, "
 sSQL = sSQL & " U.useraddress, "
 sSQL = sSQL & " U.birthdate, "
 sSQL = sSQL & " U.familyid, "
 sSQL = sSQL & " R.description, "
 sSQL = sSQL & " T.description as residenttype, "
 sSQL = sSQL & " FM.relationship as relationship_type, "
 sSQL = sSQL & " P.paymentdate, "
 sSQL = sSQL & " P.startdate, "
 sSQL = sSQL & " P.expirationdate, "
 sSQL = sSQL & " PPM.memberid, "
 sSQL = sSQL & " PPM.card_printed, "
 sSQL = sSQL & " M.membershipdesc, "
 sSQL = sSQL & " MP.period_desc, "
 sSQL = sSQL & " MP.period_interval, "
 sSQL = sSQL & " MP.period_qty, "
 sSQL = sSQL & " MP.period_type, "
 sSQL = sSQL & " R.attendancetypeid, "
 sSQL = sSQL & " R.isPunchcard, "
 sSQL = sSQL & " P.poolpassid "
	sSQL = sSQL & "FROM egov_poolpasspurchases P, "
	sSQL = sSQL &     " egov_users U, "
	sSQL = sSQL &     " egov_poolpassrates R, "
	sSQL = sSQL &     " egov_poolpassresidenttypes T, "
	sSQL = sSQL &     " egov_membership_periods MP, "
 sSQL = sSQL &     " egov_memberships M, "
	sSQL = sSQL &     " egov_familymembers FM, "
	sSQL = sSQL &     " egov_poolpassmembers PPM "
	sSQL = sSQL & "WHERE P.orgid = " & session("orgid")
	sSQL = sSQL & " AND upper(P.paymentresult) <> 'PENDING' "
	sSQL = sSQL & " AND upper(P.paymentresult) <> 'DECLINED' "
	sSQL = sSQL & " AND FM.familymemberid  = PPM.familymemberid "
	sSQL = sSQL & " AND FM.userid          = U.userid "
	sSQL = sSQL & " AND FM.belongstouserid = U.familyid "
	sSQL = sSQL & " AND P.poolpassid       = PPM.poolpassid "
	sSQL = sSQL & " AND P.rateid           = R.rateid "
	sSQL = sSQL & " AND P.periodid         = MP.periodid "
 sSQL = sSQL & " AND P.membershipid     = M.membershipid "
	sSQL = sSQL & " AND P.orgid            = MP.orgid "
	sSQL = sSQL & " AND R.residenttype     = T.resident_type "
	sSQL = sSQL & " AND T.orgid            = P.orgid "
 sSQL = sSQL & " AND R.membershipid     = " & iMembershipId

'If "Show Expired" has been checked then do not filter out the expired memberships
 if p_sc_show_expired <> "Y" then
    sSQL = sSQL & " AND CAST(P.expirationdate AS datetime) >= CAST('" & date() & "' AS datetime)"
 end if

'If the family id has been clicked on then query on the family id value.
'Otherwise, query on the first/last name
 if p_family_id <> "" then
    sSQL = sSQL & " AND  U.familyid = " & p_family_id
 else
    sSQL = sSQL & " AND  ((upper(U.userlname) like (upper('%" & dbsafe(p_sc_username) & "%'))) "
    sSQL = sSQL & " OR   (upper(U.userfname) like (upper('%" & dbsafe(p_sc_username) & "%')))) "
 end if

 sSQL = sSQL & " AND CAST(P.expirationdate AS datetime) BETWEEN CAST('" & p_sc_from_date & "' AS datetime) AND CAST('" & p_sc_to_date & "' AS datetime) "

'Setup ORDER BY
 if lcl_sc_show_expired <> "Y" then
   	sSQL = sSQL & "	ORDER BY 3, 2, 10 "
 else
   	sSQL = sSQL & "	ORDER BY 11, 3, 2, 10 "
 end if

if lcl_initial_search <> "Y" then
  	'set conn = Server.CreateObject("ADODB.Connection")
  	'conn.Open Application("DSN")
  	'set oMemberlist = Server.CreateObject("ADODB.Recordset")
  	'set oMemberlist.ActiveConnection = conn
  	'oMemberlist.CursorLocation = 3
   'oMemberlist.CursorType     = 3
  	'oMemberlist.Open sSQL

 		set oMemberlist = Server.CreateObject("ADODB.Recordset")
	 	oMemberlist.PageSize       = pagesize
	 	oMemberlist.CacheSize      = pagesize
	 	oMemberlist.CursorLocation = 3
	 	oMemberlist.Open sSQL, Application("DSN"), 3,1

 		if (len(currentpage) = 0 or clng(currentpage) < 1) and not oMemberlist.eof then
      oMemberlist.AbsolutePage = 1
 		elseif not oMemberlist.eof then
      if clng(currentpage) <= oMemberlist.PageCount then
 				    oMemberlist.AbsolutePage = currentpage
      else
	        oMemberlist.AbsolutePage = 1
			  end if
 		end if

 		Dim abspage, pagecnt
 		abspage = oMemberlist.AbsolutePage
 		pagecnt = oMemberlist.PageCount

  	lcl_totalrecords = oMemberlist.RecordCount
   lcl_totalpages   = (lcl_totalrecords \ pagesize) + 1  '\means integer/integer

  	if lcl_totalrecords Mod pagesize = 0 and lcl_totalpages > 0 then
      lcl_totalpages = lcl_totalpages-1
   end if

   if lcl_totalrecords <= pagesize then
      lcl_totalpages = 1
   end if

  	if lcl_totalpages < 1 then
      lcl_totalpages = 1
   end if

   if isNumeric(currentpage) then
  	   if currentpage < 1 then
     	   currentpage = 1
    		end if
    		if currentpage > lcl_totalpages then
 	     		currentpage = lcl_totalpages
    		end if
   else
      currentpage = 1
  	end if

   numstartid	= (currentpage-1) * PageSize
  	numendid	  = IIf(numstartid + PageSize < lcl_totalrecords, numstartid+pagesize- 1, lcl_totalrecords - 1)

  	if request("pagenum") <> "" then
    		pagenum  = request("pagenum")
	     sPageNum = "&pagenum=" & pagenum
  	else
    		pagenum  = 0
	     sPageNum = ""
  	end if

 	 if oMemberlist.eof then
    		response.write "<p><strong>No records found</strong></p>"
  	else
    		response.write vbcrlf
      response.write "<div style=""padding-bottom:5px;"">" & vbcrlf
      response.write "Number of Members: [" & oMemberlist.RecordCount & "]"

     'Set up the Previous and Next buttons
      if lcl_totalpages > 1 then
         response.write "&nbsp;&nbsp;Number of Pages: [" & lcl_totalpages & "]"
         response.write "&nbsp;&nbsp;Current Page: ["    & currentpage    & "]" & vbcrlf
         response.write "<br /><br />" & vbcrlf

        'Previous
         if currentpage > 1 then
            response.write "<a href=""search_members.asp?" & limit & "pagenum=" & currentpage-1 & "&username=" & p_sc_username & "&sc_show_expired=" & p_sc_show_expired & "&sc_from_date=" & p_sc_from_date & "&sc_to_date=" & p_sc_to_date & "&demo=" & p_demo & """>" & vbcrlf
         end if
  
         response.write "<img src='../images/arrow_back.gif' align='absmiddle' border=0>&nbsp;" & langPrev & "&nbsp;" & pagesize

         if currentpage > 1 then
            response.write "</a>" & vbcrlf
         end if

         response.write "&nbsp;&nbsp;" & vbcrlf

        'Next
        	if currentpage < lcl_totalpages then 
            response.write "<a href=""search_members.asp?" & limit & "pagenum=" & currentpage+1 & "&username=" & p_sc_username & "&sc_show_expired=" & p_sc_show_expired & "&sc_from_date=" & p_sc_from_date & "&sc_to_date=" & p_sc_to_date & "&demo=" & p_demo & """>" & vbcrlf
         end if

         response.write langNext & "&nbsp;" & pagesize & "<img src='../images/arrow_forward.gif' align='absmiddle' border='0' />" & vbcrlf

         if abspage < lcl_totalpages then
            response.write "</a>" & vbcrlf
         end if
      end if

      response.write "</div>" & vbcrlf
    		response.write "<div class=""shadow"">" & vbcrlf
  		  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""100%"">" & vbcrlf
    		response.write "  <tr align=""center"" valign=""bottom"">" & vbcrlf
    		'response.write "      <th>&nbsp;</th>" & vbcrlf
  		  response.write "      <th nowrap=""nowrap"">Member<br />#</th>" & vbcrlf

      if iOrgHasFeatureMemberShipsUseKeycards then
         response.write "      <th nowrap=""nowrap"">Barcode<br />#</th>" & vbcrlf
      end if

    		response.write "      <th nowrap=""nowrap"" align=""left"">Name</th>" & vbcrlf
  		  response.write "      <th nowrap=""nowrap"">Age</th>" & vbcrlf
      response.write "      <th nowrap=""nowrap"">Family<br />ID</th>" & vbcrlf
  		  response.write "      <th nowrap=""nowrap"" align=""left"">Relationship&nbsp;</th>" & vbcrlf
    		response.write "      <th nowrap=""nowrap"" align=""left"">Home Address</th>" & vbcrlf
  		  response.write "      <th nowrap=""nowrap"" align=""left"">Membership Type</th>" & vbcrlf
    		response.write "      <th nowrap=""nowrap"">&nbsp;Membership&nbsp;<br />Start Date</th>" & vbcrlf
    		response.write "      <th nowrap=""nowrap"">&nbsp;Membership&nbsp;<br />Expires</th>" & vbcrlf
     	response.write "      <th nowrap=""nowrap"">Photo<br />ID</th>" & vbcrlf

      if  lcl_orghasfeature_pool_attendance_view _
      AND lcl_orghasfeature_customreports_membership_scanlog _
      AND lcl_userhaspermission_customreports_membership_scanlog then
          response.write "      <th nowrap=""nowrap"">Scan Log</th>" & vbcrlf
      end if

      response.write "  </tr>" & vbcrlf

     'Find the total number of times to "loop"
      if CLng(oMemberlist.RecordCount) < CLng(pagesize) then
         lcl_total_pagerecords = oMemberlist.RecordCount
      else
         lcl_total_pagerecords = pagesize
      end if

      if abspage > 1 then
         lcl_loop_total = oMemberlist.RecordCount - (abspage*lcl_total_pagerecords)

         if CLng(lcl_loop_total) > CLng(lcl_total_pagerecords) then
            lcl_loop_total = lcl_total_pagerecords
         end if
      else
         lcl_loop_total = lcl_total_pagerecords
      end if

      if lcl_loop_total = "" then
         lcl_loop_total = 0
      end if

     '----------------------------------------------------------------------
      iRowCount = 0
  		  bgcolor   = "#eeeeee"

      for intRec = numstartid to numendid
         lcl_userid           = oMemberlist("userid")
         lcl_userfname        = oMemberlist("userfname")
         lcl_userlname        = oMemberlist("userlname")
         lcl_useraddress      = oMemberlist("useraddress")
         lcl_birthdate        = oMemberlist("birthdate")
         lcl_familyid         = oMemberlist("familyid")
         lcl_description      = oMemberlist("description")
         lcl_residenttype     = oMemberlist("residenttype")
         lcl_relationshiptype = oMemberlist("relationship_type")
         lcl_paymentdate      = oMemberlist("paymentdate")
         lcl_startdate        = oMemberlist("startdate")
         lcl_expiration_date  = oMemberlist("expirationdate")
         lcl_memberid         = oMemberlist("memberid")
         lcl_card_printed     = oMemberlist("card_printed")
         lcl_membershipdesc   = oMemberlist("membershipdesc")
         lcl_period_desc      = oMemberlist("period_desc")
         lcl_period_interval  = oMemberlist("period_interval")
         lcl_period_qty       = oMemberlist("period_qty")
         lcl_period_type      = oMemberlist("period_type")
         lcl_attendancetypeid = oMemberlist("attendancetypeid")
         lcl_isPunchcard      = oMemberlist("isPunchcard")
         lcl_poolpassid       = oMemberlist("poolpassid")

        'Get the rateid
         lcl_rateid = getRateID(lcl_poolpassid)

        'Format the age
         if lcl_birthdate = "" OR IsNull(lcl_birthdate) then
            lcl_age = "Adult"
         else
            lcl_age = GetCitizenAge(lcl_birthdate)
         end if 

        'Format the expiration date
         lcl_display_expire = lcl_expiration_date

      	 'Determine if the card has already been printed
         if lcl_card_printed = "Y" then
  	         lcl_photo_id_text = "Reprint"
  		        lcl_photo_id_url  = "javascript:goToCard('" & lcl_memberid & "','" & lcl_poolpassid & "','REPRINT');"
      	  else
            lcl_photo_id_text = "Create"
 		   	     lcl_photo_id_url  = "javascript:goToCard('" & lcl_memberid & "','" & lcl_poolpassid & "','CREATE');"
      	  end if

         response.write "  <tr bgcolor=""" &  bgcolor  & """ class=""tablelist"" valign=""top"">" & vbcrlf
      	  response.write "      <td align=""center"" nowrap=""nowrap"">" & lcl_memberid & "</td>" & vbcrlf

        'If the feature is enabled, check to see if any barcodes are associated to the member.
        'If "yes" then build the list and display the barcode icon.
        'If "no" then do not display anything.
         if iOrgHasFeatureMemberShipsUseKeycards then
            sBarcodeList       = ""
            sDisplayBarcodeImg = "&nbsp;"

            sBarcodeList = buildTooltip_BarcodeList(session("orgid"), _
                                                    lcl_memberid)

            if sBarcodeList <> "" then
               sDisplayBarcodeImg = "<img src=""images/icon_barcode2.png"" width=""22"" height=""18"" onMouseOver=""tooltip.show('" & sBarcodeList & "');"" onMouseOut=""tooltip.hide();"" style=""cursor: pointer;"" />"
            end if
         
         	  response.write "      <td align=""center"" nowrap=""nowrap"">" & sDisplayBarcodeImg & "</td>" & vbcrlf
         end if

    	    response.write "      <td nowrap=""nowrap""><a href=""javascript:Open_Profile('" & lcl_userid & "','" & lcl_familyid & "');"" onMouseOver=""tooltip.show('Click to Edit User');"" onMouseOut=""tooltip.hide();"">" & lcl_userlname & ", " & lcl_userfname & "</a></td>" & vbcrlf
      		 response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_age & "</td>" & vbcrlf
 	       response.write "      <td nowrap=""nowrap"" align=""center"">" & vbcrlf
         response.write "          <a href=""search_members.asp?familyid=" & lcl_familyid & "&demo=" & p_demo & "&username=" & p_sc_username & "&sc_show_expired=" & p_sc_show_expired & "&sc_from_date=" & p_sc_from_date & "&sc_to_date=" & p_sc_to_date & """ onMouseOver=""tooltip.show('Click to search on this Family ID');"" onMouseOut=""tooltip.hide();"">" & lcl_familyid & "</a></td>" & vbcrlf
         response.write "      <td nowrap=""nowrap"">" & lcl_relationship_type & "&nbsp;</td>" & vbcrlf
      		 response.write "      <td nowrap=""nowrap"">" & lcl_useraddress       & "</td>" & vbcrlf
 	  	    response.write "      <td nowrap=""nowrap"" style=""color:#ff0000"">&nbsp;" & lcl_residenttype & " &mdash; " & Trim(lcl_description) & "&nbsp;</td>" & vbcrlf

        'If the expiration date is LESS THAN today's date then change the font color to DARK RED
        'This will also determine if the "create" or "reprint" link is displayed or not.
         if cdate(lcl_display_expire) <= cdate(date()) then
            lcl_font_color              = " style=""color: #ff0000"""
            lcl_show_CreateReprint_link = "N"
         else
            lcl_font_color              = ""
            lcl_show_CreateReprint_link = "Y"
         end if

         response.write "      <td nowrap=""nowrap"" align=""right"">" & datevalue(lcl_startdate) & "</td>" & vbcrlf
    	    response.write "      <td nowrap=""nowrap"" align=""right""" & lcl_font_color & ">" & datevalue(lcl_display_expire) & "</td>" & vbcrlf

         if lcl_show_CreateReprint_link = "Y" then
    	       if not lcl_orghasfeature_pool_attendance_view then
        	      response.write "      <td><input type=""button"" value=""" & lcl_photo_id_text & """ style=""cursor: hand"" onMouseOver=""tooltip.show('Click to <strong style=\'color:#ffff00\'>" & UCASE(lcl_photo_id_text) & "</strong> a Membership Card');"" onMouseOut=""tooltip.hide();"" onclick=""" & lcl_photo_id_url & """ /></td>" & vbcrlf
            else
              '----------------------------------------------------------------
             	'1. If the membership purchase is NOT an attendance type of "Member" (attendancetypeid = 1 on EGOV_POOL_ATTENDANCETYPES)
              '   then hide the link to create/reprint ids.  In this case the employee (admin) will scan the custom barcode that
  	           '   represents what the user bought.
              '2. As of 01/30/09: if the attendance type <> "Member" then check to see if the rate is set up as a "punchcard".
              '   If "yes" then show the Reprint/Create button.
              '----------------------------------------------------------------
               if lcl_attendancetypeid = 1 OR lcl_isPunchcard then
                  response.write "      <td><input type=""button"" value=""" & lcl_photo_id_text & """ style=""cursor: hand"" onMouseOver=""tooltip.show('Click to <strong style=\'color:#ffff00\'>" & UCASE(lcl_photo_id_text) & "</strong> a Membership Card');"" onMouseOut=""tooltip.hide();"" onclick=""" & lcl_photo_id_url & """ /></td>" & vbcrlf
      	        else
          	       response.write "      <td nowrap=""nowrap"">&nbsp;</td>" & vbcrlf
               end if
      	     end if
         else
            response.write "      <td nowrap=""nowrap"">&nbsp;</td>" & vbcrlf
         end if

        'Set up scan log history query.
        'Only show the "View Log" button if:
        '  1. The org has the "View/Edit Pool Daily Attendance" feature turned on .
        '  2. The org and user has the "Membership Scan Log" custom report feature assigned.
        '  3. The memberid has at least one scan history record.
         if  lcl_orghasfeature_pool_attendance_view _
         AND lcl_orghasfeature_customreports_membership_scanlog _
         AND lcl_userhaspermission_customreports_membership_scanlog then
             if checkMemberIDScanned(lcl_memberid,lcl_rateid) then
                response.write "      <td><input type=""button"" value=""View Log"" style=""cursor: hand"" onMouseOver=""tooltip.show('Click to view the scan history for this membership');"" onMouseOut=""tooltip.hide();"" onclick=""openCustomReports('membershipcard_scans_by_memberid','" & lcl_memberid & "','" & lcl_rateid & "')"" /></td>" & vbcrlf
             else
                response.write "      <td>&nbsp;</td>" & vbcrlf
             end if
         end if

      	  response.write "  </tr>" & vbcrlf

         bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")

         oMemberlist.movenext
      next

   		 response.write "</table>" & vbcrlf
   		 response.write "</div>"   & vbcrlf

   end if

   oMemberlist.close
   set oMemberlist = nothing

end if

end sub

'------------------------------------------------------------------------------
Function ShowPoolPassMembers(iPoolPassId, sBgcolor)
	Dim sSQL, sNameList, sRelationList, sBreak, bFirst
	sNameList = ""
	sBreak = ""
	bFirst = True 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetPoolPassMembersList"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iPoolPassId", 3, 1, 4, iPoolPassId)
		Set oMembers = .Execute
	End With

	Do While Not oMembers.eof 
		If Not bFirst Then
			sBreak = "<br />"
		Else 
			bFirst = False 
		End If 
		sNameList = sNameList & sBreak & oMembers("firstname") & " " & oMembers("lastname")
		sRelationList = sRelationList & sBreak & TranslateMember(oMembers("relationship"))
		oMembers.movenext
	Loop 
	ShowPoolPassMembers = vbcrlf & "<tr bgcolor=""" &  sBgcolor  & """ class=""tablelist"" ><td>&nbsp;</td><td>&nbsp;</td><td class=""familylist"" nowrap=""nowrap"">" & sNameList & "</td><td nowrap=""nowrap"" colspan=""3"">" & sRelationList &"</td></tr>"
		
	oMembers.close
	Set oMembers = Nothing
	Set oCmd = Nothing

End Function 

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
Function GetInitialMembershipId( iOrgID )
	Dim sSql, oMember

	sSQL = "Select MIN(membershipid) as membershipid FROM egov_memberships WHERE orgid = " & iOrgID 
	
	Set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), 3, 1
	
	If IsNull(oMember("membershipid")) Then
		GetInitialMembershipId = 0
	Else
		GetInitialMembershipId = oMember("membershipid")
	End If 
	
	oMember.close
	Set oMember = Nothing
End Function 

'------------------------------------------------------------------------------
Function ShowMembershipPicks(iMembershipId, iOrgId)
	Dim sSQL, oMembers

	' Get the memberships
	sSQL = "Select membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & iOrgId & " order by membershipdesc"
	ShowMembershipPicks = ""

	Set oMembers = Server.CreateObject("ADODB.Recordset")
	oMembers.Open sSQL, Application("DSN"), 3, 1
	
	Do While not oMembers.eof 
		ShowMembershipPicks = ShowMembershipPicks & vbcrlf & "<option value=""" & oMembers("membershipid") & """ "
		If clng(iMembershipId) = clng(oMembers("membershipid"))  Then
			ShowMembershipPicks = ShowMembershipPicks & " selected=""selected"" "
		End If 
		ShowMembershipPicks = ShowMembershipPicks & ">" & oMembers("membershipdesc") & "</option>"
		oMembers.movenext
	Loop 

	oMembers.close
	Set oMembers = Nothing

End Function 

'--------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = ""
  end if

  dbsafe = lcl_value

end function

'--------------------------------------------------------------
function IIf(bCheck, sTrue, sFalse)
		if bCheck then IIf = sTrue Else IIf = sFalse
end function

'------------------------------------------------------------------------------
function buildTooltip_BarcodeList(iOrgID, _
                                  iMemberID)

  dim lcl_return, sSQL

  lcl_return = ""

  sSQL = "SELECT mtb.barcode, "
  sSQL = sSQL & " bs.statusname, "
  sSQL = sSQL & " bs.isActiveStatus "
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes mtb  "
  sSQL = sSQL &      " INNER JOIN egov_poolpassmembers_barcode_statuses bs ON bs.statusid = mtb.barcode_statusid "
  sSQL = sSQL & " WHERE mtb.orgid = " & iOrgID
  sSQL = sSQL & " AND mtb.memberid = " & iMemberID
  sSQL = sSQL & " ORDER BY bs.isActiveStatus DESC, bs.statusname "

  set oTooltipBarcodes = Server.CreateObject("ADODB.Recordset")
  oTooltipBarcodes.Open sSQL, Application("DSN"), 3, 1

  if not oTooltipBarcodes.eof then
     lcl_return = "<table border=\'0\' class=\'barcodeTooltipTable\'>"
     lcl_return = lcl_return & "  <tr>"
     lcl_return = lcl_return & "      <th>Barcode</th>"
     lcl_return = lcl_return & "      <th>Status</th>"
     lcl_return = lcl_return & "  </tr>"
      
     do while not oTooltipBarcodes.eof
        sDisplayBarcode    = oTooltipBarcodes("barcode")
        sDisplayStatusName = oTooltipBarcodes("statusname")

        if oTooltipBarcodes("isActiveStatus") then
           sDisplayBarcode    = "<span class=\'isActiveBarcode\'>" & sDisplayBarcode    & "</span>"
           sDisplayStatusName = "<span class=\'isActiveBarcode\'>" & sDisplayStatusName & "</span>"
        end if

        lcl_return = lcl_return & "  <tr>"
        lcl_return = lcl_return & "      <td>" & sDisplayBarcode     & "</td>"
        lcl_return = lcl_return & "      <td>" & sDisplayStatusName  & "</td>"
        lcl_return = lcl_return & "  </tr>"

        oTooltipBarcodes.movenext
     loop

     lcl_return = lcl_return & "</table>"
  end if

  oTooltipBarcodes.close
  set oTooltipBarcodes = nothing

  buildTooltip_BarcodeList = lcl_return

end function
%>

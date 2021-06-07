<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: member_list.asp
' AUTHOR: Steve Loar
' CREATED: 01/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the Pool Pass membership report
'
' MODIFICATION HISTORY
' 1.0  01/27/2006 Steve Loar - Code added to template
' 2.0  07/12/2006 Steve Loar - Changed to be membership generic
' 2.1  06/22/2010 David Boyer - Added "Send email to Members" button
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
 Dim iMembershipId, oMembership

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission( session("userid"), "membership list" ) then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine the membership type. (Currently set the membershipid to the one for "pools")

 set oMembership = New classMembership
 if request("sMembershipType") <> "" then
    lcl_membership_type = request("sMembershipType")
 else
    'lcl_membership_type = "pool"
    lcl_membership_type = oMembership.GetFirstMembershipType()
 end if
 oMembership.SetMembershipId(lcl_membership_type)

 'orderBy              = request("orderBy")
 subTotals            = request("subTotals")
 showDetail           = request("showDetail")
 fromDate             = request("fromDate")
 toDate               = request("toDate")
 today                = Date()
 sUserlname           = request("userlname")
 lcl_includeInResults = request("sc_includeinresults")

'If orderBy = "" or IsNull(orderBy) Then orderBy = "date" End If
 if toDate = "" or IsNull(toDate) then
    toDate = dateAdd("d",0,today)
 end if

 if fromDate = "" or IsNull(fromDate) then
    fromDate = DateSerial(Year(Now()),1,1)
 end if

'Check for org features
 lcl_orghasfeature_send_emails_to_members = orghasfeature("send_emails_to_members")
%>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>E-Gov Administration Console {Membership List By Purchaser}</title>
  
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="./style_pool.css" />
	<link rel="stylesheet" href="poolpass.css" />
	<link rel="stylesheet" media="print" href="receiptprint.css" />
  <link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">
  
  <script src="https://code.jquery.com/jquery-1.9.1.js"></script>
  <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
  <script src="../scripts/selectAll.js"></script>

<script>
  <!--
//  function checkStat() {
//  if ( !(form1.statusInProgress.checked) &&  !(form1.statusPending.checked) && !(form1.statusRefund.checked) && !(form1.statusDenied.checked) &&  !(form1.statusCompleted.checked) && !(form1.statusProcessed.checked)) {
//   		alert("You must select the status.");
//   		form1.statusPending.focus();
//   		return false;
//  }
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

//  function doCalendar(ToFrom) {
//    w = (screen.width - 350)/2;
//    h = (screen.height - 350)/2;
//    eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
//  }

  function emailMembers() {
    location.href="poolpass_sendemail.asp";
  }
  
  $(function() {
    $( "#toDate" ).datepicker({
      showOn: "button",
      buttonImage: "../images/calendar.gif",
      buttonImageOnly: true,
      changeMonth: true,
      changeYear: true
    });
  });
  
  $(function() {
    $( "#fromDate" ).datepicker({
      showOn: "button",
      buttonImage: "../images/calendar.gif",
      buttonImageOnly: true,
      changeMonth: true,
      changeYear: true
    });
  });
  
  
//-->
</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>" & session("sOrgName") & "&nbsp;Membership List</strong></font></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  'response.write "<tr><td><form action=""member_list.asp"" method=""post"" name=""nameform"">" & vbcrlf
  'response.write "<font size=""+1""><strong>" & Session("sOrgName") & "&nbsp;" & oMembership.GetMembershipName() & "&nbsp;Membership List</strong></font></td></tr></form>" & vbcrlf

  response.write "  <tr class=""noprint"">" & vbcrlf
  response.write "      <td class=""noprint"">" & vbcrlf

 'BEGIN: Search Options -------------------------------------------------------
  lcl_selected_includeall     = ""
  lcl_selected_includecurrent = ""
  lcl_selected_includeexpired = ""

  if lcl_includeInResults = "currentmembers" then
     lcl_selected_includecurrent = " selected=""selected"""
  elseif lcl_includeInResults = "expiredmembers" then
     lcl_selected_includeexpired = " selected=""selected"""
  else
     lcl_selected_includeall = " selected=""selected"""
  end if

  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "      		    <legend><strong>Search/Sorting Option(s)</strong></legend>" & vbcrlf
  'response.write "            <form name=""searchform"" id=""searchform"" action=""member_list.asp"" method=""post"" onSubmit=""return checkStat()"">" & vbcrlf
  response.write "            <form name=""searchform"" id=""searchform"" action=""member_list.asp"" method=""post"">" & vbcrlf
  response.write "            <table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""searchtable"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td><strong>Membership Type:</strong></td>" & vbcrlf
  response.write "                  <td colspan=""4"" valign=""top"" nowrap=""nowrap"">" & vbcrlf
  response.write "                       <select name=""sMembershipType"" id=""sMembershipType"">" & vbcrlf
  			showMembershipTypePicks(lcl_membership_type)
  response.write "                       </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td><strong>From:</strong></td>" & vbcrlf
  response.write "                  <td valign=""top"" nowrap=""nowrap"">" & vbcrlf
  response.write "                       <input type=""text"" name=""fromDate"" id=""fromDate"" value=""" & fromDate & """ />" & vbcrlf
  'response.write "                      <a href=""javascript:void doCalendar('From');""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td nowrap=""nowrap""><strong>Purchaser Last Name:</strong></td>" & vbcrlf
  response.write "                  <td valign=""top"" nowrap=""nowrap"">" & vbcrlf
  response.write "                      <input type=""text"" name=""userlname"" id=""userlname"" value=""" & sUserlname & """ size=""30"" maxlength=""50"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td width=""100%"">&nbsp;</td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td><strong>To:</strong></td>" & vbcrlf
  response.write "                  <td valign=""top"" nowrap=""nowrap"">" & vbcrlf
  response.write "                      <input type=""text"" name=""toDate"" id=""toDate"" value=""" & toDate & """ />" & vbcrlf
  'response.write "                      <a href=""javascript:void doCalendar('To');""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td nowrap=""nowrap""><strong>Include in results:</strong></td>" & vbcrlf
  response.write "                  <td valign=""top"" nowrap=""nowrap"">" & vbcrlf
  response.write "                      <select name=""sc_includeinresults"" id=""sc_includeinresults"">" & vbcrlf
  response.write "                        <option value=""all""" & lcl_selected_includeall & ">All</option>" & vbcrlf
  response.write "                        <option value=""currentmembers""" & lcl_selected_includecurrent & ">Current Members Only</option>" & vbcrlf
  response.write "                        <option value=""expiredmembers""" & lcl_selected_includeexpired & ">Expired Members Only</option>" & vbcrlf
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td width=""100%"">&nbsp;</td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td><input type=""submit"" class=""button"" value=""Search"" /></td>" & vbcrlf
  response.write "                  <td>&nbsp;</td>" & vbcrlf
  response.write "                  <td colspan=""3"" align=""right"" nowrap=""nowrap"">" & vbcrlf
  response.write "                      <input type=""button"" class=""button"" value=""Print"" onclick=""javascript:window.print();"" />" & vbcrlf

  if lcl_orghasfeature_send_emails_to_members then
     response.write "                      <input type=""button"" class=""button"" value=""Send Email to Members"" onclick=""emailMembers();"" />" & vbcrlf
  end if

  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "            </form>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
 'END: Search Options ---------------------------------------------------------

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""3"" valign=""top"">" & vbcrlf
                            PoolPassMembershipList fromDate, _
                                                   toDate, _
                                                   sUserlname, _
                                                   oMembership.MembershipId, _
                                                   lcl_includeInResults
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

  set oMembership = nothing 

'------------------------------------------------------------------------------
sub PoolPassMembershipList(fromDate, toDate, sUserlname, iMembershipId, iIncludeInResults)
  lcl_userlname = ""

 	if sUserlname <> "" then
     lcl_userlname = ucase(sUserlname)
     lcl_userlname = dbsafe(lcl_userlname)
     lcl_userlname = "'%" & lcl_userlname & "%'"
  end if

 	toDate = dateAdd("d",1,toDate)

  sSQL = "SELECT P.poolpassid, "
  sSQL = sSQL & " U.userid, "
  sSQL = sSQL & " U.userfname, "
  sSQL = sSQL & " U.userlname, "
  sSQL = sSQL & " U.useraddress, "
  sSQL = sSQL & " U.userhomephone, "
  sSQL = sSQL & " U.useremail, "
  sSQL = sSQL & " P.paymenttype, "
  sSQL = sSQL & " R.description, "
  sSQL = sSQL & " T.description as residenttype, "
  sSQL = sSQL & " P.paymentdate, "
  sSQL = ssQL & " P.expirationdate, "
  sSQL = sSQL & " P.paymentresult, "
  sSQL = sSQL & " M.membershipdesc, "
  sSQL = sSQL & " MP.period_desc, "
  sSQL = sSQL & " MP.period_interval, "
  sSQL = sSQL & " MP.period_qty, "
  sSQL = sSQL & " MP.period_type, p.note, au.firstname, au.lastname "
  sSQL = sSQL & " FROM egov_poolpasspurchases P "
		sSql = sSql & " INNER JOIN egov_users U ON U.userid = p.userid "
		sSql = sSql & " INNER JOIN egov_poolpassrates R ON P.rateid = r.rateid "
		sSql = sSql & " INNER JOIN egov_poolpassresidenttypes T ON r.residenttype = t.resident_type AND T.orgid = P.orgid "
		sSql = sSql & " INNER JOIN egov_memberships M ON M.membershipid = P.membershipid "
		sSql = sSql & " INNER JOIN egov_membership_periods MP ON mp.periodid = p.periodid "
		sSql = sSql & " LEFT JOIN users au ON au.userid = P.adminid"
  sSQL = sSQL & "	WHERE P.orgid = " & session("orgid")
  sSQL = sSQL & "	AND P.paymentresult <> 'Pending' "
  sSQL = sSQL & "	AND P.paymentresult <> 'Declined' "
  sSQL = sSQL & "	AND R.membershipid = " & iMembershipID
  sSQL = sSQL & "	AND (P.paymentdate >= '" & fromDate & "' AND P.paymentdate < '" & dateAdd("d",1,toDate)& "') "

  if lcl_userlname <> "" then
     sSQL = sSQL & "	AND upper(U.userlname) like (" & lcl_userlname & ") "
  end if

  if iIncludeInResults <> "" AND iIncludeInResults <> "all" then
     if iIncludeInResults = "expiredmembers" then
        sSQL = sSQL & " AND P.expirationdate < '" & dateadd("d",1,date()) & "' "
     else
        sSQL = sSQL & " AND P.expirationdate >= '" & date() & "' "
     end if
  end if

  sSQL = sSQL & "	ORDER BY userlname, userfname "

 	set oRequests = Server.CreateObject("ADODB.Recordset")
	 oRequests.Open sSQL, Application("DSN"), 0, 1

 	if oRequests.EOF then
     session("sendEmailsToMembers_query") = ""
   		response.write "<p><strong>No records found</strong></p>" & vbcrlf
  else
     session("sendEmailsToMembers_query") = sSQL

   		'response.write "<div class=""shadow"">" & vbcrlf
  	 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" id=""memberlist"" width=""100%"">" & vbcrlf
   		response.write "  <tr class=""tablelist"" align=""left"">" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "      <th>Pass ID</th>" & vbcrlf
     response.write "      <th>Name</th>" & vbcrlf
     response.write "      <th>Home Address</th>" & vbcrlf
     response.write "      <th>Home Phone</th>" & vbcrlf
     response.write "      <th>Membership Type</th>" & vbcrlf
     if session("orgid") = "26" then
     response.write "      <th>Note</th>" & vbcrlf
     response.write "      <th>Admin Processor</th>" & vbcrlf
     end if
     response.write "  </tr>" & vbcrlf

   		bgcolor = "#eeeeee"

  		 do while not oRequests.eof
        bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")
																					
    			 response.write "  <tr bgcolor=""" &  bgcolor  & """ class=""tablelist"">" & vbcrlf
        response.write "      <td>&nbsp;</td>" & vbcrlf
 						 response.write "      <td align=""left"" valign=""top"">"    & oRequests("poolpassid")                 & "</td>" & vbcrlf
 						 response.write "      <td nowrap=""nowrap"" valign=""top"">" & oRequests("userlname") & ", " & oRequests("userfname") & "</td>" & vbcrlf
 						 response.write "      <td nowrap=""nowrap"" valign=""top"">" & oRequests("useraddress")                & "</td>" & vbcrlf
 						 response.write "      <td nowrap=""nowrap"" valign=""top"">" & FormatPhone(oRequests("userhomephone")) & "</td>" & vbcrlf
 						 response.write "      <td nowrap=""nowrap"" valign=""top"">" & oRequests("residenttype") & " &mdash; " & Trim(oRequests("description")) & " &mdash; " & Trim(oRequests("period_desc")) 
 						 response.write           "<br />Purchased: " & DateValue(oRequests("paymentdate"))

        if oRequests("expirationdate") < today then
           lcl_expiration_label = "<span style=""color:#800000"">Expired: </span>"
        else
           lcl_expiration_label = "Expires: "
        end if

  						if oRequests("period_type") = "season" then
  			 					'response.write "<br />Expires: 12/31/" & Year(oRequests("paymentdate"))
    						 response.write "<br />" & lcl_expiration_label & "12/31/" & Year(oRequests("expirationdate"))
 	 					else
  			  				'response.write "<br />Expires: " & DateAdd(oRequests("period_interval"),clng(oRequests("period_qty")),DateValue(oRequests("paymentdate")))
           'if  oRequests("period_interval") <> "" _
           'AND oRequests("period_qty")      <> "" _
           'AND oRequests("paymentdate")     <> "" then
      			  '				response.write "<br />Expires: " & DateAdd(oRequests("period_interval"),CLng(oRequests("period_qty")),DateValue(oRequests("paymentdate")))
           'end if

    						 response.write "<br />" & lcl_expiration_label & DateValue(oRequests("expirationdate"))

 	 					end if

  						response.write "      </td>" & vbcrlf
     						if session("orgid") = "26" then
							response.write "<td>" & oRequests("note") & "</td>" & vbcrlf
							response.write "<td><nobr>" & oRequests("firstname") & " " & oRequests("lastname") & "</td>" & vbcrlf
						end if
        response.write "  </tr>" & vbcrlf
 	 					response.write ShowPoolPassMembers(oRequests("poolpassid"), bgcolor)

 		 				oRequests.MoveNext 
						response.Flush
				 loop

  		 response.write "</table>" & vbcrlf
  		 'response.write "</div>" & vbcrlf

  end if

	oRequests.close
	set oRequests = nothing 
	set oCmd      = nothing

end sub

'------------------------------------------------------------------------------
Function ShowPoolPassMembers(iPoolPassId, sBgcolor)
	Dim sSQL, sNameList, sRelationList, sBreak, bFirst

	sNameList  = ""
	sBreak     = ""
	bFirst     = True
 lcl_return = ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "GetPoolPassMembersList"
		.CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iPoolPassId", 3, 1, 4, iPoolPassId)
		Set oMembers = .Execute
	End With

	do while not oMembers.eof
   	if not bFirst then
    			sBreak = "<br />"
  		else
    			bFirst = False
    end if

   	sNameList     = sNameList & sBreak & oMembers("firstname") & " " & oMembers("lastname") & " (" & oMembers("memberid") & ")"
  		sRelationList = sRelationList & sBreak & TranslateMember(oMembers("relationship"))

		  oMembers.movenext
 loop 

 lcl_return = "  <tr bgcolor=""" &  sBgcolor  & """ class=""tablelist"">" & vbcrlf
 lcl_return = lcl_return & "      <td>&nbsp;</td>" & vbcrlf
 lcl_return = lcl_return & "      <td>&nbsp;</td>" & vbcrlf
 lcl_return = lcl_return & "      <td class=""familylist"" nowrap=""nowrap"">" & sNameList & "</td>" & vbcrlf
 lcl_return = lcl_return & "      <td nowrap=""nowrap"" colspan=""3"">" & sRelationList &"</td>" & vbcrlf
 lcl_return = lcl_return & "  </tr>" & vbcrlf

	ShowPoolPassMembers = lcl_return
		
	oMembers.close
	set oMembers = nothing
	set oCmd     = nothing

end function

'------------------------------------------------------------------------------
function TranslateMember( sRelationship )
  if UCase(sRelationship) = "YOURSELF" then
   		TranslateMember = "Purchaser"
  else
   		TranslateMember = sRelationship
  end if

end function

'------------------------------------------------------------------------------
function MakeProper( sString )
 	if sString = "" then
	   	MakeProper = ""
 	else
	   	MakeProper = UCase(Left(sString,1)) & LCase(Mid(sString,2))
 	end if

end function

'------------------------------------------------------------------------------
function FormatPhone( Number )
 	if Len(Number) = 10 then
   		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  else
   		FormatPhone = Number
  end if

end function

'------------------------------------------------------------------------------
function GetInitialMembershipId( iOrgID )
 	dim sSql, oMember

	 sSQL = "SELECT MIN(membershipid) as membershipid FROM egov_memberships WHERE orgid = " & iOrgID 
	
 	set oMember = Server.CreateObject("ADODB.Recordset")
	 oMember.Open sSQL, Application("DSN"), 0, 1
	
 	if IsNull(oMember("membershipid")) then
	   	GetInitialMembershipId = 0
  else
		   GetInitialMembershipId = oMember("membershipid")
 	end if
	
 	oMember.close
	 Set oMember = nothing

end function

'------------------------------------------------------------------------------
function ShowMembershipPicks(iMembershipId, iOrgId)
	 dim sSQL, oMembers

	'Get the memberships
	 sSQL = "SELECT membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & iOrgId & " ORDER BY membershipdesc"
	 ShowMembershipPicks = ""

 	set oMembers = Server.CreateObject("ADODB.Recordset")
 	oMembers.Open sSQL, Application("DSN"), 0, 1
	
 	do while not oMembers.eof 
   		if clng(iMembershipId) = clng(oMembers("membershipid")) then
     			lcl_selected_membershippicks = " selected=""selected"""
     else
     			lcl_selected_membershippicks = ""
   		end if

   		ShowMembershipPicks = ShowMembershipPicks & "<option value=""" & oMembers("membershipid") & """" & lcl_selected_membershippicks & ">" & oMembers("membershipdesc") & "</option>" & vbcrlf

   		oMembers.movenext
 	loop 

 	oMembers.close
 	set oMembers = nothing

end function

'------------------------------------------------------------------------------
sub showMembershipTypePicks(p_membership_type)

  sSQL = "SELECT membershipid, membership, membershipdesc "
  sSQL = sSQL & " FROM egov_memberships "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY membershipdesc "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     while not rs.eof
        if UCASE(p_membership_type) = UCASE(rs("membership")) then
           lcl_selected = " selected=""selected"""
	   sMembershipTypeName = rs("membershipdesc")
        else
           lcl_selected = ""
        end if

        response.write "  <option value=""" & rs("membership") & """" & lcl_selected & ">" & rs("membershipdesc") & "</option>" & vbcrlf
        rs.movenext
     wend
  end if

  set rs = nothing

end sub
%>

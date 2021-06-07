<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../class/classMembership.asp" -->
<!-- #include file="poolpass_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_receipt.asp
' AUTHOR: Steve Loar
' CREATED: 02/7/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  02/07/06 Steve Loar - Code added
' 1.1  09/09/08 David Boyer - Added Membership Renewals
' 1.2  01/30/09 David Boyer - Added "Punchcard" and "Punchcard Count".
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim iPoolPassId, oMembership
	
	sLevel = "../" ' Override of value from common.asp

	If Not UserHasPermission( Session("UserId"), "purchase membership" ) Then
  		response.redirect sLevel & "permissiondenied.asp"
	End If 

	Set oMembership = New classMembership

	'oMembership.SetMembershipId( "pool" )
	iPoolPassId              = request("iPoolPassId")
	oMembership.MembershipId = GetMembershipId( iPoolPassId ) 
 lcl_renewal_text         = ""

 if request("isRenewedPass") = "Y" then
    lcl_renewal_text = "Renewal "
 end if
%>
<html>
<head>
	<title>E-Gov Administration Consolue {Membership Purchase Receipt}</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="./style_pool.css" />
	<link rel="stylesheet" type="text/css" href="./poolpass.css" />
	<link rel="stylesheet" type="text/css" media="print" href="receiptprint.css" />

<script language="javascript">
	function GoBack() {
   <%
     if UCASE(lcl_renewal_text) = "RENEWAL" then
        lcl_button_text = "Return to Membership Details List"

       'Retrieve the search criteria
        orderBy       = session("orderBy")
        subTotals     = session("subTotals")
        showDetail    = session("showDetail")
        fromDate      = session("fromDate")
        toDate        = session("toDate")
        sUserlname    = session("userlname")
        iMembershipId = session("imembershipid")
       	iPeriodId     = session("iperiodid")

        lcl_url = "?"
        lcl_url = lcl_url & "orderBy="           & orderBy
        lcl_url = lcl_url & "&subTotals="        & subTotals
        lcl_url = lcl_url & "&showDetail="       & showDetail
        lcl_url = lcl_url & "&fromDate="         & fromDate
        lcl_url = lcl_url & "&toDate="           & toDate
        lcl_url = lcl_url & "&userlname=' + escape(""" & sUserlname & """) + '"
        lcl_url = lcl_url & "&membershipid="     & iMembershipId
        lcl_url = lcl_url & "&periodid="         & iperiodid

        response.write "location.href='poolpass_list.asp" & lcl_url & "';" & vbcrlf
     else
        lcl_button_text = "Purchase Another Membership"
        lcl_url         = "poolpass_form.asp"

        response.write "location.href='poolpass_form.asp';" & vbcrlf
     end if
   %>
 }
	</script> 
</head>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<body>" & vbcrlf

 'BEGIN: Third Party Print Control --------------------------------------------
  response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
  'response.write " 	<input disabled type=""button"" value=""Print the page"" onclick=""factory.printing.Print(true)"" />&nbsp;&nbsp;" & vbcrlf
  'response.write " 	<input class=""ie55"" disabled type=""button"" value=""Print Preview..."" onclick=""factory.printing.Preview()"" />" & vbcrlf
  response.write "  <input type=""button"" name=""printButton"" id=""printButton"" value=""Print the page"" onclick=""window.print();"" />" & vbcrlf
  response.write "</div>" & vbcrlf

  'response.write "<object id=""factory"" viewastext style=""display:none"" classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814"" codebase=""../includes/smsx.cab#Version=6,3,434,12""></object>" & vbcrlf
 'END: Third Party Print Control ----------------------------------------------

 'BEGIN: Receipt --------------------------------------------------------------
  lcl_pageTitle = GetCityName()
  lcl_pageTitle = lcl_pageTitle & "&nbsp;"
  lcl_pageTitle = lcl_pageTitle & oMembership.GetMembershipName()
  lcl_pageTitle = lcl_pageTitle & "&nbsp;"
  lcl_pageTitle = lcl_pageTitle & "Membership " & lcl_renewal_text & "Receipt"

  response.write "<div id=""receiptcontent"">" & vbcrlf
  response.write "  <font size=""+1""><strong>" & lcl_pageTitle & "</strong></font>" & vbcrlf
  response.write "  <div id=""receiptlinks"">" & vbcrlf
  response.write "    <input type=""button"" class=""button"" onclick=""GoBack();"" value=""" & lcl_button_text & """ />" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  <div class=""shadow"">" & vbcrlf
  response.write "  		<table border=""0"" cellpadding=""5"" cellspacing=""0"">" & vbcrlf
  response.write "    		<tr>" & vbcrlf
  response.write "          <th colspan=""2"">" & lcl_pageTitle & "</th>" & vbcrlf
  response.write "      </tr>" & vbcrlf
                        oMembership.ShowMembershipInfo( iPoolPassId )
  response.write "  		</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  <div class=""shadow"">" & vbcrlf
  response.write "    <table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  		  <tr>" & vbcrlf
  response.write "          <th colspan=""2"">This membership " & lcase(lcl_renewal_text) & "includes the following people:</th>" & vbcrlf
  response.write "      </tr>" & vbcrlf
                        oMembership.ShowMembers( iPoolPassId )
  response.write "  		</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
 'END: Receipt ----------------------------------------------------------------
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

  set oMembership = nothing

'------------------------------------------------------------------------------
function GetMembershipId( iPoolPassId )
	 dim sSQL, oMembership, lcl_return

  lcl_return = 0

 	sSQL = "SELECT membershipid "
  sSQL = sSQL & " FROM egov_poolpasspurchases "
  sSQL = sSQL & " WHERE poolpassid = " & iPoolPassId

 	set oMembership = Server.CreateObject("ADODB.Recordset")
 	oMembership.Open sSQL, Application("DSN") , 3, 1

 	if not oMembership.EOF then
   		lcl_return = clng(oMembership("membershipid"))
 	end if
		
 	oMembership.Close
	 set oMembership = nothing

  GetMembershipId = lcl_return

end function
%>
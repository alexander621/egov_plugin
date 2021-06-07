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
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the details page called from poolpass_list.asp
'
' MODIFICATION HISTORY
' 1.0	02/07/06	Steve Loar - Initial Version
' 1.1	10/05/06	Steve Loar - Header and nav changed
' 1.2  09/09/08 David Boyer - Added Renewal Memberships
' 2.0	07/28/2010	Steve Loar - Changes for Point and Pay payments
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	 dim iPoolPassId
	
	 sLevel = "../" ' Override of value from common.asp

	 if not UserHasPermission( Session("UserId"), "membership detail" ) then
   		response.redirect sLevel & "permissiondenied.asp"
  end if

 	set oMembership = New classMembership

 	iPoolPassId = CLng(request("iPoolPassId"))

 'Retrieve the search criteria
  orderBy       = session("orderBy")
  subTotals     = session("subTotals")
  showDetail    = session("showDetail")
  fromDate      = session("fromDate")
  toDate        = session("toDate")
  sUserlname    = session("userlname")
  iMembershipId = session("imembershipid")
  iPeriodId     = session("iperiodid")

  lcl_search_criteria = "?"
  lcl_search_criteria = lcl_search_criteria & "orderBy="       & orderBy
  lcl_search_criteria = lcl_search_criteria & "&subTotals="    & subTotals
  lcl_search_criteria = lcl_search_criteria & "&showDetail="   & showDetail
  lcl_search_criteria = lcl_search_criteria & "&fromDate="     & fromDate
  lcl_search_criteria = lcl_search_criteria & "&toDate="       & toDate
  lcl_search_criteria = lcl_search_criteria & "&userlname="    & sUserlname
  lcl_search_criteria = lcl_search_criteria & "&membershipid=" & iMembershipId
  lcl_search_criteria = lcl_search_criteria & "&periodid="     & iperiodid

		lcl_message = ""

		if request("success") = "SN" then
		   lcl_message = "Successfully Created..."
		elseif request("success") = "SU" then
		   lcl_message = "Successfully Updated..."
		elseif request("success") = "SD" then
		   lcl_message = "Successfully Deleted..."
		end if

  if lcl_message <> "" then
		   lcl_message = "<span class=""screenMsg"">*** " & lcl_message & " ***</span>"
  else
     lcl_message = "&nbsp;"
  end if
%>
<html>
<head>
 	<title>E-Gov Administration Consule { Membership Details }</title>

 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="style_pool.css" />
 	<link rel="stylesheet" type="text/css" href="poolpass.css" />
 	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

<style type="text/css">
.screenMsg {
   color: #ff0000;
}

#tableReceiptHeader {
   width: 600px;
   margin-left: 5px;
}

#tableReceipt td {
   padding: 5px;
}

#receiptcontent .shadow,
#receiptcontent table {
   border-radius: 5px;
}

#receiptcontent table th {
   border-top-left-radius: 5px;
   border-top-right-radius: 5px;
}

</style>

<script type="text/javascript">
	<!--

		function validate() 
		{
			var bPicked = false;

			if (document.PassForm.familymemberid.length) 
			{   // Several picked
				var checklength = document.PassForm.familymemberid.length;
				for (i = 0; i < checklength; i++) 
				{
					if (document.PassForm.familymemberid[i].checked) 
					{
						bPicked = true;
						break;
					}
				}
			}
			else
			{  // Just one picked
				if (document.PassForm.familymemberid.checked) 
				{
					bPicked = true;
				}
			}

			if(bPicked) 
			{
				document.PassForm.submit();
			}
			else
			{
				alert("Please select at least one family member from the list before updating.");
			}
		}

	//-->
</script>
</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
 'BEGIN: Receipt Header -------------------------------------------------------
  response.write "<table id=""tableReceipt"" border=""0"">" & vbcrlf
  response.write "	 <tr>" & vbcrlf
  response.write "		    <td>" & vbcrlf
  response.write "			       <font size=""+1""><strong>" & GetCityName() & "&nbsp;" & GetMembershipDesc( iPoolPassId ) & "&nbsp;Membership</strong></font>" & vbcrlf
  response.write "		    </td>" & vbcrlf
  response.write "		    <td align=""right"">" & lcl_message & "</td>" & vbcrlf
  response.write "		</tr>" & vbcrlf
  response.write "		<tr>" & vbcrlf
  response.write "				  <td colspan=""2"">" & vbcrlf
  response.write "					     <input type=""button"" class=""button"" value=""<< Back"" onclick=""location.href='poolpass_list.asp" & lcl_search_criteria & "'"" />" & vbcrlf
  response.write "					     <input type=""button"" class=""button"" value=""Print"" onclick=""window.print();"" />" & vbcrlf
  response.write "				  </td>" & vbcrlf
  response.write "		</tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 'END: Receipt Header ---------------------------------------------------------

 'BEGIN: Receipt Content ------------------------------------------------------
  response.write "<div id=""receiptcontent"">" & vbcrlf
  response.write "		<div class=""shadow"">" & vbcrlf
  response.write "			 <table border=""0"" cellpadding=""5"" cellspacing=""0"">" & vbcrlf
  response.write "  				<tr><th colspan=""2"">&nbsp;</th></tr>" & vbcrlf
                        oMembership.ShowMembershipInfo( iPoolPassId )
  response.write " 			</table>" & vbcrlf
  response.write "		</div>" & vbcrlf
 'END: Receipt Content --------------------------------------------------------

 'BEGIN: Members --------------------------------------------------------------
  response.write "<form name=""PassForm"" id=""PassForm"" method=""post"" action=""updatefamilyonpass.asp"">" & vbcrlf
  response.write "		<input type=""hidden"" name=""poolpassid"" id=""poolpassid"" value=""" & iPoolPassId & """ />" & vbcrlf
  response.write "  <div class=""shadow"">" & vbcrlf
  response.write "		  <table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
  response.write "  		  <tr>" & vbcrlf
  response.write "    		    <th colspan=""3"">" & vbcrlf
  response.write "              This pass includes the following people.&nbsp;" & vbcrlf
  response.write "							       <input type=""button"" class=""button"" onclick=""validate();"" value=""Update Pass Members"" name=""update"" />" & vbcrlf
  response.write "						    </th>" & vbcrlf
  response.write "				  </tr>" & vbcrlf
                        oMembership.ShowFamilyMembers( iPoolPassId )
  response.write "				</table>" & vbcrlf
  response.write "			</div>" & vbcrlf
  response.write "		</form>" & vbcrlf
  response.write "	</div>" & vbcrlf
 'END: Members ----------------------------------------------------------------
%>
	<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
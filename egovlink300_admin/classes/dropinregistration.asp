<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: dropinregistration.asp
' AUTHOR: Steve Loar
' CREATED: 07/5/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows drop in registration of classes using a scanner
'
' MODIFICATION HISTORY
' 1.0   07/5/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iUserIdScan, bUserIsValid, bRegistrationBlocked, re, matches

Set re = New RegExp
re.Pattern = "^\d+$"

sLevel = "../"  'Override of value from common.asp

'Check the page availability and user access rights in one call
PageDisplayCheck "dropinregistration", sLevel	 'In common.asp

iUserIdScan = request("useridscan")

If iUserIdScan <> "" Then
	Set matches = re.Execute(iUserIdScan)
	If matches.Count > 0 Then
		iUserIdScan = CLng(iUserIdScan)
	Else 
		iUserIdScan = ""
	End If 
Else
	iUserIdScan = ""
End If 

bUserIsValid = False 
bRegistrationBlocked = False 

%>
<html lang="en">
<head>
	<meta charset="UTF-8">
 	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="classes.css" />

	<script src="https://code.jquery.com/jquery-1.5.min.js"></script>

	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	 <script>
	<!--

		function validateSearch() {
			if ( validateFields() ) {
				document.frmUserIdScan.submit();
			}
		}
		
		function validateForm() {
			return validateFields();
		}
		
		function validateFields() {
			if ($("#useridscan").val() == '') {
	            	$("#useridscan").focus();
				inlineMsg("useridscan",'<strong>Missing Id: </strong>Please enter an ID Number before searching.',5,"useridscan");
				return false;
			}
			else {
				var useridscan = $("#useridscan").val();
				var rege = /^\d+$/;
				var Ok = rege.test(useridscan);

				if (!Ok) {
					$("#useridscan").focus();
					inlineMsg("useridscan",'<strong>Invalid Format: </strong>The ID Number must be numeric.',5,"useridscan");
					return false;
				}
			}
			return true;
		}

		function updatePurchaseTotal()
		{
			var totalPurchase = parseFloat( "0.00" );
			var activityCount = parseInt( $("#activitycount").val() );
			var rege = /^\d+\.?\d{0,6}$/;
			var Ok;

			for ( var i = 1; i <= activityCount; i++ )
			{
				// Check each checkbox and sum the total
				if ( $("#activity" + i).is(":checked") == true )
				{
					Ok = rege.test( $("#amount" + i).val() );
					if ( ! Ok ) 
					{
						inlineMsg("amount" + i,'<strong>Invalid Format: </strong>The amount must be numeric.',8,"amount" + i);
						// uncheck the pick
						//$("#activity" + i).removeAttr('checked');
					}
					else
					{
						// format the amount
						$("#amount" + i).val(  parseFloat( $("#amount" + i).val() ).toFixed(2) );

						// Add to the total
						totalPurchase += parseFloat( $("#amount" + i).val() );
					}
				}
			}

			$("#purchasetotaldisplay").html( totalPurchase.toFixed(2) );
		}

		function validatePrice( AmountFieldId )
		{
			var rege = /^\d+\.?\d{0,6}$/;
			var Ok = rege.test( $("#" + AmountFieldId).val() );
			if ( ! Ok ) 
			{
				//$("#" + AmountFieldId).focus();
				inlineMsg(AmountFieldId,'<strong>Invalid Format: </strong>The amount must be numeric.',8,AmountFieldId);
    				//return false;
			}
			else
			{
				// format the amount
				$("#" + AmountFieldId).val(  parseFloat( $("#" + AmountFieldId).val() ).toFixed(2) );
			}

			// Re do the totals
			updatePurchaseTotal();
		}

		function completePurchase()
		{
			var totalPurchase = parseFloat( "0.00" );
			var accountBalance = parseFloat( $("#availablebalance").val() );
			var activityCount = parseInt( $("#activitycount").val() );
			var activitiesChecked = 0;
			var errorCount = 0;
			var rege = /^\d+\.?\d{0,6}$/;
			var Ok;

			if ( accountBalance == 0 )
			{
				alert( "There is nothing on their account.\nPlace some funds into their account and try again." );
				return;
			}

			for ( var i = 1; i <= activityCount; i++ )
			{
				// Check each checkbox and sum the total
				if ( $("#activity" + i).is(":checked") == true )
				{
					Ok = rege.test( $("#amount" + i).val() );
					if ( ! Ok ) 
					{
						inlineMsg("amount" + i,'<strong>Invalid Format: </strong>The amount must be numeric.',8,"amount" + i);
						errorCount++;
					}
					else
					{
						// format the amount
						$("#amount" + i).val(  parseFloat( $("#amount" + i).val() ).toFixed(2) );

						// Add to the total
						totalPurchase += parseFloat( $("#amount" + i).val() );

						activitiesChecked++;
					}
				}
			}

			if ( activitiesChecked == 0 )
			{
				alert( "Please select at least one activity.\nThen try your purchase again." );
				return;
			}
			
			if ( errorCount == 0 )
			{
				if ( totalPurchase > accountBalance )
				{
					//alert( totalPurchase );
					//alert( accountBalance );
					alert( "The purchase total is greater than their account balance." );
					return;
				}

				// everything is OK so submit the purchase
				//alert( "OK to submit" );
				document.frmPurchase.submit();
			}
		}

		function resetPage()
		{
			location.href = "dropinregistration.asp";
		}

		$(document).ready(function(){
			$(':input:visible:enabled:first').focus(); 
		});

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->

<div id="content">
	<div id="centercontent">

		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Drop In Registration</strong></font><br />
		</p>
		<!--END: PAGE TITLE-->

		<form name="frmUserIdScan" method="post" action="dropinregistration.asp" onsubmit="return validateForm()">
			<fieldset id="scanentryfieldset">
				Scan the barcode on the ID card<br /><br />
				<input type="text" name="useridscan" id="useridscan" size="9" maxlength="9" placeholder="ID Number" value="<%= iUserIdScan %>" /> &nbsp; 
				<input type="button" class="button" value="Search" onclick="validateSearch();" /> &nbsp; &nbsp;
				<input type="button" class="button" value="Reset Page" onclick="resetPage();" />
			</fieldset>
		</form>

<%			
		If iUserIdScan <> "" Then		
			'response.write "iUserIdScan = " & iUserIdScan & "<br /><br />"
%>	
			<form name="frmPurchase" method="post" action="dropinprocessing.asp">

<%				
				bUserIsValid = ShowUserInfo( iUserIdScan, bRegistrationBlocked )

				If bUserIsValid Then	

					ShowAvailableActivities session("orgid"), bRegistrationBlocked

				End If	
%>
			</form>
<%
		End If						
%>			
	</div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<!--#Include file="class_global_functions.asp"-->  


<%

'--------------------------------------------------------------------------------------------------
' boolean ShowUserInfo( iUserIdScan, bRegistrationBlocked )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( ByVal iUserIdScan, ByRef bRegistrationBlocked )
	Dim sSql, oRs, sPhotoPath

	bRegistrationBlocked = False  

	sPhotoPath = "../images/MembershipCard_Photos/users/" & session("orgid") & "_"

	sSql = "SELECT userid, orgid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & "ISNULL(useraddress,'') AS useraddress, ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, "
	sSql = sSql & "ISNULL(userzip,'') AS userzip, ISNULL(useremail,'') AS useremail, ISNULL(userhomephone,'') AS userhomephone, "
	sSql = sSql & "ISNULL(residenttype,'N') AS residenttype, ISNULL(accountbalance,0.00) AS accountbalance, registrationblocked, "
	sSql = sSql & "blockeddate, blockedadminid, blockedexternalnote, blockedinternalnote, card_pic_uploaded "
	sSql = sSql & "FROM egov_users "
	sSql = sSql & "WHERE userid = " & iUserIdScan
	sSql = sSql & "AND orgid = " & session("orgid")
	sSql = sSql & "AND isdeleted = 0"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "<font size=""+1""><strong>Purchaser Information</strong></font><br />"

		response.write vbcrlf & "<table id=""purchaserinfo"" cellspacing=""0"" cellpadding=""2"" border=""0"">"

		response.write vbcrlf & "<tr>"
		response.write "<th>Photo</th><th>Details</th><th>Account<br />Balance</th>"
		response.write vbcrlf & "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td id=""purchaserphoto"" align=""center"">"
		If oRs("card_pic_uploaded") Then
			response.write "<img border=""0"" src=""" & sPhotoPath & oRs("userid") & ".jpg"" height=""200"" width=""266"" />"
		Else
			response.write "<span id=""photounavailable"">Photo Unavailable</span>"
		End If 
		response.write "</td>"

		response.write "<td id=""purchaserdetails"">"
		response.write "<span id=""purchasername"">" & oRs("userfname") & " " & oRs("userlname") & "</span><br /><br />"
		response.write oRs("useraddress") & "<br />"
		response.write oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "<br />"
		response.write oRs("useremail") & "<br />"
		response.write FormatPhoneNumber( oRs("userhomephone") ) & "<br /><br />"
		' Link to add to account
		response.write "<input type=""button"" class=""button"" onclick=""location.href='../dirs/citizen_account_history.asp?u=" & oRs("userid") & "'"" value=""Add Funds To My Account"" />"

		' Handle blocked
		If oRs("registrationblocked") Then
			bRegistrationBlocked = True 
			response.write vbcrlf & "<br /><br /><span id=""warningmsg""> *** Registration Blocked *** </span><br /><br />"
			response.write vbcrlf & "Date: &nbsp;" & oRs("blockeddate") & "<br />"
			response.write vbcrlf & "By: &nbsp;" & GetAdminName( oRs("blockedadminid") ) & "<br />"
			response.write vbcrlf & "Internal Note: &nbsp;" & oRs.Fields("blockedinternalnote") & "<br />"
			response.write vbcrlf & "External Note: &nbsp;" & oRs.Fields("blockedexternalnote") & "<br />"
		End If 
		response.write "</td>"

		response.write "<td id=""accountbalance"">"
		response.write FormatNumber(oRs("accountbalance"),2,,,0) 
		response.write "<input type=""hidden"" id=""availablebalance"" name=""availablebalance"" value=""" & oRs("accountbalance") & """ />"
		response.write "</td>"

		response.write vbcrlf & "</tr>"
		response.write vbcrlf & "</table>"
		response.write "<input type=""hidden"" id=""userid"" name=""userid"" value=""" & iUserIdScan & """ />"
		ShowUserInfo = True 
	Else
		response.write "<p id=""idnotfound"">No information could be found for the ID Number entered. Please check the number and try again.</p>"
		ShowUserInfo = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string getWeekDayNameForOrg( )
'--------------------------------------------------------------------------------------------------
Function getWeekDayNameForOrg( )
	Dim sSql, oRs, sWeekDayName

	sWeekDayName = ""

	sSql = "SELECT dbo.GetWeekDayNameOfDate( dbo.GetLocalDate( " & Session("orgid") & ", getdate() ) ) AS dayname "
	sSql = sSql & "FROM organizations WHERE orgid = " & Session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sWeekDayName = oRs("dayname" )
	End If 

	oRs.Close
	Set oRs = Nothing 

	getWeekDayNameForOrg = sWeekDayName

End Function 


'--------------------------------------------------------------------------------------------------
' showAvailableActivities iOrgID, bRegistrationBlocked
'--------------------------------------------------------------------------------------------------
Sub showAvailableActivities( ByVal iOrgID, ByVal bRegistrationBlocked )
	Dim sSql, oRs, iActivityCount, iMinAge, iMaxAge, sBtnState, sOrgID

	iActivityCount = clng(0)
	sOrgID         = 0
	sBtnState      = ""

	if iOrgID <> "" then
		sOrgID = clng(iOrgID)
	end if

	if bRegistrationBlocked then
  		sBtnState = " disabled=""disabled"" "
	end if

	sSql = "SELECT C.classid, "
	sSql = sSql & " T.timeid, "
	sSql = sSql & " T.activityno, "
	sSql = sSql & " C.classname, "
	sSql = sSql & " ISNULL(D.starttime,'') AS starttime, "
	sSql = sSql & " ISNULL(D.endtime,'') AS endtime, "
	sSql = sSql & " P.pricetypeid, "
	sSql = sSql & " PT.accountid, "
	sSql = sSql & " PT.amount, "
	sSql = sSql & " ISNULL(C.locationid,0) AS locationid, "
	sSql = sSql & " ISNULL(C.minage,0.0) AS minage, "
	sSql = sSql & " ISNULL(C.maxage,99.9) AS maxage "
	sSql = sSql & " FROM egov_class C, "
	sSql = sSql & " egov_class_status S, "
	sSql = sSql & " egov_class_time T, "
	sSql = sSql & " egov_class_time_days D, "
	sSql = sSql & " egov_class_pricetype_price PT, "
	sSql = sSql & " egov_price_types P "
	sSql = sSql & " WHERE C.statusid = S.statusid "
	sSql = sSql & " AND C.classid = T.classid "
	sSql = sSql & " AND T.timeid = D.timeid "
	sSql = sSql & " AND C.classid = PT.classid "
	sSql = sSql & " AND PT.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND C.orgid = " & sOrgID
	sSql = sSql & " AND C.orgid = P.orgid "
	'sSql = sSql & " AND dbo.GetLocalDate( " & sOrgID & ", getdate( ) ) >= C.startdate "
	'sSql = sSql & " AND dbo.GetLocalDate( " & sOrgID & ", getdate() ) <= C.enddate "
	sSql = sSql & " AND datediff(day,dbo.GetLocalDate(" & sOrgID & ", getdate()), C.startdate) <= 0 "
	sSql = sSql & " AND datediff(day,dbo.GetLocalDate(" & sOrgID & ", getdate()), C.enddate) >= 0 "
	sSql = sSql & " AND S.iscancelled = 0 "
	sSql = sSql & " AND T.iscanceled = 0 "
	sSql = sSql & " AND D." & getWeekDayNameForOrg() & " = 1 "
	sSql = sSql & " AND P.isdropin = 1 "
	sSql = sSql & " ORDER BY C.classname, D.starttime, T.activityno"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write "<font size=""+1""><strong>Available Activities</strong></font><br />" & vbcrlf
		response.write "<table id=""activityinfo"" cellspacing=""0"" cellpadding=""2"" border=""0"">" & vbcrlf
		response.write "<tr>" & vbcrlf
		response.write "<th colspan=""5"">Activity</th>" & vbcrlf
  		response.write "<th>Price</th>" & vbcrlf
		response.write "</tr>" & vbcrlf

		Do While Not oRs.EOF

   			iActivityCount = iActivityCount + 1

     			response.write "<tr"

			If iActivityCount Mod 2 = 0 Then
				response.write " class=""alt_row"""
			End If 

			response.write ">" & vbcrlf
			response.write "<td align=""center"">"
			response.write "<input type=""checkbox"" id=""activity"  & iActivityCount & """ name=""activity"    & iActivityCount & """ onclick=""updatePurchaseTotal();"" />"
			response.write "<input type=""hidden"" id=""classid"     & iActivityCount & """ name=""classid"     & iActivityCount & """ value=""" & oRs("classid") & """ />"
			response.write "<input type=""hidden"" id=""timeid"      & iActivityCount & """ name=""timeid"      & iActivityCount & """ value=""" & oRs("timeid") & """ />"
			response.write "<input type=""hidden"" id=""pricetypeid" & iActivityCount & """ name=""pricetypeid" & iActivityCount & """ value=""" & oRs("pricetypeid") & """ />"
			response.write "<input type=""hidden"" id=""accountid"   & iActivityCount & """ name=""accountid"   & iActivityCount & """ value=""" & oRs("accountid") & """ />"
			response.write "</td>"
			response.write "<td align=""center"">" & oRs("activityno")
			response.write "</td>"
			response.write "<td class=""availableclassname"">" & oRs("classname")

			iMinAge = CDbl(oRs("minage"))
			iMaxAge = CDbl(oRs("maxage"))

			If iMinAge > CDbl(0.0) And iMaxAge < CDbl(99.9) Then
				response.write " (Ages " & iMinAge & " - " & iMaxAge & ")"
			Else
				If iMinAge > CDbl(0.0) Then
					response.write " (Ages " & iMinAge & " and above)"
				End If 

				If iMaxAge < CDbl(99.9) Then
					response.write " (Ages " & iMaxAge & " and below)"
				End If 
			End If 

			response.write "</td>"
			response.write "<td align=""center"">"
			response.write GetLocationName(  oRs("locationid") )
			response.write "</td>"
			response.write "<td align=""center"">"

			If oRs("starttime") <> "" Then 
				response.write oRs("starttime") & " &ndash; " & oRs("endtime")
			Else
				response.write "&nbsp;"
			End If 

			response.write "</td>"
			response.write "<td align=""center"">" 
			response.write "<input type=""text"" id=""amount" & iActivityCount & """ name=""amount" & iActivityCount & """ value=""" & FormatNumber(oRs("amount"),2,,,0) & """ size=""6"" maxlength=""6"" onchange=""validatePrice( 'amount" & iActivityCount & "' );"" />"
			'response.write "&nbsp;" & FormatCurrency(oRs("amount"),2)
			response.write "</td>" & vbcrlf
			response.write "</tr>" & vbcrlf

			oRs.MoveNext 
		Loop 

		' Purchase total Row
		response.write vbcrlf & "<tr id=""purchasetotalrow"">"
		response.write "<td align=""right"" colspan=""5"">Purchase Total</td>"
		response.write "<td id=""purchasetotal"" align=""center""><span id=""purchasetotaldisplay"">0.00</span></td>"
		response.write vbcrlf & "</tr>"
		
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<input type=""hidden"" id=""activitycount"" name=""activitycount"" value=""" & iActivityCount & """ />"

		' Purchase Notes
		response.write vbcrlf & "<strong>Purchase Notes:</strong><br /><textarea class=""purchasenotes"" name=""purchasenotes"" id=""dropinpurchasenotes""></textarea><br />"

		' Cpmplete Purchase button
		response.write vbcrlf & "<div id=""purchasecomplete"">"
		response.write vbcrlf & "<input type=""button"" class=""button"" id=""purchasebtn"" name=""purchasebtn"" value=""Complete Purchase"" onclick=""completePurchase();""" & sBtnState & " />"
		response.write vbcrlf & "</div>"
	Else
		response.write vbcrlf & "<p id=""activitynotfound""No Activities could be found.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub

%>

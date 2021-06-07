<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="facility_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/17/06  John Stullenberger - INITIAL VERSION
' 1.1  10/10/08  David Boyer - Added "isFacilityAvail" check
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iFacilityId, sLodgeName

iFacilityId = CLng(request("L"))
sLodgeName = getFacilityName( iFacilityId )  ' in facility_global_functions.asp

'Check to see if this facility has been reserved while user has been filling out reservation form for same facility.
 if request("success") <> "NA" then
    lcl_facility_avail = isFacilityAvail("", request("checkindate"), request("checkoutdate"), request("selTimePartID"), request("L"), request("D"))

    if Not lcl_facility_avail then
       response.redirect "facility_reservation.asp?L=" & iFacilityId & "&D=" & request("D") & "&TP=" & request("selTimePartID") & "&success=NA"
    end if
 end if
%>
<html lang="en">
<head>
	<meta charset="UTF-8">
<%
   if iorgid = 7 then
      lcl_title = sOrgName
   else
      lcl_title = "E-Gov Services " & sOrgName
   end if

   response.write "<title>" & lcl_title & "</title>" & vbcrlf
%>
	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="facility_styles.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />

	<script src="../scripts/modules.js"></script>
	<script src="../scripts/easyform.js"></script>

	<script>
	<!--

		function openwindow(sURL) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			window.open(sURL);
		}

		function view_waivers( surl )
		{
			// CHANGE FORM'S ACTION URL AND SUBMIT
			document.frmReserver.action = surl;
			document.frmReserver.target = '_NEW';
			document.frmReserver.submit();
		}

		function view_form()
		{
			// CHANGE FORM'S ACTION URL AND SUBMIT
			//document.frmReserver.action = '../recreation_payments/verisign_form.asp';
			document.frmReserver.action = '<%=Application("PAYMENTURL")%>/<%=sorgVirtualSiteName%>/recreation_payments/VERISIGN_FORM.ASP';
			
			document.frmReserver.target = '_self';
			document.frmReserver.submit();
		}

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->
<%	RegisteredUserDisplay( "../" ) %>

<!--<form name="frmReserver" action="../recreation_payments/verisign_form.asp" method="POST">-->
<form name="frmReserver" id="frmAvail" action="<%=Application("PAYMENTURL") & "/" & sorgVirtualSiteName & "/recreation_payments/VERISIGN_FORM.ASP"%>" method="POST">

	<input type="hidden" name="ITEM_NAME" value="FACILITY RESERVATION" />
	<input type="hidden" name="ITEM_NUMBER" value="F300" />
	<input type="hidden" name="iPAYMENT_MODULE" value="3" />
	<input type="hidden" name="amount" value="<%=request("amounttotal")%>" />

	<!--REQUIRED INFORMATION-->
	<input type="hidden" name="checkintime" value="<%=request("checkintime")%>" />
	<input type="hidden" name="checkindate" value="<%=request("checkindate")%>" />
	<input type="hidden" name="checkouttime" value="<%=request("checkouttime")%>" />
	<input type="hidden" name="checkoutdate" value="<%=request("checkoutdate")%>" />
	<input type="hidden" name="lesseeid" value="<%=request.cookies("userid")%>" />
	<input type="hidden" name="timepartid" value="<%=request("selTimePartID")%>" />
	<input type="hidden" name="facilityid" value="<%= iFacilityId %>" />
	<input type="hidden" name="lodgename" value="<%= sLodgeName %>" />
	<input type="hidden" name="paymenttype" value="1" />
	<input type="hidden" name="paymentlocation" value="3" />
	<input type="hidden" name="iuserid" value="<%=request.cookies("userid") %>" />
	<input type="hidden" name="D" value="<%=request("D")%>" />
	<input type="hidden" name="S" value="<%=session.sessionid%>" />
	<%
	  for each oField in request.form
		 if left(oField,7) = "custom_" then
				response.write "<input type=""hidden"" name=""" & oField & """ value=""" & request(oField) & """ />" & vbcrlf
		 end if
	  next
	%>
	<!--REQUIRED INFORMATION-->

	<!--RESERVATION DETAILS-->
	<%
	  if request("success") <> "" then
		 lcl_message = getSuccessMessage(request("success"))

		 if lcl_message <> "" then
			response.write "<div align=""right"" style=""width: 750px"">" & lcl_message & "</div>" & vbcrlf
		 end if
	  end if
	%>
	<div class="reserveformtitle">Reservation Details</div>
	<div class="reserveforminputarea"><% showReservationDetails sLodgeName %></div>

	<!--TERM/CONDITIONS AND WAIVER DOWNLOAD-->
	<div class="reserveformtitle">Terms/Conditions and Waiver Downloads</div>
	<div class="reserveforminputarea">

	<% 
		DisplayTerms CLng(request("L"))
	%>

	<%	If FacilityHasWaivers() Then	%>
			<p>
				<strong>Important!</strong> You must download several forms to print, sign, and bring with you when picking up the key.  
				You will need a PDF viewer in order to view and print these documents.  You can download the free Adobe Reader plugin 
				by clicking the link below.
			</p>
	<%	End If		%>

		<p><% ListWaivers %></p>

	</div>

	<!--TOTAL COSTS-->
	<div class="reserveformtitle">Totals Costs</div>
	<div class="reserveforminputarea">
	<% 
		GetPaymentDetails 
	%>
	</div>

<!--MAKE PURCHASE-->
<%
  'If the reservation is no longer available to be reserved then hide the "continue" button and show the "return" button.
   If request("success") = "NA" Then 
      lcl_purchasetitle = "Date/Time Period is no longer available."

      lcl_reserveformtext = "The button below will return you to the ""Check Availability and Reserve"" screen for this facility.  "
      lcl_reserveformtext = lcl_reserveformtext & "Although this time slot was available when you began your reservation, it has "
      lcl_reserveformtext = lcl_reserveformtext & "since been reserved and is no longer available.  We apologize for any inconvenience. "

      lcl_purchasebutton_value   = "Return to Facility Availability"
      lcl_purchasebutton_onclick = "location.href='facility_availability.asp?L=" & request("L") & "&Y=" & year(request("D")) & "&M=" & month(request("D")) & "';"
   Else 
      lcl_purchasetitle = "Make Purchase"

      lcl_reserveformtext = "The button below will take you to a credit card processing screen.  Your reservation is not complete "
      lcl_reserveformtext = lcl_reserveformtext & " until your payment is processed and you receive an online receipt."

      lcl_purchasebutton_value   = "Continue with Reservation"
      lcl_purchasebutton_onclick = "if (validateForm('frmReserver')) {view_form();}"
   End If 

   response.write "<div class=""reserveformtitle"">" & lcl_purchasetitle & "</div>" & vbcrlf
   response.write "<div class=""reserveforminputarea"" align=""center"">" & vbcrlf

   If lcl_message <> "" Then 
      response.write "<div align=""center"" style=""width: 750px"">" & lcl_message & "</div><br />" & vbcrlf
   End If 

   response.write "  <p class=""reserveformtext"">" & lcl_reserveformtext & "</p>" & vbcrlf
   response.write "  <input type=""button"" value=""" & lcl_purchasebutton_value & """ class=""facilitybutton"" onclick=""" & lcl_purchasebutton_onclick & """ />" & vbcrlf
   response.write "</div>" & vbcrlf
%>
</form>

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'------------------------------------------------------------------------------
Sub DisplayTerms( ByVal iFacilityId )
	Dim sSql, oRs

	sSql = "SELECT termid, termdescription FROM egov_recreation_Terms WHERE facilityid = " & iFacilityId & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	if not oRs.eof then

		response.write "<strong>Important! You must read and agree to each of the terms/conditions below.</strong>" & vbcrlf

		Do While Not oRs.EOF
			response.write "<p>" & vbcrlf
			response.write "  <span style=""color: #ff0000"">*&nbsp;</span>" & vbcrlf
			response.write "  <input value=""on"" type=""checkbox"" name=""term" & oRs("termid") & """ /> " & oRs("termdescription") & vbcrlf
			response.write "</p>" & vbcrlf
			response.write "<input type=""hidden"" name=""ef:term" & oRs("termid") & "-checkbox/req"" value=""" & Left(clearHTMLTags( oRs("termdescription") ),50) & "..."" />" & vbcrlf
			
			oRs.MoveNext
		Loop 

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub showReservationDetails( ByVal sLodgeName )

	response.write "<strong>Facility Name</strong>: " & sLodgeName & "<br /><br />"
	response.write "<strong>Check In:</strong> " & request("checkindate") & " " & request("checkintime") & "<br /><br />"
	'response.write "<strong>Check In Date</strong>: " & request("checkindate") & "<br />"
	response.write "<strong>Check Out:</strong> "& request("checkoutdate") & " " & request("checkouttime") & "<br />"
	'response.write "<strong>Check Out Date</strong>: " & request("checkoutdate") & "<br />"

End Sub 


'------------------------------------------------------------------------------
Sub GetPaymentDetails()

	response.write "<strong>Facility Reservation Cost:</strong> " & request("reservetotal") & "<br />" & vbcrlf

	If LCase(request("keydeposit")) = "on" Then 
  		response.write "<strong>Key Deposit Charge:</strong> " & request("reservetotal") & "<br />" & vbcrlf
	End If 

	response.write "<strong>Total:</strong> " & request("amounttotal") & "<br />" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub ListWaivers()
	Dim iWaiverCount, sList

	iWaiverCount = 0
	sList        = ""

	For Each oField In request.Form 
		If Left(oField,11) = "chkwaivers_" Then 
			iWaiverCount = iWaiverCount + 1
			If sList = "" Then 
				sList = sList & request(oField)
			Else 
				sList = sList & "X" & request(oField)
			End If 
		End If 
	Next 

	If iWaiverCount <> 0 Then 
		response.write "<p>" & vbcrlf
		response.write "<input style=""text-align: center;"" type=""button"" value=""Click to download required reservation forms in PDF format"" onClick=""view_waivers('display_waiver.aspx?MASK=" & sList & "');"" >" & vbcrlf
		response.write "</p>" & vbcrlf

		'Link to ADOBE
		response.write "<p>" & vbcrlf
		response.write "<a href=""http://www.adobe.com/products/acrobat/readstep2.html"" target=""_blank"" title=""Get Adobe Acrobat Reader Plug-in Here""><img border=""0"" src=""../images/adreader.gif"">""Get Adobe Reader.""</a>" & vbcrlf
		response.write "</p>" & vbcrlf
	Else 
		response.write "<p>No waivers required.</p>" & vbcrlf
	End If 
	
End Sub


'------------------------------------------------------------------------------
' boolean FacilityHasWaivers()
'------------------------------------------------------------------------------
Function FacilityHasWaivers()
	Dim iWaiverCount, sList

	iWaiverCount = 0
	sList        = ""

	For Each oField In request.Form 
		If Left(oField,11) = "chkwaivers_" Then 
			iWaiverCount = iWaiverCount + 1
			If sList = "" Then 
				sList = sList & request(oField)
			Else 
				sList = sList & "X" & request(oField)
			End If 
		End If 
	Next 

	If iWaiverCount <> 0 Then 
		FacilityHasWaivers = True 
	Else
		FacilityHasWaivers = False 
	End If 

End Function 



%>

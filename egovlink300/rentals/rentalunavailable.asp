<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalunavailable.asp
' AUTHOR: Steve Loar
' CREATED: 02/04/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Message given when rental is no longer available during the reservation process.
'
' MODIFICATION HISTORY
' 1.0   02/04/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, sPhoneNumber, iReservationTempId, sRentalName

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

If request("rti") = "" Then
	response.redirect "rentalcategories.asp"
Else 
	If Not IsNumeric(request("rti")) Then
		response.redirect "rentalcategories.asp"
	Else 
		iReservationTempId = CLng(request("rti"))
	End If 
End If 

sPhoneNumber = GetRentalSupervisorPhoneByRTI( iReservationTempId )

If sPhoneNumber = "" Then
	' Use the City Default Phone Number 
	' sDefaultPhone is set in include_top_functions.asp
	sPhoneNumber = FormatPhoneNumber( sDefaultPhone )	' in common.asp
End If 

sPhoneNumber = Trim(sPhoneNumber)

sRentalName = GetRentalNameByRTI( iReservationTempId )

ClearTempReservation iReservationTempId, iOrgId


%>

<html>
<head>

	<title><%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />


</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<!-- Show the Rental Category navagation -->
<%	DisplayCategoryMenu	iorgid	%>

<p id="limitationmessage">
	We are sorry but the date and time that you are trying to reserve <%=sRentalName%> are no longer available.  
	If you would like help with your reservation, please call us at <%=sPhoneNumber%>.
</p>


<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------



%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalunavailable.asp
' AUTHOR: Steve Loar
' CREATED: 10/08/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page is displayed when the final reservation availability check fails. 
'				Someone has beaten them to the day and time they wanted.
'
' MODIFICATION HISTORY
' 1.0   10/08/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, sRentalName

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create edit rentals", sLevel	' In common.asp

'iRentalId = CLng(request("rentalid"))
iRentalID = 0

sRentalName = GetRentalName( iRentalId )		' in rentalscommonfunctions.asp

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Rental Unavailable</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p class="limitationmessage">
				We are sorry but our final check of availability shows that the date and time that you are 
				trying to reserve the <%=sRentalName%> are no longer available.  
			</p>
			<p class="limitationmessage">
				To try another date or time <a href="rentalsearch.asp">click here to start your reservation again</a>.
			</p>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>
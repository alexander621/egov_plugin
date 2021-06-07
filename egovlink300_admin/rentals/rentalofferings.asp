<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalslist.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals. From here you can create or edit rentals
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
' 1.1	03/24/2011	Steve Loar - hide deactivated rentals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iCategoryId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create simple reservations", sLevel	' In common.asp

iCategoryId = CLng(request("categoryid"))


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="Javascript">
	<!--
		
		function goBack()
		{
			location.href='rentalcategoryselection.asp';
		}

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
				<font size="+1"><strong>Make Simple Reservations</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p>
				<input type="button" class="button" value="<< Back to Categories" onclick="goBack()" />
			</p>


			<!-- Display the category details -->
			<%	DisplayCategoryDetails iCategoryId		%>

			<!-- List out the details for the rentals in this category -->
			<%	ShowRentalsList iCategoryId		%>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayCategoryDetails iCategoryId
'--------------------------------------------------------------------------------------------------
Sub DisplayCategoryDetails( ByVal iCategoryId )
	Dim sSql, oRs

	sSql = "SELECT recreationcategoryid, ISNULL(categorytitle,'') AS categorytitle, "
	sSql = sSql & "ISNULL(categorydescription,'') AS categorydescription "
	sSql = sSql & "FROM egov_recreation_categories "
	sSql = sSql & "WHERE recreationcategoryid = " & iCategoryId & " AND orgid = " & session("orgid") 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<table class=""availablecategory"" id=""availablerentalcategory"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td class=""spacerrow"" colspan=""2"">&nbsp;</td></tr>"

		response.write "</td><td valign=""top"" class=""availabledescription"">"
		response.write "<p><span class=""availableschedulerentalname"">"
		response.write oRs("categorytitle")
		response.write "</span></p>"

		response.write "<p>" & oRs("categorydescription") & "</p>"

		If GetRentalsInCategoryCount( oRs("recreationcategoryid") ) > CLng(1) Then 
			response.write "<div class=""checkbutton"">"
			response.write "<input type=""button"" class=""button"" value=""Check Availability On All " & oRs("categorytitle") & """ onclick=""location.href='rentalavailability.asp?cid=" & iCategoryId & "';"" />"
			response.write "</div>"
		End If 

		response.write "</td></tr></table>"
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetRentalsInCategoryCount( iRecreationCategoryId )
'--------------------------------------------------------------------------------------------------
Function GetRentalsInCategoryCount( ByVal iRecreationCategoryId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(rentalid) AS hits FROM egov_rentals_to_categories WHERE recreationcategoryid = " & iRecreationCategoryId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRentalsInCategoryCount = CLng(oRs("hits"))
	Else
		GetRentalsInCategoryCount = CLng(0)
	End If 

	oRs.Close 
	Set oRs = Nothing  

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowRentalsList iCategoryId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalsList( ByVal iCategoryId )
	Dim sSql, oRs, sMainImg, iRecCount

	iRecCount = 0

	'  AND R.publiccanview = 1
	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, ISNULL(R.width,'') AS width, ISNULL(R.length,'') AS length, "
	sSql = sSql & "ISNULL(R.capacity,'') AS capacity, R.publiccanreserve, usehtmlonlongdesc, "
	sSql = sSql & "ISNULL(R.description,'') AS description, ISNULL(R.iconimageurl,'') AS iconimageurl "
	sSql = sSql & "FROM egov_rentals R, egov_rentals_to_categories C, egov_class_location L "
	sSql = sSql & "WHERE C.rentalid = R.rentalid AND R.isdeactivated = 0 AND R.locationid = L.locationId "
	sSql = sSql & "AND C.recreationcategoryid = " & iCategoryId
	sSql = sSql & "ORDER BY L.name, R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iRecCount = iRecCount + 1
		response.write "<table class=""availablerentals"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		'response.write "<tr><td class=""spacerrow"" colspan=""2"">&nbsp;</td></tr>"

		response.write "<tr>"

		response.write "<td valign=""top"" align=""left"" width=""100%"" class=""availabledescription"">"

		response.write "<p><span class=""availableschedulerentalname"">"
		If oRs("locationname")  <> "" Then 
			response.write oRs("locationname") & " &ndash; " 
		End If 
		response.write oRs("rentalname")
		response.write "</span>"
		response.write " &nbsp; <input type=""button"" class=""button"" value=""Check Availability"" onclick=""location.href='rentalavailability.asp?rid=" & oRs("rentalid") & "&cid=" & iCategoryId & "';"" />"
		response.write "</p>"

		response.write "<p>"
		If oRs("usehtmlonlongdesc") Then 
			response.write oRs("description") 
		Else 
			response.write Replace(oRs("description"), Chr(10), "<br />")
		End If 
		response.write "</p>"

		If oRs("locationname")  <> "" Or oRs("width") <> "" Or oRs("capacity") <> "" Then 
			response.write vbcrlf & "<p>"
			If oRs("width") <> "" Then 
				response.write "<strong>Dimensions: </strong>" & oRs("width") & " x " & oRs("length") & "<br />"
			End If 
			If oRs("capacity") <> "" Then 
				response.write "<strong>Capacity: </strong>" & oRs("capacity") & "<br />"
			End If 
			response.write vbcrlf & "</p>"
		End If 

'		DisplayRentalDocuments oRs("rentalid")

		response.write "</td>"
		response.write "</tr>"

		response.write "</table>"
		oRs.MoveNext 
	Loop

	oRs.Close
	Set oRs = Nothing 

End Sub


%>

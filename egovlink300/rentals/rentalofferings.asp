<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalofferings.asp
' AUTHOR: Steve Loar
' CREATED: 01/14/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals in a category.
'
' MODIFICATION HISTORY
' 1.0   01/14/2010	Steve Loar - INITIAL VERSION
' 1.1	03/24/2011	Steve Loar - Hiding deactivated rentals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iCategoryId

If request("categoryid") = "" Then
	response.redirect "../rentals/rentalcategories.asp"
Else
	If Not IsNumeric(request("categoryid")) Then
		response.redirect "../rentals/rentalcategories.asp"
	Else 
		iCategoryId = 0
		on error resume next
		iCategoryId = CLng(request("categoryid"))
		on error goto 0

		if iCategoryId = 0 then response.redirect "../rentals/rentalcategories.asp"
	End If 
End If

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

' Want to be sure that the category is viewable to the public. If not take them to the category list page'
If categoryIsNotViewable( iCategoryId ) Then 
	response.redirect "rentalcategories.asp"
End If 

%>

<html lang="en">
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<meta charset="UTF-8">

	<title><%=sTitle%></title>

	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalstyles.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />

	<script src="../scripts/mootools-12.js" type="text/javascript"></script> <!-- MOOTOOLS 1.2 BETA -->
	<script src="../scripts/mootools-12-more.js" type="text/javascript"></script> <!-- MOOTOOLS 1.2 BETA -->

	<script>
	<!--

		function imageSwap( sItem, iFullImgNo )
		{
				/*
			var item = $(sItem);
			var drop = $('fullimg'+ iFullImgNo );

			drop.removeEvents();
			drop.empty();


			var newImage = item.getElement('img');

			var a = newImage.clone();
			a.inject(drop);
			var b = drop.getElement('img');
			b.set('class', 'templateimg');
			*/
			newImage = document.getElementById('img' + sItem).src;
			document.getElementById('imgfullimg' + iFullImgNo).src = newImage;
		}

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<!--BEGIN: Page Top Display-->
<% 
	If OrgHasDisplay( iorgid, "rentalscategorypagetop" ) Then
		response.write vbcrlf & "<div id=""rentalscategorypagetop"">" & GetOrgDisplay( iOrgId, "rentalscategorypagetop" ) & "</div>"
	End If 
%>
<!--END: Page Top Display-->

<!-- Show the Rental Category navagation -->
<%	DisplayCategoryMenu	iorgid	%>


<!-- Display the category details -->
<%	DisplayCategoryDetails iCategoryId		%>


<!-- List out the details for the rentals in this category -->
<%	ShowRentalsList iCategoryId		%>


<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' void DisplayCategoryDetails iCategoryId
'--------------------------------------------------------------------------------------------------
Sub DisplayCategoryDetails( ByVal iCategoryId )
	Dim sSql, oRs

	sSql = "SELECT recreationcategoryid, ISNULL(categorytitle,'') AS categorytitle, "
	sSql = sSql & "ISNULL(categorydescription,'') AS categorydescription, ISNULL(imgurl,'') AS imgurl "
	sSql = sSql & "FROM egov_recreation_categories "
	sSql = sSql & "WHERE recreationcategoryid = '" & iCategoryId & "' AND orgid = '" & iorgid  & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<table class=""availablecategory"" id=""availablerentalcategory"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td class=""spacerrow"" colspan=""2"">&nbsp;</td></tr>"

		response.write "</td><td valign=""top"" class=""availabledescription"">"
		response.write "<p><span class=""schedulerentalname"">"
		response.write oRs("categorytitle")
		response.write "</span></p>"

		response.write "<p>" & oRs("categorydescription") & "</p>"

		If PublicCanReserveRentalsInCategory( oRs("recreationcategoryid") ) Then 
			If GetRentalsInCategoryCount( oRs("recreationcategoryid") ) > CLng(1) Then 
				response.write "<div class=""checkbutton"">"
				response.write "<input type=""button"" class=""button"" value=""Check Availability On All " & oRs("categorytitle") & """ onclick=""location.href='rentalavailability.asp?cid=" & iCategoryId & "';"" />"
				response.write "</div>"
			End If 
		End If

		response.write "</td></tr></table>"
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalsList iCategoryId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalsList( ByVal iCategoryId )
	Dim sSql, oRs, sMainImg, iRecCount

	iRecCount = 0

	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, ISNULL(R.width,'') AS width, ISNULL(R.length,'') AS length, "
	sSql = sSql & "ISNULL(R.capacity,'') AS capacity, R.publiccanreserve, usehtmlonlongdesc, "
	sSql = sSql & "ISNULL(R.description,'') AS description, ISNULL(R.iconimageurl,'') AS iconimageurl "
	sSql = sSql & "FROM egov_rentals R, egov_rentals_to_categories C, egov_class_location L "
	sSql = sSql & "WHERE C.rentalid = R.rentalid AND R.publiccanview = 1 AND R.locationid = L.locationId "
	sSql = sSql & "AND isdeactivated = 0 AND C.recreationcategoryid = '" & iCategoryId & "'"
	sSql = sSql & "ORDER BY L.name, R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iRecCount = iRecCount + 1
		response.write "<table class=""availablerentals"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td class=""spacerrow"" colspan=""2"">&nbsp;</td></tr>"

		response.write "<tr>"
		response.write "<td valign=""top"" align=""center"" class=""imgcell2"">"
		response.write "<div class=""fullimg"" id=""fullimg" & iRecCount & """>" 
		If oRs("iconimageurl") <> "" Then 
			response.write "<img id=""imgfullimg" & iRecCount & """ src=""" & replace(oRs("iconimageurl"),"http://www.egovlink.com","") & """ alt=""" & oRs("rentalname") & """ class=""templateimg"" />"
			sMainImg = oRs("iconimageurl")
		Else
			response.write "&nbsp;"
			sMainImg = ""
		End If 
		response.write "</div>"

		If PublicCanReserveRental( oRs("rentalid") ) then
			response.write "<div class=""checkbutton"">"
			response.write "<input type=""button"" class=""button"" value=""Check Availability"" onclick=""location.href='rentalavailability.asp?rid=" & oRs("rentalid") & "&cid=" & iCategoryId & "';"" />"
			response.write "<br /><br /><input type=""button"" class=""button"" value=""Browse Calendar for Availability"" onclick=""location.href='rentalcalendar.asp?rid=" & oRs("rentalid") & "';"" />"
			response.write "</div>"
		End If 

		response.write "</td>"
		response.write "<td valign=""top"" align=""left"" width=""100%"" class=""availabledescription"">"

		response.write "<p><span class=""schedulerentalname"">"
		If oRs("locationname")  <> "" Then 
			response.write oRs("locationname") & " &ndash; " 
		End If 
		response.write oRs("rentalname")
		response.write "</span></p>"

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

		DisplayRentalDocuments oRs("rentalid")

		response.write "</td>"
		response.write "</tr>"

		' Display any extra images
		DisplayExtraImages oRs("rentalid"), sMainImg, oRs("rentalname"), iRecCount

		response.write "</table>"
		oRs.MoveNext 
	Loop

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void DisplayExtraImages iRentalId 
'--------------------------------------------------------------------------------------------------
Sub DisplayExtraImages( ByVal iRentalId, ByVal sMainImg, ByVal sAltTag, ByVal iRecCount )
	Dim sSql, oRs, iImageCount

	iImageCount = 0

	sSql = "SELECT imageid, imageurl, alttag FROM egov_rentalimages "
	sSql = sSql & "WHERE rentalid = '" & iRentalId & "' ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr><td align=""center"" class=""rentalimgcell"" colspan=""2"">"
		response.write "<div class=""imgrow"">"

		response.write "<span class=""item"" id=""item" & iRecCount & "0"" onclick=""imageSwap( 'item" & iRecCount & "0', " & iRecCount & " );"">"
		response.write "<img id=""imgitem" & iRecCount & "0"" class=""thumb"" alt=""" & sAltTag & """ src=""" & replace(sMainImg,"http://www.egovlink.com","") & """ />"
		response.write "</span>"

		Do While Not oRs.EOF
			iImageCount = iImageCount + 1
			response.write "<span class=""item"" id=""item" & iRecCount & iImageCount & """ onclick=""imageSwap( 'item" & iRecCount & iImageCount & "', " & iRecCount & " );"">"
			response.write "<img id=""imgitem" & iRecCount & iImageCount & """ class=""thumb"" alt=""" & oRs("alttag") & """ src=""" & replace(oRs("imageurl"),"http://www.egovlink.com","") & """ />"
			response.write "</span>"
			oRs.MoveNext
		Loop
		response.write "<div class=""imageinstructions"">Click small images to see in larger view</div>"
		response.write "</div>"
		
		response.write "</td></tr>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean categoryIsNotViewable iCategoryId 
'--------------------------------------------------------------------------------------------------
Function categoryIsNotViewable( ByVal iCategoryId )
	Dim sSql, oRs 
	
	sSql = "SELECT hidefrompublic FROM egov_recreation_categories WHERE orgid = '" & iOrgId & "' AND recreationcategoryid = '" & iCategoryId & "'"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If oRs("hidefrompublic") Then 
			categoryIsNotViewable = True 
		Else
			categoryIsNotViewable = False 
		End If 
	Else
		categoryIsNotViewable = True 
	End If 
	
	oRs.Close
	Set oRs = Nothing 
	
End Function 


%>

<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: facilitycategory.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This was the facility_list.asp page, but has been altered to integrate with the
'				newer rentals feature.
'
' MODIFICATION HISTORY
' 1.0   01/17/2006	JOHN STULLENBERGER - INITIAL VERSION
' 2.0	01/13/2010	Steve Loar - Initial Version as facilitycategory.asp
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCategoryId, sTitle

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title><%=sTitle%></title>

	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../rentals/rentalstyles.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" href="facility_styles.css" />

	<script src="../scripts/modules.js"></script>
	<script src="../scripts/easyform.js"></script>

</head>

<!--#Include file="../include_top.asp"-->
<%	RegisteredUserDisplay( "../" ) %>

<!--BEGIN PAGE CONTENT-->
<p>
	<font class="pagetitle"><%=GetOrgFeatureName( "rentals" )%></font>
	<br />
</p>

<%	

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

' Page Top Display
If OrgHasDisplay( iorgid, "rentalscategorypagetop" ) Then
	response.write vbcrlf & "<div id=""rentalscategorypagetop"">" & GetOrgDisplay( iOrgId, "rentalscategorypagetop" ) & "</div>"
End If 


DisplayCategoryInformation iCategoryId, 1

DisplayFacility iCategoryId 

%>
<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../include_bottom.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
' DisplayFacility iCategoryId
'--------------------------------------------------------------------------------------------------
 Sub DisplayFacility( ByVal iCategoryId )
	Dim sSql, oRs

    ' GET SELECTED FACILITY INFORMATION
	sSql = "SELECT facilityid, facilityname, isviewable, isreservable, facilitytemplateid "
	sSql = sSql & "FROM egov_recreation_item_to_category "
	sSql = sSql & "WHERE categoryid = " & iCategoryId & " ORDER BY facilityname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Response.Write vbCrLf & "<!-- begin DisplayFacility-->"
	' DISPLAY FACILITY INFORMATION

	If Not oRs.EOF Then

		Response.Write vbCrLf & "<div class=""facilitylist"">"

		Do While Not oRs.EOF
			' WRITE TITLE
			Response.Write vbCrLf & "<div class=""facilityname"">" & oRs("facilityname") & "</div>"

			' WRITE LINK TO AVAILABILITY
			If oRs("isviewable") Then

				If oRs("isreservable") Then
					sMsg = "Check Availability and Reserve"
				Else
					sMsg = "Check Availability"
				End If
				Response.Write vbCrLf & "<div class=""reserve_link"" align=""left""><a href=""facility_availability.asp?L=" & oRs("facilityid") & """ class=""linkbutton"">" & sMsg &" </a></br></div>"
			End If

			' WRITE DESCRIPTION
			DisplayFacilityDetail oRs("facilityid"), oRs("facilitytemplateid")

			' WRITE LINK TO AVAILABILITY
			If oRs("isviewable") Then

				If oRs("isreservable") Then
					sMsg = "Check Availability and Reserve"
				Else
					sMsg = "Check Availability"
				End If
				Response.Write vbCrLf & "<div class=""reserve_link"" align=""right""><a href=""facility_availability.asp?L=" & oRs("facilityid") & """ class=""linkbutton"">" & sMsg &"</a></br></div>"
			End If

			oRs.MoveNext

		Loop

		Response.Write vbcrlf & "</div>"

	End If

		' CLOSE OBJECTS
		oRs.Close
		Set oRs = Nothing 
		Response.Write vbCrLf & "<!-- finish DisplayFacility-->"
 End Sub


'--------------------------------------------------------------------------------------------------
' DisplayCategoryInformation iCategoryId, blnShowBreadCrumbs
'--------------------------------------------------------------------------------------------------
Sub DisplayCategoryInformation( ByVal iCategoryId, ByVal blnShowBreadCrumbs )
	Dim sSql, oRs

    ' GET SELECT CATEGORY ROW
		sSql = "SELECT categorytitle, imgurl, categorysubtitle, categorydescription FROM egov_recreation_categories "
		sSql = sSql & "WHERE recreationcategoryid = " &  iCategoryId & " AND orgid = " & iorgid

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
		' DISPLAY BREADCRUMBS
			If blnShowBreadCrumbs = 1  Then
				Response.Write vbCrLf & "<p><div class=""subcategorymenu"">"
				Response.Write vbcrlf & "<a class=""subcategorymenu"" href=""../rentals/rentalcategories.asp"" >" & GetOrgFeatureName( "rentals" ) & "</a> | <a class=""subcategorymenu"" href=""facilitycategory.asp?categoryid=" & iCategoryId & """ >" & oRs("categorytitle") & "</a> "
				Response.Write vbCrLf & "</div></p>" 
			End If

			' DISPLAY CATEGORY INFORMATION
			Response.Write "<!--start display-->"
			Response.Write "<p><div class=""categorygroup"">"
        
			' WRITE PHOTO
			sImgURL = ""
			If oRs("imgurl") = "" OR ISNULL(oRs("imgurl"))  Then
				sImgURL = "images/park_category_default.jpg"
			Else
				sImgURL = replace(oRs("imgurl"),"http://www.egovlink.com","https://www.egovlink.com")
			End If

			' WRITE IMAGE LINK
			Response.Write "<img class=""categoryimage"" align=""left"" src=""" & sImgURL & """ />"

            ' WRITE TITLE
            Response.Write "<font class=""categorytitle"">" & oRs("categorytitle") & "<br /></font>" 
            ' WRITE SUBTITLE
			If oRs("categorysubtitle") <> "" Then
				Response.Write "<font class=""categorysubtitle"">" & oRs("categorysubtitle") & "</font><br /><br />" 
			End If
            ' WRITE DESCRIPTION
            Response.Write "<font class=""categorydescription"" >" & oRs("categorydescription") & "</font><br />"

			Response.Write vbCrLf & "</div></p>" 
	        Response.Write "<BR clear=""all""><!--end display-->"

        End If

        ' CLOSE OBJECTS
		oRs.Close
        Set oRs = Nothing 

 End Sub


'--------------------------------------------------------------------------------------------------
' DisplayFacilityDetail iFacilityID, iTemplateId
'--------------------------------------------------------------------------------------------------
Sub DisplayFacilityDetail( ByVal iFacilityID, ByVal iTemplateId )

	' GET FACILITY ELEMENTS
	Dim arrImgUrl(4)
	Dim arrText(4)
	For i = 1 to 4
		arrText(i) = GetText(iFacilityID,i)
		arrImgUrl(i) = GetImage(iFacilityID,i+4)
	Next 

	' DISPLAY SELECTED TEMPLATE
	response.write "<div id=""templatecontainer"">"
	
	Select Case iTemplateId

	Case 1
		response.write "<table class=""template"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr>"
		response.write "<td valign=""top"">" & arrImgUrl(1) & "</td>"
		response.write "<td colspan=""2"" valign=""top"">" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td class=""bottomrow"" colspan=""3"" valign=""top"" nowrap>"
		response.write arrImgUrl(2) & arrImgUrl(3) & arrImgUrl(4) 
		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 2
		response.write "<table class=""template"">"
		response.write "<tr>"
		response.write "<td valign=""top"">" & arrImgUrl(1) & "<br>" & arrImgUrl(2) & "</td>"
		response.write "<td valign=""top"">" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td class=""bottomrow"" colspan=""2"" valign=""top"" align=""center"" nowrap>"
		response.write arrImgUrl(3) & arrImgUrl(4) 
		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 3
		response.write "<table class=""template"">"
		response.write "<tr>"
		response.write "<td valign=""top"">" & arrImgUrl(1) & "</td>"
		response.write "<td valign=""top"" colspan=""2"">" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td valign=""top"">" & arrImgUrl(2) & "</td>"
		response.write "<td colspan=""2"" valign=""top"">" & arrText(2) & "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 4
		response.write "<table class=""template"">"
		response.write "<tr>"
		response.write "<td valign=""top"">" & arrImgUrl(1) 
		response.write "<br>" & arrText(2) & "</td>"
		response.write "<td  valign=""top"">" & arrText(1) 
		response.write "<br />" & arrImgUrl(2) & "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 5
		response.write "<table class=""template"">"
		response.write "<tr>"
		response.write "<td  valign=""top"">" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "</table>"
	End Select

	response.write "</div>"

End Sub 


'--------------------------------------------------------------------------------------------------
' string = GetImage(iFacilityId,iSequence)
'--------------------------------------------------------------------------------------------------
Function GetImage( ByVal iFacilityId, ByVal iSequence )
	Dim sSql, oRs, sReturnValue
	
	sReturnValue = " "

	sSql = "SELECT elementid, content, alt_tag "
	sSql = sSql & "FROM egov_facilityelements "
	sSql = sSql & "WHERE facilityid =" & iFacilityID & " AND sequence = " & iSequence 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If  oRs("content") <> "" Then
			sReturnValue = "<img class=""templateimg"" src=""" &  replace(oRs("content"),"http://www.egovlink.com","https://www.egovlink.com") & """ alt=""" & oRs("alt_tag") & """ title=""" &  oRs("alt_tag") & """ />"
		End If
	End If

	oRs.Close
	Set oRs = Nothing

	GetImage = sReturnValue 

End Function


'--------------------------------------------------------------------------------------------------
' string = GetText( iFacilityId, iSequence )
'--------------------------------------------------------------------------------------------------
Function GetText( ByVal iFacilityId, ByVal iSequence )
	Dim sSql, oRs, sReturnValue
	
	sReturnValue = " "
	
	sSql = "SELECT elementid, content FROM egov_facilityelements "
	sSql = sSql & "WHERE facilityid = " & iFacilityID & " AND sequence = " & iSequence 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If Not oRs.EOF Then
		sReturnValue = oRs("content") 	
	End If

	oRs.Close
	Set oRs = Nothing

	GetText = sReturnValue

End Function


%>


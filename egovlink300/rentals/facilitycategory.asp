<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
Dim iCategoryID

%>
<html>
<head>

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/easyform.js"></script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->
<p>
	<font class="pagetitle">Facility Reservations</font>
	<br />
</p>

<%	

RegisteredUserDisplay( "../" ) 

If request("categoryid") = "" Then
	response.redirect "../rentals/rentalcategories.asp"
Else
	If Not IsNumeric(request("categoryid")) Then
		response.redirect "../rentals/rentalcategories.asp"
	Else 
		iCategoryID = CLng(request("categoryid"))
	End If 
End If


' DISPLAY CATEGORY SELECTION
'If CLng(iCategoryID) = CLng(GetFirstCategory) Then
'	DisplaySubCategoryMenu iorgid, iCategoryID
'End If


'  DISPLAY ROOT CATEGORY DESCRIPTION
'If CLng(iCategoryID) = CLng(GetFirstCategory) Then
'	DisplayCategoryInformation icategoryid,0
'Else
'	DisplayCategoryInformation icategoryid,1
'End If


' LIST ALL THE facilities in this category
'If CLng(iCategoryID) = CLng(GetFirstCategory) Then
'	ListCategories iCategoryID, 1, 0, 0
'Else
	ListCategories iCategoryID, 1, 1, 1
'End If

%>
<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../include_bottom.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB DISPLAYFACILITY(SUBCATEGORYID)
'--------------------------------------------------------------------------------------------------
 Sub DisplayFacility( ByVal subcategoryid )
	Dim sSql, oRs

    ' GET SELECTED FACILITY INFORMATION
	sSql = "select * from egov_recreation_item_to_category where categoryid = " & subcategoryid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1


	Response.Write(vbCrLf & "<!-- begin DisplayFacility-->" & vbCrLf)
	' DISPLAY FACILITY INFORMATION
	If Not oRs.EOF Then
		Response.Write vbCrLf & "<div class=""facilitylist"">"
		Do While NOT oRs.EOF
			' WRITE TITLE
			Response.Write vbCrLf & "<div class=""facilityname"">" & oRs("facilityname") & "</div>" & vbCrLf

			' WRITE LINK TO AVAILABILITY
			If oRs("isviewable") Then

				If oRs("isreservable") Then
					sMsg = "Check Availability and Reserve"
				Else
					sMsg = "Check Availability"
				End If

				Response.Write vbCrLf & "<div class=""reserve_link"" align=""left""><a href=""facility_availability.asp?L=" & oRs("facilityid") & """>" & sMsg &" </a></br></div>" & vbCrLf
			End If

			' WRITE DESCRIPTION
			DisplayFacilityDetail oRs("facilityid"),oRs("facilitytemplateid")


			' WRITE LINK TO AVAILABILITY
			If oRs("isviewable") Then

				If oRs("isreservable") Then
					sMsg = "Check Availability and Reserve"
				Else
					sMsg = "Check Availability"
				End If

				Response.Write vbCrLf & "<div class=""reserve_link"" align=""right""><a href=""facility_availability.asp?L=" & oRs("facilityid") & """>" & sMsg &"</a></br></div>" & vbCrLf

			End If
			oRs.MoveNext
		Loop
		Response.Write("</div>" & vbCrLf)
	End If

		' CLOSE OBJECTS
		Set oRs = Nothing 
		Response.Write(vbCrLf & "<!-- finish DisplayFacility-->" & vbCrLf)
 End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYSUBCATEGORYMENU(ORGID, ICATEGORYID)
'--------------------------------------------------------------------------------------------------
Sub ListCategories( ByVal icategoryid, ByVal ipos, ByVal blnShowDetail, ByVal blnShowBreadCrumbs )
	Dim sSql, oRs
 			
	Response.Write(vbCrLf & "<!-- start ListCategories -->" & vbCrLf)
 
    ' GET CATEGORY INFORMATION
	sSql = "select * from dbo.View_1 where recreationcategoryid = " & icategoryid & " order by sequenceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

    ' LOOP THRU CATEGORIES AND DISPLAY 
    Response.Write vbCrLf & "<div class=""FACILITYLIST"">"

	' DISPLAY ITEMS UNDER CATEGORY
	If ipos = 1 Then
		DisplayFacility icategoryid 
	End If

	Do While Not oRs.EOF
		' SPACER TO SHOW CATEGORY HIEARCHY	
		iSpace = string(ipos*5,"-")

		'DEBUG CODE: Response.Write(iSpace & oRs("subcategorytitle") & "" & vbCrLf)

		' DISPLAY CATEGORY DESCRIPTION/IMAGE INFORMATION
		DisplayCategoryInformation oRs("subcategoryid"),blnShowBreadCrumbs

		' DISPLAY ITEMS UNDER SUBCATEGORY
		If blnShowDetail = 1 Then
			DisplayFacility(oRs("subcategoryid"))
		End If

		ipos = ipos + 1

		' RECURSIVE CALL TO LIST ALL SUBCATEGORIES FOR THIS CATEGORY
		'ListCategories oRs("recreationsubcategoryid"), ipos, blnShowDetail,blnShowBreadCrumbs
		oRs.MoveNext
	Loop

	ipos = 1 

	Response.Write("</div>" & vbCrLf)

	' CLOSE OBJECTS
	oRs.Close
	Set oRs = Nothing 

	Response.Write(vbCrLf & "<!-- end ListCategories -->" & vbCrLf)
End Sub


'--------------------------------------------------------------------------------------------------
' PUBLIC SUB DISPLAYSUBCATEGORYMENU(ORGID, ICATEGORYID)
'--------------------------------------------------------------------------------------------------
Sub DisplaySubCategoryMenu( orgid, iCategoryID )
	Dim sSql, oRs

	' GET SUBCATEGORY FOR THIS CATEGORY
	sSql = "select * from dbo.recreation_categories where orgid='" & orgid & "' AND recreationcategoryid='" & iCategoryID & "' ORDER BY sequenceid, subcategorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1
	blnFirst = True

	' DISPLAY LIST OF LINK TO SUBCATEGORIES OF THIS CATEGORY
	If Not oRs.EOF Then
		Response.Write("<p><div class=""subcategorymenu"">")
		' DISPLAY ROOT CATEGORY
		response.write("<font class=""subcategorymenuheader"">Browse " & oRs("categorytitle") & ":<br></font>")

		Do While Not oRs.EOF
			' WRITE SPACER
			If Not blnFirst Then
				Response.Write(" | ")
			End If
			blnFirst = False

			' DISPLAY SUBCATEGORY LINKS
			Response.Write "<a class=""subcategorymenu"" href=""facility_list.asp?categoryid=" & oRs("recreationsubcategoryid") & """ >" & oRs("subcategorytitle") & "</a> " & vbCrLf

			oRs.MoveNext
		Loop
		Response.Write vbCrLf & "</div></p>" 
	End If

	' CLEAN UP OBJECTS
	oRs.Close
	Set oRs = Nothing		

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCATEGORYINFORMATION(ICATEGORYID)
'--------------------------------------------------------------------------------------------------
Sub DisplayCategoryInformation( icategoryid, blnShowBreadCrumbs )
	Dim sSql, oRs

    ' GET SELECT CATEGORY ROW
		sSql = "select * from egov_recreation_categories where recreationcategoryid = '" &  icategoryid & "' and orgid='" & iorgid & "'"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		' DISPLAY CATEGORY INFORMATION
        Response.Write("<!--start display-->")
        Response.Write("<p><div class=""categorygroup"" onClick=""location.href='facility_list.asp?categoryid=" & icategoryid &"';"">")
        If Not oRs.EOF Then
            ' WRITE PHOTO
            sImgURL = ""
            If oRs("imgurl") = "" OR ISNULL(oRs("imgurl"))  Then
                sImgURL = "images/park_category_default.jpg"
						Else
								sImgURL = oRs("imgurl")
            End If

			' WRITE IMAGE LINK
			If CLng(icategoryid) <> CLng(GetFirstCategory) Then
				''**Response.Write("<P><a href=""facility_list.asp?categoryid=" & icategoryid & """><img class=""categoryimage"" ALIGN=""left"" src=""" & sImgURL & """></a>")
				Response.Write "<a href=""facility_list.asp?categoryid=" & icategoryid & """><img class=""categoryimage"" align=""left"" src=""" & sImgURL & """></a>"
			End If

            ' WRITE TITLE
            Response.Write "<font class=""categorytitle"" ><a class=""categorytitle"" href=""facility_list.asp?categoryid=" & icategoryid & """>" & oRs("categorytitle") & "</a><br></font>" 
            ' WRITE SUBTITLE
			If oRs("categorysubtitle") <> "" Then
				Response.Write "<font class=""categorysubtitle"" >" & oRs("categorysubtitle") & "</font><br /><br />" 
			End If
            ' WRITE DESCRIPTION
            ''**Response.Write("<font class=""categorydescription"" >" & oRs("categorydescription") & "</font><br></p>")
            Response.Write "<font class=""categorydescription"" >" & oRs("categorydescription") & "</font><br>"

        End If
        Response.Write vbCrLf & "</div></p>" 
        Response.Write("<BR clear=all><!--end display-->")


		' DISPLAY BREADCRUMBS
		If blnShowBreadCrumbs = 1  Then
			Response.Write vbCrLf & "<p><div class=""subcategorymenu"">"
			
			Response.Write("<a class=subcategorymenu href=""facility_list.asp?categoryid=" & GetFirstCategory & """ >Facility Reservations</a> | <a class=subcategorymenu href=""facility_list.asp?categoryid=" & oRs("recreationcategoryid") & """ >" & oRs("categorytitle") & "</a> " & vbCrLf)

			
			Response.Write vbCrLf & "</div></p>" & vbCrLf
		End If

        ' CLOSE OBJECTS
		oRs.CLose
        Set oRs = Nothing 

 End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYFACILITYDETAIL(IFACILITYID,iTemplateId)
'--------------------------------------------------------------------------------------------------
Sub DisplayFacilityDetail( iFacilityID, iTemplateId )

	' GET FACILITY ELEMENTS
	Dim arrImgUrl(4)
	Dim arrText(4)
	For i=1 to 4
		arrText(i) = GetText(iFacilityID,i)
		arrImgUrl(i) = GetImage(iFacilityID,i+4)
	Next 

	' DISPLAY SELECTED TEMPLATE
	response.write "<div id=""templatecontainer"">"
	
	Select Case iTemplateId

	Case 1
		response.write "<table class=template>"
		response.write "<tr>"
		response.write "<td valign=top>" & arrImgUrl(1) & "</td>"
		response.write "<td colspan=2 valign=top>" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "<tr >"
		response.write "<td class=bottomrow colspan=3 valign=top  >"
		response.write arrImgUrl(2) & arrImgUrl(3) & arrImgUrl(4) 
		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 2
		response.write "<table class=template>"
		response.write "<tr>"
		response.write "<td valign=top>" & arrImgUrl(1) & "<br>" & arrImgUrl(2) & "</td>"
		response.write "<td valign=top >" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td class=bottomrow colspan=2 valign=top align=center>"
		response.write arrImgUrl(3) & arrImgUrl(4) 
		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 3
		response.write "<table class=template>"
		response.write "<tr>"
		response.write "<td valign=top>" & arrImgUrl(1) & "</td>"
		response.write "<td valign=top colspan=2>" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<td valign=top>" & arrImgUrl(2) & "</td>"
		response.write "<td colspan=2 valign=top>" & arrText(2) & "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 4
		response.write "<table class=template>"
		response.write "<tr>"
		response.write "<td valign=top>" & arrImgUrl(1) 
		response.write "<br>" & arrText(2) & "</td>"
		response.write "<td  valign=top>" & arrText(1) 
		response.write "<br>" & arrImgUrl(2) & "</td>"
		response.write "</tr>"
		response.write "</table>"
	Case 5
		response.write "<table class=template>"
		response.write "<tr>"
		response.write "<td  valign=top>" & arrText(1) & "</td>"
		response.write "</tr>"
		response.write "</table>"
	End Select

	response.write "</div>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetImage(iFacilityId,iSequence)
'--------------------------------------------------------------------------------------------------
Function GetImage( iFacilityId, iSequence )
	Dim sSql, oRs
	
	sReturnValue = " "

	sSql = "Select elementid, content, alt_tag FROM egov_facilityelements WHERE facilityid =" & iFacilityID & " and sequence = " & iSequence & ""
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If NOT oRs.EOF Then
		If  oRs("content") <> "" Then
			sReturnValue = "<img class=""templateimg"" src=""" &  oRs("content") & """ alt=""" & oRs("alt_tag") & """ title=""" &  oRs("alt_tag") & """>"
		End If
	End If

	oRs.Close
	Set oRs = Nothing

	GetImage = sReturnValue 

End Function


'--------------------------------------------------------------------------------------------------
' Function GetText(iFacilityId,iSequence)
'--------------------------------------------------------------------------------------------------
Function GetText( iFacilityId, iSequence )
	Dim sSql, oRs
	
	sReturnValue = " "
	
	sSql = "Select elementid, content FROM egov_facilityelements WHERE facilityid =" & iFacilityID & " and sequence = " & iSequence & ""

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If NOT oRs.EOF Then
		sReturnValue = oRs("content") 	
	End If

	oRs.Close
	Set oRs = Nothing

	GetText = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' Function GetFirstCategory()
'--------------------------------------------------------------------------------------------------
Function GetFirstCategory()
	Dim sSql, oRs

	iReturnValue= "0"

	sSql = "select Top 1 * from egov_recreation_categories where orgid='" & iorgid & "' and isroot = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If not oRs.EOF Then
		iReturnValue = oRs("recreationcategoryid") 
	End If

	' CLEAN UP OBJECTS
	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetFirstCategory = iReturnValue

End Function


%>


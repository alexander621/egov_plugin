<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcategories.asp
' AUTHOR: Steve Loar
' CREATED: 01/13/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rental categories and their details. Also has links back to the old
'				facilities stuff.
'
' MODIFICATION HISTORY
' 1.0   01/13/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle 

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If
%>

<html lang="en">
<head>
	<meta charset="UTF-8">
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	
	<title><%=sTitle%></title>

	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalstyles.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<p>
	<font class="pagetitle"><%=GetOrgFeatureName( "rentals" )%></font>
	<br />
</p>


<!--BEGIN: Page Top Display-->
<% 
	If OrgHasDisplay( iorgid, "rentalscategorypagetop" ) Then
		response.write vbcrlf & "<div id=""rentalscategorypagetop"">" & GetOrgDisplay( iOrgId, "rentalscategorypagetop" ) & "</div>"
	End If 
%>
<!--END: Page Top Display-->

<%
	ShowRentalCategories
%>

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' void ShowRentalCategories
'--------------------------------------------------------------------------------------------------
Sub ShowRentalCategories( )
	Dim sSql, oRs, sGoTo

	sSql = "SELECT recreationcategoryid, ISNULL(categorytitle,'') AS categorytitle, "
	sSql = sSql & "ISNULL(categorydescription,'') AS categorydescription, ISNULL(imgurl,'') AS imgurl, isforrentals "
	sSql = sSql & "FROM egov_recreation_categories "
	sSql = sSql & "WHERE isroot = 0 AND hidefrompublic = 0 AND orgid = " & iorgid & " ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		if (iorgid = 37 and oRs("isforrentals")) or iorgid <> 37 then
			If oRs("isforrentals") then
				sGoTo = "rentalofferings"
			Else
				sGoTo = "../recreation/facilitycategory"
			End If 
			response.Write vbcrlf & "<div class=""rentalcategorygroup"" onClick=""location.href='" & sGoTo & ".asp?categoryid=" & oRs("recreationcategoryid") &"';"">"
			response.write "<table class=""rentalcategory"" cellpadding=""0"" cellspacing=""0"" border=""0""><tr><td valign=""top"">"
			If oRs("imgurl") <> "" Then 
				response.Write vbcrlf & "<a href=""" & sGoTo & ".asp?categoryid=" & oRs("recreationcategoryid") & """><img class=""categoryimage"" align=""left"" src=""" & replace(oRs("imgurl"),"http://www.egovlink.com","") & """></a>"
			End If 
			response.write "</td><td valign=""top"">"
			response.Write vbcrlf & "<div class=""categorydesc"">"
			response.write oRs("categorytitle") & "<br />"
			response.write "<p class=""categorydesc"">" & oRs("categorydescription") & "</p>"
			response.Write vbcrlf & "</div>"
			response.write "</td></tr></table>"
			response.write vbcrlf & "</div>"
		end if

		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	
End Sub


%>

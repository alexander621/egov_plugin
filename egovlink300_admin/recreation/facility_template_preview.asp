<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> TEMPLATE PREVIEW </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<style>
	.template {width:700px;}
	body {font-family: verdana,sans-serif; font-size: 12px;}
	td {font-family: verdana,sans-serif; font-size: 12px;}
	img {margin: 0px 20px 20px 20px;}
	.bottomrow {padding: 20px 0px 0px 00px;}
</style>
</HEAD>

<BODY onLoad="window.focus();">

<%
	' DISPLAY SELECTED TEMPLATE
	DisplayFacilityDetail request("ifacilityid")
%>


</BODY>
</HTML>



<%

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' PUBLIC SUB DISPLAYFACILITYDETAIL(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Public Sub DisplayFacilityDetail(iFacilityID)

	' GET FACILITY ELEMENTS
	Dim arrImgUrl(4)
	Dim arrText(4)
	For i=1 to 4
		arrText(i) = GetText(iFacilityID,i)
		arrImgUrl(i) = GetImage(iFacilityID,i+4)
	Next 

	' DISPLAY SELECTED TEMPLATE
	response.write "<div id=""templatecontainer"">"
	
	Select Case GetFacilityTemplateId(iFacilityID)

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
' PUBLIC SUB DISPLAYFACILITYDETAIL(ITEMPLATEID,IFACILITYID)
'--------------------------------------------------------------------------------------------------
Function GetImage(iFacilityId,iSequence)
	
	sReturnValue = " "

	sSQL = "Select elementid, content, alt_tag FROM egov_facilityelements WHERE facilityid =" & iFacilityID & " and sequence = " & iSequence & ""
	Set oImageInfo = Server.CreateObject("ADODB.Recordset")
	oImageInfo.Open sSQL, Application("DSN") , 3, 1

	If NOT oImageInfo.EOF Then
		If  oImageInfo("content") <> "" Then
			sReturnValue = "<img src=""" &  oImageInfo("content") & """ alt=""" & oImageInfo("alt_tag") & """ title=""" &  oImageInfo("alt_tag") & """>"
		End If
	End If

	Set oImageInfo = Nothing

	GetImage = sReturnValue 

End Function


'--------------------------------------------------------------------------------------------------
' PUBLIC SUB DISPLAYFACILITYDETAIL(ITEMPLATEID,IFACILITYID)
'--------------------------------------------------------------------------------------------------
Function GetText(iFacilityId,iSequence)
	
	sReturnValue = " "
	
	sSQL = "Select elementid, content FROM egov_facilityelements WHERE facilityid =" & iFacilityID & " and sequence = " & iSequence & ""
	Set oText = Server.CreateObject("ADODB.Recordset")
	oText.Open sSQL, Application("DSN") , 3, 1

	If NOT oText.EOF Then
		sReturnValue = oText("content") 	
	End If

	Set oText = Nothing

	GetText = sReturnValue

End Function

'--------------------------------------------------------------------------------------------------
' FUNCTION GETFACILITYTEMPLATEID(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Function GetFacilityTemplateId(iFacilityID)

	iReturnValue = 0

	sSQL = "select facilitytemplateid from egov_facility WHERE facilityid='" & ifacilityid & "'"
	Set oID = Server.CreateObject("ADODB.Recordset")
	oID.Open sSQL, Application("DSN") , 3, 1

	If NOT oID.EOF Then
		iReturnValue = oID("facilitytemplateid") 	
	End If

	Set oID = Nothing

	GetFacilityTemplateId = iReturnValue
End Function
%>
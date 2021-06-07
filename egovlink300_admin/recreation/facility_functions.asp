<%
'--------------------------------------------------------------------------------------------------
' Function GetFacilityName(iFacilityID)
'--------------------------------------------------------------------------------------------------
Function GetFacilityName(iFacilityID)
	
	sSQL = "Select facilityname FROM egov_facility WHERE facilityid =" & iFacilityID & ""
	Set oFacName = Server.CreateObject("ADODB.Recordset")
	oFacName.Open sSQL, Application("DSN") , 3, 1
	If oFacName.eof Then
		GetFacilityName = ""
	Else
		GetFacilityName = oFacName("facilityname")
	End If 
	oFacName.close
	Set oFacName = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetTextArea(iFacilityId,iSequence)
'--------------------------------------------------------------------------------------------------
Function GetTextArea(iFacilityId,iSequence, ByRef sFound)
	sSQL = "Select elementid, content FROM egov_facilityelements WHERE facilityid =" & iFacilityID & " and sequence = " & iSequence & ""
	Set oTextArea = Server.CreateObject("ADODB.Recordset")
	oTextArea.Open sSQL, Application("DSN") , 3, 1

	If oTextArea.eof Then
		GetTextArea = "<textarea name=" & Chr(34) & "content" & Chr(34) & "></textarea> "
		GetTextArea = GetTextArea & vbcrlf & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "elementid" & Chr(34) & " value=" & Chr(34) & "0" & Chr(34) & " />"
		sFound = "no"
	Else
		If Trim(oTextArea("content")) <> "" then
			GetTextArea = "<textarea name=" & Chr(34) & "content" & Chr(34) & ">" & oTextArea("content") & "</textarea> " 
			GetTextArea = GetTextArea & vbcrlf & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "elementid" & Chr(34) & " value=" & Chr(34) & oTextArea("elementid") & Chr(34) & " />"
			sFound = "yes"
		Else
			GetTextArea = "<textarea name=" & Chr(34) & "content" & Chr(34) & "></textarea> "
			GetTextArea = GetTextArea & vbcrlf & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "elementid" & Chr(34) & " value=" & Chr(34) & oTextArea("elementid") & Chr(34) & " />"
			sFound="no"
		End If 
	End If 

	oTextArea.close
	Set oTextArea = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetTextArea(iFacilityId,iSequence)
'--------------------------------------------------------------------------------------------------
Function GetImageInfo(iFacilityId,iSequence, ByRef sFound)
	sSQL = "Select elementid, content, alt_tag FROM egov_facilityelements WHERE facilityid =" & iFacilityID & " and sequence = " & iSequence & ""

'	GetImageInfo = sSQL

	Set oImageInfo = Server.CreateObject("ADODB.Recordset")
	oImageInfo.Open sSQL, Application("DSN") , 3, 1

	If oImageInfo.eof Then
		GetImageInfo = "No Image</td>"
		GetImageInfo = GetImageInfo & vbcrlf & "<td class=" & Chr(34) & "imageinput" & Chr(34) & ">"
		GetImageInfo = GetImageInfo & vbcrlf & "<a class=" & Chr(34) & "selectimage" & Chr(34) & " href=" & Chr(34) & "javascript:doPicker('element" & iSequence & ".content');" & Chr(34) & ">Select an Image</a><br />"
		GetImageInfo = GetImageInfo & vbcrlf & "URL: <input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "content" & Chr(34) & " size=" & Chr(34) & "50" & Chr(34) & " maxlength=" & Chr(34) & "250" & Chr(34) & " /><br />"
		GetImageInfo = GetImageInfo & vbcrlf & "Alt Tag: <input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "alt_tag" & Chr(34) & " size=" & Chr(34) & "50" & Chr(34) & " maxlength=" & Chr(34) & "250" & Chr(34) & " />"
		GetImageInfo = GetImageInfo & vbcrlf & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "elementid" & Chr(34) & " value=" & Chr(34) & "0" & Chr(34) & " />"
		sFound = "no"
	Else
		If Trim(oImageInfo("content")) <> "" then
			GetImageInfo = "<img src=" & Chr(34) & oImageInfo("content") & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " alt=" & Chr(34) & oImageInfo("alt_tag") & Chr(34) & " /></td>"
			GetImageInfo = GetImageInfo & vbcrlf & "<td class=" & Chr(34) & "imageinput" & Chr(34) & ">"
			GetImageInfo = GetImageInfo & vbcrlf & "<a class=" & Chr(34) & "selectimage" & Chr(34) & " href=" & Chr(34) & "javascript:doPicker('element" & iSequence & ".content');" & Chr(34) & ">Select an Image</a><br />"
			GetImageInfo = GetImageInfo & vbcrlf & "URL: <input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "content" & Chr(34) & "value= " & Chr(34) &  oImageInfo("content") & Chr(34) & " size=" & Chr(34) & "50" & Chr(34) & " maxlength=" & Chr(34) & "250" & Chr(34) & " /><br />"
			GetImageInfo = GetImageInfo & vbcrlf & "Alt Tag: <input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "alt_tag" & Chr(34) & " value= " & Chr(34) &  oImageInfo("alt_tag") & Chr(34) & " size=" & Chr(34) & "50" & Chr(34) & " maxlength=" & Chr(34) & "250" & Chr(34) & " />"
			GetImageInfo = GetImageInfo & vbcrlf & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "elementid" & Chr(34) & " value=" & Chr(34) & oImageInfo("elementid") & Chr(34) & " />"
			sFound = "yes"
		Else
			GetImageInfo = "No Image</td>"
			GetImageInfo = GetImageInfo & vbcrlf & "<td class=" & Chr(34) & "imageinput" & Chr(34) & ">"
			GetImageInfo = GetImageInfo & vbcrlf & "<a class=" & Chr(34) & "selectimage" & Chr(34) & " href=" & Chr(34) & "javascript:doPicker('element" & iSequence & ".content');" & Chr(34) & ">Select an Image</a><br />"
			GetImageInfo = GetImageInfo & vbcrlf & "URL: <input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "content" & Chr(34) & " size=" & Chr(34) & "50" & Chr(34) & " maxlength=" & Chr(34) & "250" & Chr(34) & " /><br />"
			GetImageInfo = GetImageInfo & vbcrlf & "Alt Tag: <input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & "alt_tag" & Chr(34) & " size=" & Chr(34) & "50" & Chr(34) & " maxlength=" & Chr(34) & "250" & Chr(34) & " />"
			GetImageInfo = GetImageInfo & vbcrlf & "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "elementid" & Chr(34) & " value=" & Chr(34) & oImageInfo("elementid") & Chr(34) & " />"
			sFound = "no"
		End If 
	End If 
	
	oImageInfo.close
	Set oImageInfo = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetRateSelect(iFacilityId, iRateId)
'--------------------------------------------------------------------------------------------------
Function GetRateSelect(iFacilityId, iRateId)

	sSQL = "Select ratedescription, rateid FROM egov_rate WHERE facilityid =" & iFacilityID & ""

	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.Open sSQL, Application("DSN") , 3, 1
	
	GetRateSelect = "<select name=" & Chr(34) & "rateid" & Chr(34) & ">"
	Do While not oRates.eof 
		GetRateSelect = GetRateSelect & vbcrlf & "<option value=" & Chr(34) & oRates("rateid") & Chr(34) 
		If oRates("rateid") = iRateId Then
			GetRateSelect = GetRateSelect & " selected=" & Chr(34) & "selected" & Chr(34) 
		End If 
		GetRateSelect = GetRateSelect & ">"
		GetRateSelect = GetRateSelect & oRates("ratedescription") & "</option>"
		oRates.MoveNext
	Loop 
	GetRateSelect = GetRateSelect & vbcrlf & "</select>"

	oRates.close
	Set oRates = nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetMaxDisplayOrder(iFacilityID)
'--------------------------------------------------------------------------------------------------
Function GetMaxDisplayOrder(iFacilityID)
	Dim sSql

	sSQL = "Select MAX(displayorder) as MaxOrder FROM egov_recreation_terms WHERE facilityid =" & iFacilityID & ""
	Set oMax = Server.CreateObject("ADODB.Recordset")
	oMax.Open sSQL, Application("DSN") , 3, 1
	If IsNull(oMax("MaxOrder")) Then
		GetMaxDisplayOrder = 0
	Else
		GetMaxDisplayOrder = oMax("MaxOrder")
	End If 
	oMax.close
	Set oMax = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function CheckWaiverDisplay(iFacilityId, iWaiverId)
'--------------------------------------------------------------------------------------------------
Function CheckWaiverDisplay(iFacilityId, iWaiverId)
	
	CheckWaiverDisplay = ""
	sSQL = "Select waiverid FROM egov_facilitywaivers WHERE facilityid =" & iFacilityID & ""

	Set oWavers = Server.CreateObject("ADODB.Recordset")
	oWavers.Open sSQL, Application("DSN") , 3, 1
	
	Do While not oWavers.eof 
		If oWavers("waiverid") = iWaiverId Then
			CheckWaiverDisplay = "checked=checked"
			Exit Do 
		End if
		oWavers.MoveNext
	Loop 

	oWavers.close
	Set oWavers = nothing

End Function 


'--------------------------------------------------------------------------------------------------
'  PUBLIC SUB SUBDRAWSELECTTEMPLATE(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Public Sub subDrawSelectTemplate(ifacilityid)
	
	If ifacilityid = "" Then
		ifacilityid = 0
	End If

	' GET SELECT CATEGORY ROW
	sSQL = "select *,(select facilitytemplateid from egov_facility WHERE facilityid='" & ifacilityid & "') AS iseltemplateid from egov_facility_templates Order by templatename"
	Set oTemplates = Server.CreateObject("ADODB.Recordset")
	oTemplates.Open sSQL, Application("DSN") , 3, 1

    ' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
    Response.Write("<select name=""seltemplate"" >")
    Do While NOT oTemplates.EOF
		sSelected = ""

		If oTemplates("iseltemplateid") = clng(oTemplates("templateid")) Then
			sSelected = "SELECTED"
		End If
		
		Response.Write("<option " & sSelected & " value=""" & oTemplates("templateid") & """>" & oTemplates("templatename") & "</option>" & vbCrLf)
		oTemplates.MoveNext
	Loop
    Response.Write("</select>" & vbCrLf)

	' DESTROY OBJECTS
	Set oTemplates = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETFACILITYTEMPLATENAME(IFACILITYID)
'--------------------------------------------------------------------------------------------------
Function GetFacilityTemplateName(iFacilityID)

	sReturnValue = "UNKNOWN"

	sSQL = "select templatename from egov_facility INNER JOIN egov_facility_templates ON egov_facility.facilitytemplateid=egov_facility_templates.templateid WHERE facilityid='" & ifacilityid & "'"
	Set oID = Server.CreateObject("ADODB.Recordset")
	oID.Open sSQL, Application("DSN") , 3, 1

	If NOT oID.EOF Then
		sReturnValue = oID("templatename") 	
	End If

	Set oID = Nothing

	GetFacilityTemplateName = sReturnValue
End Function


'--------------------------------------------------------------------------------------------------
' Function GetFacilityName(iFacilityID)
'--------------------------------------------------------------------------------------------------
Sub SetFacilityInformation(iFacilityID)
	
	sSQL = "Select * FROM egov_facility WHERE facilityid =" & iFacilityID & ""
	Set oFacName = Server.CreateObject("ADODB.Recordset")
	oFacName.Open sSQL, Application("DSN") , 3, 1
	If oFacName.eof Then
		sFacilityName = ""
	Else
		sFacilityName= oFacName("facilityname")
		chkisviewable = oFacName("isviewable")
		chkisreservable = oFacName("isreservable")
	End If 
	oFacName.close
	Set oFacName = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYRATES(ISELECTEDID)
'--------------------------------------------------------------------------------------------------
Sub SubDisplayRates( iSelectedID )
	Dim sSql, oRates, sSelected
		
	sSql = "SELECT rateid,ratedescription FROM egov_facility_rates WHERE orgid = " & session("orgid") & " ORDER BY ratedescription"

	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.Open sSql, Application("DSN"), 3, 1

	If Not oRates.EOF Then
		
		response.write vbcrlf & "<select name=""ratechoice"">"
		response.write vbcrlf & "<option value=""0"">Select rate from list ...</option>"
		
		Do While Not oRates.EOF
			If iSelectedID = clng(oRates("rateid")) Then
				sSelected = " selected=""selected"" "
			Else
				sSelected = "" 
			End If
			response.write vbcrlf & "<option " & sSelected  & " value=""" &  oRates("rateid") & """>" & oRates("ratedescription") & "</option>"
			oRates.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If
	oRates.Close 
	Set oRates = Nothing
End Sub


'--------------------------------------------------------------------------------------------------
' Function GetFacilityCategory( iFacilityId ) 
'--------------------------------------------------------------------------------------------------
Function GetFacilityCategory( iFacilityId ) 
	Dim sSql, oCategory

	sSql = "Select categoryid From egov_recreation_category_to_item where itemid = " & iFacilityId

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSQL, Application("DSN"), 3, 1

	If Not oCategory.EOF Then
		GetFacilityCategory = oCategory("categoryid")
	Else
		GetFacilityCategory = 0
	End If 

	oCategory.Close
	Set oCategory = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string  GetResidencyDescription( iOrgId, sResidentType ) 
'--------------------------------------------------------------------------------------------------
Function GetResidencyDescription( ByVal iOrgId, ByVal sResidentType )
	Dim sSql, oRs
	
	sSql = "SELECT description FROM egov_poolpassresidenttypes WHERE orgid = " & iOrgId & " AND resident_type = '" & sResidentType & "'"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetResidencyDescription = oRs("description")
	Else
		GetResidencyDescription =""
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowCategoryPicks( iCurrentCategory )
'--------------------------------------------------------------------------------------------------
Sub ShowCategoryPicks( iCurrentCategory )
	Dim sSql, oCategories
		
	sSQL = "Select recreationcategoryid, categorytitle From egov_recreation_categories where orgid = " & session("orgid") & " and isroot = 0 order by categorytitle"

	Set oCategories = Server.CreateObject("ADODB.Recordset")
	oCategories.Open sSQL, Application("DSN"), 3, 1

	If Not oCategories.EOF Then
		
		response.write "<select name=""categoryid"">"
		response.write "<option value=""0"">Select category from list ...</option>"
		
		Do While Not oCategories.EOF
			If clng(iCurrentCategory) = clng(oCategories("recreationcategoryid")) Then
				sSelected = " selected=""selected"" "
			Else
				sSelected = "" 
			End If
			response.write "<option " & sSelected  & " value=""" &  oCategories("recreationcategoryid") & """>" & oCategories("categorytitle") & "</option>"
			oCategories.MoveNext
		Loop
		response.write "</select>"
	End If
	oCategories.Close 
	Set oCategories = Nothing

End Sub 
%>

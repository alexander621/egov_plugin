<%
Call subOrderCategories(request("iCatId"),UCASE(request("direction")),session("orgid"))

'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBORDERCATEGORIES(ICATID)
'--------------------------------------------------------------------------------------------------
Sub subOrderCategories(iCatId,sDirection,iorgID)
	
	' REORDER QUESTIONS
	iSequence = 0
	sSQL = "Select * FROM egov_form_categories WHERE orgid='" & iorgID & "' ORDER BY form_category_sequence"
	Set oOrder = Server.CreateObject("ADODB.Recordset")
	oOrder.Open sSQL, Application("DSN") , 3, 2
	iNumberOfCats = oOrder.Recordcount
	
	' REPLACE ANY NULL SEQUENCE WITH CURRENT SEQUENCE
	If NOT oOrder.EOF Then

		Do While NOT oOrder.EOF 
			iSequence = iSequence + 1
			If clng(iCatId) = clng(oOrder("form_category_id")) Then
				iCurrentSequence = iSequence
			End If
			oOrder("form_category_sequence") = iSequence
			oOrder.Update
		oOrder.MoveNext
		Loop

	End If
	
	Set oOrder = Nothing
	
	' PROCESS QUESTION MOVE
	If sDirection = "UP" Then
		iNewSequence = iCurrentSequence - 1
		If iNewSequence < 1 Then
			iNewSequence = 1
		End If
	End If

	If sDirection = "DOWN" Then
		iNewSequence = iCurrentSequence + 1
		If iNewSequence > iNumberOfCats Then
			iNewSequence = iNumberOfCats
		End If
	End If

	If sDirection = "TOP" Then
		iNewSequence = 0
	End If

	If sDirection = "BOTTOM" Then
		iNewSequence = iNumberOfCats + 1
	End If


	' APPLY QUESTION MOVE
	If iNewSequence <> iCurrentSequence Then
		sSQL =  "UPDATE egov_form_categories SET form_category_sequence='" & iCurrentSequence & "' WHERE orgid='" & iOrgID & "' AND form_category_sequence='" & iNewSequence & "'"
		sSQL2 = "UPDATE egov_form_categories SET form_category_sequence='" & iNewSequence & "' WHERE orgid='" & iOrgID & "' AND form_category_id='" & iCatId & "'"
		RESPONSE.WRITE sSQL & "<BR>"
		RESPONSE.WRITE sSQL2 & "<BR>"
		
		Set oOrder = Server.CreateObject("ADODB.Recordset")
		oOrder.Open sSQL, Application("DSN") , 3, 1
		oOrder.Open sSQL2, Application("DSN") , 3, 1
		Set oOrder = Nothing
	End If

	' REDIRECT TO MANANGE FORM PAGE
	 response.redirect("actioncategories.asp")

End Sub
%>
<%
Call subOrderQuestions(request("ifieldid"),UCASE(request("direction")),request("iformid"))
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
' SUB SUBGETFORMINFORMATION(IFORMID)
'--------------------------------------------------------------------------------------------------
Sub subOrderQuestions(iFieldID,sDirection,iFormID)
	Dim oOrder, iSequence, sSQL, iNumberofQuestions, iCurrentSequence
	Dim oCmd
	
'REORDER QUESTIONS
	iSequence = CLng(0)
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_action_form_questions "
 sSQL = sSQL & " WHERE formid = " & iFormID
 sSQL = sSQL & " ORDER BY SEQUENCE "
	Set oOrder = Server.CreateObject("ADODB.Recordset")
' oOrder.Open sSQL, Application("DSN"), 3, 2
	oOrder.Open sSQL, Application("DSN"), 0, 1
'	iNumberofQuestions = oOrder.Recordcount
	
'REPLACE ANY NULL SEQUENCE WITH CURRENT SEQUENCE
	If NOT oOrder.EOF Then
  		Do While NOT oOrder.EOF 
		    	iSequence = iSequence + CLng(1)

     		If clng(iFieldID) = clng(oOrder("questionid")) Then
      				iCurrentSequence = iSequence
    			End If

     			UpdateSequence oOrder("questionid"), iSequence
     			'oOrder("sequence") = iSequence
     			'oOrder.Update
      		oOrder.MoveNext
    Loop
    iNumberofQuestions = iSequence
 else
    iNumberofQuestions = 0
	End If

	oOrder.Close 
	Set oOrder = Nothing

'PROCESS QUESTION MOVE
	If sDirection = "UP" Then
	  	iNewSequence = iCurrentSequence - 1
	  	If iNewSequence < 1 Then
    			iNewSequence = 1
  		End If
	End If

	If sDirection = "DOWN" Then
  		iNewSequence = iCurrentSequence + 1
  		If iNewSequence > iNumberofQuestions Then
		    	iNewSequence = iNumberofQuestions
  		End If
	End If

	If sDirection = "TOP" Then
		  iNewSequence = 0
	End If

	If sDirection = "BOTTOM" Then
  		iNewSequence = iNumberofQuestions + 1
	End If

'APPLY QUESTION MOVE
	If iNewSequence <> iCurrentSequence Then
  		Set oCmd = Server.CreateObject("ADODB.Command")
		  oCmd.ActiveConnection = Application("DSN")

  		sSQL = "UPDATE egov_action_form_questions SET sequence = " & iCurrentSequence
    sSQL = sSQL & " WHERE formid = " & iFormID
    sSQL = sSQL & " AND sequence = " & iNewSequence

		  sSQL2 = "UPDATE egov_action_form_questions SET sequence = " & iNewSequence
    sSQL2 = sSQL2 & " WHERE formid = "   & iFormID
    sSQL2 = sSQL2 & " AND questionid = " & iFieldID 

	  	'Set oOrder = Server.CreateObject("ADODB.Recordset")
 		 'oOrder.Open sSQL, Application("DSN") , 3, 1
  		'oOrder.Open sSQL2, Application("DSN") , 3, 1
  		oCmd.CommandText = sSQL
  		oCmd.Execute
  		oCmd.CommandText = sSQL2
  		oCmd.Execute
		
		Set oCmd = Nothing
	End If

'REDIRECT TO MANANGE FORM PAGE
	response.redirect("manage_form.asp?iformid=" & iFormID)
End Sub

'-------------------------------------------------------------------------------------------------
' Sub UpdateSequence( iQuestionid, iSequence )
'-------------------------------------------------------------------------------------------------
Sub UpdateSequence( iQuestionid, iSequence )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = "UPDATE egov_action_form_questions SET sequence = " & iSequence & " WHERE questionid = " & iQuestionid
	oCmd.Execute
	Set oCmd = Nothing
End Sub 


%>
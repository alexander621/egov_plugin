<%
Call subOrderLetters(request("iLetterID"),UCASE(request("direction")),session("orgid"))

'------------------------------------------------------------------------------
sub subOrderLetters(iLetterID,sDirection,iorgID)

  if sDirection <> "" then
     lcl_direction = UCASE(sDirection)
  else
     lcl_direction = ""
  end if

	'BEGIN: Reorder questions ----------------------------------------------------
 	iSequence = 0

 	sSQL = "SELECT * "
  sSQL = sSQL & " FROM Formletters "
  sSQL = sSQL & " WHERE orgid = '" & iorgID & "' "
  sSQL = sSQL & " ORDER BY SEQUENCE"

 	set oOrder = Server.CreateObject("ADODB.Recordset")
 	oOrder.Open sSQL, Application("DSN"), 3, 2

 	iNumberOfLtrs = oOrder.Recordcount
	
	'Replace any NULL sequence with current sequence
 	if not oOrder.eof then
   		do while not oOrder.eof
     			iSequence = iSequence + 1

     			if clng(iLetterID) = clng(oOrder("FLid")) then
       				iCurrentSequence = iSequence
     			end if

     			oOrder("sequence") = iSequence
     			oOrder.Update
      		oOrder.MoveNext
   		loop
	 end if

 	set oOrder = nothing

	'END: Reorder questions ------------------------------------------------------

	'BEGIN: Process question move ------------------------------------------------
 	if lcl_direction = "UP" then
   		iNewSequence = iCurrentSequence - 1

   		if iNewSequence < 1 then
     			iNewSequence = 1
   		end if

     lcl_success = "MOVE_UP"
  end if

 	if lcl_direction = "DOWN" then
   		iNewSequence = iCurrentSequence + 1

   		if iNewSequence > iNumberOfLtrs then
     			iNewSequence = iNumberOfLtrs
     end if

     lcl_success = "MOVE_DOWN"

  end if

	 if lcl_direction = "TOP" then
   		iNewSequence = 0
     lcl_success  = "MOVE_TOP"
  end if

	 if lcl_direction = "BOTTOM" then
   		iNewSequence = iNumberOfLtrs + 1
     lcl_success  = "MOVE_BOTTOM"

  end if

  if lcl_direction = "DELETE" then
     lcl_success = "SD"
  end if
	'END: Process question move --------------------------------------------------

	'BEGIN: Apply question move --------------------------------------------------
 	if iNewSequence <> iCurrentSequence then
   		sSQL =  "UPDATE Formletters SET sequence='" & iCurrentSequence & "' WHERE orgid='" & iOrgID & "' AND sequence='" & iNewSequence & "'"
   		sSQL2 = "UPDATE Formletters SET sequence='" & iNewSequence     & "' WHERE orgid='" & iOrgID & "' AND FLid='"     & iLetterID    & "'"
		
   		set oOrder = Server.CreateObject("ADODB.Recordset")
   		oOrder.Open sSQL, Application("DSN"), 3, 1
   		oOrder.Open sSQL2, Application("DSN"), 3, 1

   		set oOrder = nothing
  end if
	'END: Apply question move ----------------------------------------------------

  response.redirect "list_letter.asp?success=" & lcl_success

end sub
%>
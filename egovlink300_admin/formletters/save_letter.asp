<%
  if request("containsHTML") = "Y" then
     lcl_containsHTML = 1
  else
     lcl_containsHTML = 0
  end if

subUpdateFL request("iLetterID"), session("orgid"), lcl_containsHTML

'--------------------------------------------------------------------------------------------------
sub subUpdateFL( iLetterID, iOrgID, iContainsHTML )
	Dim oCmd, sSQL, oCheck
 lcl_success = ""

 if iLetterID <> "" then
    if CLng(iLetterID) = CLng(0) then
       iLetterID = ""
    end if
 end if

 if iLetterID <> "" then
  		sSQL = "SELECT FLid FROM FormLetters WHERE FLid = " & CLng(iLetterID)

		  set oCheck = Server.CreateObject("ADODB.Recordset")
  		oCheck.Open sSql, Application("DSN"), 3, 1

  		if not oCheck.eof then
    			sSQL = "UPDATE FormLetters SET "
 						sSQL = sSQL & " FLtitle = '"     & DBsafe(request("FLtitle")) & "', "
 						sSQL = sSQL & " FLbody = '"      & DBsafe(request("FLbody"))  & "', "
       sSQL = sSQL & " containsHTML = " & iContainsHTML
 						sSQL = sSQL & " WHERE orgid = " & iOrgID
 						sSQL = sSQL & " AND FLid = "    & iLetterID

      	set oUpdateFormLetter = Server.CreateObject("ADODB.Recordset")
      	oUpdateFormLetter.Open sSQL, Application("DSN") , 3, 1
      	set oUpdateFormLetter = nothing

       'lcl_success = "SU"

 						'set oCmd = Server.CreateObject("ADODB.Command")
 						'With oCmd
 						'	.ActiveConnection = Application("DSN")
				 		'	.CommandText = sSQL
 						'	.Execute
 						'End With

       'set oCmd = nothing
    end if

  		oCheck.Close
  		set oCheck = nothing

   	response.redirect "manage_letter.asp?iletterid=" & iLetterID & "&success=SU"

 else
  		sSQL = "INSERT INTO FormLetters ("
    sSQL = sSQL & "FLtitle, "
    sSQL = sSQL & "FLbody, "
    sSQL = sSQL & "orgid, "
    sSQL = sSQL & "sequence, "
    sSQL = sSQL & "containsHTML "
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & "'" & DBsafe(request("FLtitle")) & "', "
    sSQL = sSQL & "'" & DBsafe(request("FLbody"))  & "', "
    sSQL = sSQL &       iOrgID                     & ", "
    sSQL = sSQL &       fnGetSequenceNumber()      & ", "
    sSQL = sSQL &       iContainsHTML              & ")"

   	set oInsertFormLetter = Server.CreateObject("ADODB.Recordset")
   	oInsertFormLetter.Open sSQL, Application("DSN"), 3, 1
   	set oInsertFormLetter = nothing

    'lcl_success = "SA"

  		'set oCmd = Server.CreateObject("ADODB.Command")
  		'With oCmd
		  '	.ActiveConnection = Application("DSN")
  		'	.CommandText = sSQL
		  '	.Execute
  		'End With

  		set oCmd = nothing

   	response.redirect "list_letter.asp?success=SA"

	end if

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )
  if not VarType( strDB ) = vbString then DBsafe = strDB : exit function
  DBsafe = Replace( strDB, "'", "''" )
end function

'------------------------------------------------------------------------------
function fnGetSequenceNumber()
	Dim sSQLNewSequence, oSequence, iReturnValue

	iReturnValue = CLng(1)

	sSQLNewSequence = "SELECT isnull(MAX(Sequence),0) as newsequence "
 sSQLNewSequence = sSQLNewSequence & " FROM Formletters "
 sSQLNewSequence = sSQLNewSequence & " WHERE OrgID = " & session("orgid") 

	set oSequence = Server.CreateObject("ADODB.Recordset")
	oSequence.Open sSQLNewSequence, Application("DSN"), 3, 1

	if not oSequence.eof then
  		iReturnValue = CLng(oSequence("newsequence")) + 1
	end if

	oSequence.Close
	set oSequence = nothing

	fnGetSequenceNumber = iReturnValue

end function
%>
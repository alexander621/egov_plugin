<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: DELETE_LETTER.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 3/1/2007
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  DELETE A FORM LETTER.
'
' MODIFICATION HISTORY
' 1.0	03/14/07	John Stullenberger - Initial Version
' 1.1 10/26/09 David Boyer - Added redirect to reorder sequence
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'Delete the form letter
 lcl_deleted = ""

 if request("iletterid") <> "" then
    subDeleteLetter request("iletterid")
    
    lcl_deleted = "SD"
 end if

 'response.redirect "list_letter.asp?success=" & lcl_deleted
 response.redirect "order_letter.asp?iLetterID = " & request("iletterid") & "&direction=DELETE"

'------------------------------------------------------------------------------
sub subDeleteLetter(iLetterID)
	
	sSQL = "DELETE FROM FormLetters "
 sSQL = sSQL & " WHERE FLid = '" & iLetterID & "' "
 sSQL = sSQL & " AND orgid = " & session("orgid")

	set oDeleteLetter = Server.CreateObject("ADODB.Recordset")
	oDeleteLetter.Open sSQL, Application("DSN") , 3, 1
	set oDeleteLetter = nothing

end sub
%>
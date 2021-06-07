<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: catch404.asp
' AUTHOR: Steve Loar
' CREATED: 02/04/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This catches the 404 errors and routes accoringly.
'
' MODIFICATION HISTORY
' 1.0   02/04/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sReferer

sReferer = request.servervariables("HTTP_REFERER")

If InStr(sReferer, "&") > 0 Then
	sReferer = Replace(sReferer, "&", "and")
	response.redirect sReferer
Else
	response.redirect "http://www.egovlink.com/custom404b.htm"
End If 


%>
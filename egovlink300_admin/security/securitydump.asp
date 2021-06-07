<%
Server.ScriptTimeout = 1200
iOrgID = 106
adOpenStatic = 3
adLockReadOnly = 1
	sSql = "SELECT userid, firstname, lastname FROM users "
	sSql = sSql & "WHERE orgid = " & iOrgID & " AND isdeleted = 0"

		sSql = sSql & " AND (isrootadmin IS NULL OR isrootadmin = 0) "

	sSql = sSql & " ORDER BY lastname, firstname"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	x = 0
	If Not oRs.EOF Then 

		Do While Not oRs.EOF
				x = x + 1
				if x > 72 then
				response.write x & "/" & oRs.RecordCount & " - " & oRs("lastname") & ", " & oRs("firstname") & "<br />" & vbcrlf

				set fs=Server.CreateObject("Scripting.FileSystemObject") 
				set f=fs.CreateTextFile(Server.MapPath("securitydump") & "\" & x & ".txt",true)
				f.write getHTML("http://www.egovlink.com/romeoville/admin/security/edit_user_security_dump.asp?iuserid=" & oRs("userid"))
				response.flush
				f.close
				set f=nothing
				set fs=nothing
				end if

			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

function getHTML (strUrl)
		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

		' Set timeouts of resolve(0), connection(60000), send(30000), receive(30000) in milliseconds. 0 = infinite
		objWinHttp.SetTimeouts 0, 120000, 60000, 120000

		'response.write strURL & "<br />" & vbcrlf
		objWinHttp.Open "GET", strURL, False

		objWinHttp.setRequestHeader "Content-Type", "text/namevalue"
		objWinHttp.Send

		If objWinHttp.Status = 200 Then 
			' Get the text of the response.
			results = objWinHttp.ResponseText
		End If 

		' Trash our object now that we are finished with it.
		Set objWinHttp = Nothing

    getHTML = results
end function
%>

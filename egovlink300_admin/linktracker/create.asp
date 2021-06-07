<%
if request.servervariables("REQUEST_METHOD") = "POST" then
	'response.write request.form("body") & "<br />"
	strSendToName = replace(request.form("sendtoname"),"'","''")
	blnLinks = true
	strBody = replace(request.form("body"),"'","''")
	x = 0
	Do While blnLinks
		if instr(strBody,"[") > 0 then
			'strBody = replace(strBody,"[","")
			intTextStart = instr(strBody,"[") + 1
			intTextEnd = instr(strBody,"]") - intTextStart
			intLinkStart = instr(strBody,"(") + 1
			intBlockEnd = instr(strBody,")")
			intLinkEnd = intBlockEnd - intLinkStart
			'response.write intTextStart & " - " & intTextEnd & " - " & intLinkStart & " - " & intLinkEnd & " - " & intBlockEnd & "<br />"

			strLinkText = mid(strBody,intTextStart,intTextEnd)
			strLinkURL =  mid(strBody,intLinkStart,intLinkEnd)

			'Generate Random Code
			strCode = getCode

			'Insert into database
			sSQL = "INSERT INTO LinkTracker (linkcode,sendtoname,linktourl,linktotext) VALUES('" & strCode & "','" & strSendToName & "','" & strLinkURL & "','" & strLinkText & "')"
			Set oCmd = Server.CreateObject("ADODB.Connection")
			oCmd.Open Application("DSN")
			oCmd.Execute(sSQL)
			oCmd.Close
			Set oCmd = Nothing


			strLinkURL = "http://www.egovlink.com/eclink/link.asp?code=" & strCode

			strTranslatedLink = "<a href=""" & strLinkURL & """>" & strLinkText & "</a>"

			strBody = Left(strBody,intTextStart-2) & strTranslatedLink & mid(strBody,intBlockEnd + 1)
			'response.write strBody & "<hr />"
		else
			blnLinks = false
		end if
	loop
	response.write "Copy and paste the text below into your email:<br />"
	response.write "<pre>" & strBody & "</pre><br />"
	response.flush
	'response.write x
	response.end
end if

Function getCode()

	NotUnique = true
	y = 0
	Do While NotUnique

		Randomize
		num = int(Rnd() * 1000000)
		'response.write num & " = "

		alphabet = "123456789abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ"
		base_count = len(alphabet)
        	encoded = ""
        	div = 0
        	intmod = 0
	
		Do While num >= base_count
			div = num / base_count
			intmod = ((num - (base_count * int(div)))) + 1
	
			'response.write intmod & "<br />"
			encoded = mid(alphabet,intmod,1) & encoded
			num = div
		loop
	
		if num > 0 then
			encoded = mid(alphabet,num,1) & encoded
		end if
		'response.write encoded & "<br />"
		
		sSQL = "SELECT linktrackerid FROM linktracker WHERE linkcode = '" & encoded & "'"
		set oCode = Server.CreateObject("ADODB.RecordSet")
		oCode.Open sSQL, Application("DSN"), 3, 1
		if oCode.EOF then NotUnique = false
		oCode.Close
		set oCode = Nothing
	loop
	
        getCode = encoded
End Function

%>
<form method="POST">
	Who are you sending email to?<br />
	<input type="text" name="sendtoname" value="" maxlength="50" placeholder="Betty at Sycamore" />
	<br />
	What's the contents of the email?
	<br />
	<textarea name="body" cols="110" rows="20"></textarea>
	<br />
	<input type="submit" />
</form>

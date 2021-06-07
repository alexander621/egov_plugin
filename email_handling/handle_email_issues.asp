<!--#include file="includes/aspJSON1.17.asp" -->
<%
dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.OpenTextFile(Server.MapPath("emaillog.txt"),8,true)
'f.WriteLine("This text will be added to the end of file")

	f.WriteLine("")
	f.WriteLine("")
	f.WriteLine("################################################################################################")
	f.WriteLine("NEW LOG: " & Now())
%>
<!--
	f.WriteLine("")
	f.WriteLine("Session:")
	For each session_name in Session.Contents
		f.WriteLine(sSessionLog & session_name & ":  " & session(session_name))
	Next

	f.WriteLine("")
	f.WriteLine("Cookies:")
	For Each Item in Request.Cookies
		f.WriteLine(sCookieLog & Item & ":  " & request.cookies(Item))
	Next

	f.WriteLine("")
	' GET POST INFORMATION
	f.WriteLine("Form Data:")
	For each item in Request.Form
		f.WriteLine(sPostLog & item & ":  " &	 request.form(item))
	Next
	-->
	<%

	f.WriteLine("")
	' GET POST INFORMATION
	f.WriteLine("Form Body:")
	If Request.TotalBytes > 0 Then
    	Dim lngBytesCount
        	lngBytesCount = Request.TotalBytes
    	f.WriteLine(BytesToStr(Request.BinaryRead(lngBytesCount)))
		strJSONBody = BytesToStr(Request.BinaryRead(lngBytesCount))
	End If

	Set oJSON = New aspJSON

	'Load JSON string
	oJSON.loadJSON(jsonstring)

	f.WriteLine("")
	f.WriteLine("")
	f.WriteLine("")
	'Get single value
	Response.Write oJSON.data("firstName") & "<br>"
	

	f.WriteLine("################################################################################################")
%>
<!--
	f.WriteLine("")
	f.WriteLine("Server Variables:")
	For each item in Request.ServerVariables 
		f.WriteLine(sHTTPLog & item & ":  " &	 request.servervariables(item))
	Next
	

f.Close
set f=Nothing
set fs=Nothing
-->
<%

Function BytesToStr(bytes)
    Dim Stream
    Set Stream = Server.CreateObject("Adodb.Stream")
        Stream.Type = 1 'adTypeBinary
        Stream.Open
        Stream.Write bytes
        Stream.Position = 0
        Stream.Type = 2 'adTypeText
        Stream.Charset = "iso-8859-1"
        BytesToStr = Stream.ReadText
        Stream.Close
    Set Stream = Nothing
End Function
	
%>


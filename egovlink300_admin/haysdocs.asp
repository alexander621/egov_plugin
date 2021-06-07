<%

response.buffer = false
dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
		'ProcessFiles "IS NULL", "" 
	'response.end
sSQL = "SELECT * FROM haysfolders WHERE parentid IS null ORDER BY ModuleID"
set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1

'response.write "<ul>" & vbcrlf
Do While not oRs.EOF

	if not fs.folderexists("d:\haysdocs\" & oRs("ModuleID") & "\" & oRs("foldername")) then
		fs.CreateFolder("d:\haysdocs\" & oRs("ModuleID") & "\" & oRs("foldername"))
	end if

	'response.write "	<li>" & oRs("ModuleID") & "\" & oRs("foldername") & vbcrlf
		ProcessFolder oRs("folderid"), oRs("ModuleID") & "\" & oRs("foldername") & "\"
	'response.write "	</li>" & vbcrlf


	oRs.MoveNext
loop
'response.write "</ul>" & vbcrlf

oRs.Close
set oRs = Nothing

set f=nothing
set fs=nothing


Function ProcessFolder(intID, strFolderName)

'response.write "		<ul>" & vbcrlf
	ProcessFiles intID,strFolderName 

	sSQL = "SELECT * FROM haysfolders WHERE parentid = " & intID
	set oRsf = Server.CreateObject("ADODB.RecordSet")
	oRsf.Open sSQL, Application("DSN"), 3, 1
	Do While Not oRsf.EOF

		strLCLFolderName = replace(trim(oRsf("foldername")),"/","-")
		
		'response.write "			<li>" & strFolderName & strLCLFolderName
		if not fs.folderexists("d:\haysdocs\" & strFolderName & strLCLFolderName) then
			fs.CreateFolder("d:\haysdocs\" & strFolderName & strLCLFolderName)
			'response.write "(created)"
		else
			'response.write "(exists)"
		end if
		ProcessFolder oRsf("folderid"), strFolderName & strLCLFolderName & "\"
		'response.write "			</li>" & vbcrlf


		oRsf.MoveNext
	loop
	oRsf.Close
	set oRsf = Nothing

	

'response.write "		</ul>" & vbcrlf

	

End Function


Sub ProcessFiles(intID, strFolderName)
	'sSQL = "SELECT * FROM haysdocs WHERE folderid IS NULL"
	sSQL = "SELECT * FROM haysdocs_20190412 WHERE folderid = " & intID
	'response.write sSQL
	set oRsD = Server.CreateObject("ADODB.RecordSet")
	oRsD.Open sSQL, Application("DSN"), 3, 1
	intAppend = 0
	Do While Not oRsD.EOF
		response.write "<li>" & strFolderName & oRsD("originalFileName")
		strAppend = ""

		if fs.FileExists("d:\haysdocs\" & strFolderName & oRsD("originalfilename")) then 
			intAppend = intAppend + 1
			strAppend = "-" & intAppend
			'response.write strAppend
			if intAppend > 1 then response.write strFolderName & oRsD("originalFileName") & strAppend & "<br />"
		end if
		if fs.FileExists("d:\haysdocs\" & oRsD("serverfilename")) then

	 		fs.MoveFile "d:\haysdocs\" & oRsD("serverfilename"), "d:\haysdocs\" & strFolderName & oRsD("originalfilename") & strAppend
		end if
		response.write "</li>" & vbcrlf
		oRsD.MoveNext
	loop
	oRsD.Close
	set oRsD = Nothing


End Sub
%>

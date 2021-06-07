Sites using &gt; 1GB of space:<br />
<%
Dim fs,fo,x,fo2




'dim fs,fo,x
set fs=Server.CreateObject("Scripting.FileSystemObject")
set fo=fs.GetFolder("E:\egovlink300_docs\custom\pub\")
response.write "<table><tr><th>Name</th><th>Size (in GB)</th></tr>"
for each x in fo.SubFolders
  'Print the name of all subfolders in the test folder
  'Response.write(x.Name & "<br>")
	Set fo=fs.GetFolder("E:\egovlink300_docs\custom\pub\" & x.Name)
	if fo.size > 1000000000 then
		Response.Write("<tr><td>" & x.Name & "</td><td>" & round(fo.size/1000000000.00,2) & "</td></tr>")
	end if
	set fo=nothing
next
response.write "</table>"

set fo=nothing
set fs=nothing

%>

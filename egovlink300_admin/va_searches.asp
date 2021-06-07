  <link rel="stylesheet" type="text/css" href="global.css" />
<%
intOrgID = session("orgid")
if request.querystring("orgid") <> "" then
	intOrgID = request.querystring("orgid")
end if

URL = ""
'sSQL = "SELECT TOP 1 URL FROM alexaquerylog WHERE orgid = '" & intOrgID & "'"
sSQL = "SELECT TOP 1 configvalue AS URL FROM alexaconfigs WHERE configtype = 'faqURL' and orgid = '" & intOrgID & "'"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
if not oRs.EOF then 
	'URL = mid(oRs("URL"),1,instr(oRs("URL"),"/feed")-1)
	URL = oRs("URL")
end if
oRs.Close
Set oRs = Nothing


%>


<table id="bodytable" border="0" cellpadding="0" cellspacing="0" class="start">
  <tr valign="top">
    	<td>
	<h2>Voice Assistant Searches</h2>

<form action="<%=mid(URL,1,instr(URL,"?")-1)%>" id="searchform" method="get" role="search" target="_blank">
<%
	qs = replace(mid(URL,instr(URL,"?")+1),"&s=","")
	arrQS = split(qs,"&")
	for each item in arrQS
		arrItem = split(item,"=")
		name = arrItem(0)
		value = arrItem(1)
		response.write "<input type=""hidden"" name=""" & name & """ value=""" & value & """ />"
	next

%>
    <div><label for="s" class="screen-reader-text">Test A Search for:</label>
    <input type="text" id="s" name="s" value="">
    <input type="submit" value="Search" id="searchsubmit">

    </div>
</form>
<br />
<br />
<br />
	<%

	sSQL = "SELECT TOP 100 * FROM alexaquerylog WHERE orgid = '" & intOrgID & "' ORDER BY alexaquerylogid DESC"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1


	response.write "<table class=""tablelist"" cellspacing=""0"" cellpadding=""2"" border=""0"" style=""min-width:1000px"">"
	response.write "<tr><th>When</th><th>User Asked/<br /> Assistant Provided</th><th>Query Phrase to WordPress</th><th>Assistant Response to User</th><th>Search URL</th></tr>"
	Do While Not oRs.EOF
		response.write "<tr>"
		response.write "<td>" & oRs("logdate") & "</td>"
		response.write "<td>" & oRs("query") & "</td>"
		response.write "<td>" & oRs("shortquery") & "</td>"
		response.write "<td>" & oRs("reply") & "</td>"
		response.write "<td><a href=""" & oRs("URL") & """>link</a></td>"
		response.write "<tr>"
		oRs.MoveNext
	loop
	response.write "</table>"

	oRs.Close
	Set oRs = Nothing
	%>
      </td>
  </tr>
</table>

<script>
 function onElementHeightChange(elm, callback){
    var lastHeight = elm.clientHeight, newHeight;
    (function run(){
        newHeight = elm.clientHeight;
        if( lastHeight != newHeight )
            callback();
        lastHeight = newHeight;

        if( elm.onElementHeightChangeTimer )
            clearTimeout(elm.onElementHeightChangeTimer);

        elm.onElementHeightChangeTimer = setTimeout(run, 200);
    })();
}


	if (window.top!=window.self)
	{
		var height = document.body.scrollHeight;
		parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: true }},"*")

	}
 </script>


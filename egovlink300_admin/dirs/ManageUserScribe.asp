<%
dim conn,rs,strSQL,thisname,cmd,groupid,delimeter,sGroups,strResult,sBgcolor,sChecked0,sChecked1,sChecked2,bDisabled,CommitteeName,ResultID
const howoften0="No"
const howoften1="intantly"
const howoften2="Daily"
const howoften3="Weekly"
const howoften4="monthly"

set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 2 
strSQL = "select subscribeid,name,source from subscribe"
rs.Open strSQL,,, 2 

While Not rs.EOF
sOptions=  sOptions & "<tr bgcolor=""" & sBgcolor & """>"
sOptions=  sOptions & "<td>"&rs("name")&"</td>"&vbcrlf
sOptions = sOptions & "<td><input type=radio name='howoften"&rs("subscribeid")&"' "&sChecked0&" value="&howoften0&" >&nbsp;</td>"&vbcrlf
sOptions = sOptions & "<td><input type=radio name='howoften"&rs("subscribeid")&"' "&sChecked0&" value="&howoften1&" >&nbsp;</td>"&vbcrlf
sOptions = sOptions & "<td><input type=radio name='howoften"&rs("subscribeid")&"' "&sChecked1&" value="&howoften2&">&nbsp;</td>"&vbcrlf
sOptions = sOptions & "<td><input type=radio name='howoften"&rs("subscribeid")&"' "&sChecked2&" value="&howoften3&">&nbsp;</td>"&vbcrlf
sOptions = sOptions & "<td><input type=radio name='howoften"&rs("subscribeid")&"' "&sChecked3&" value="&howoften4&">&nbsp;</td></tr>"&vbcrlf
sOptions = sOptions & "</tr>"
rs.movenext
wend
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<form name="subscribe" action="insert_subscribe.asp" method="post" ID="subscribe">
<TABLE cellpadding=5 width=400 cellspacing=0 class="tablelist" ID="Table1">
<th align="left">Subscribed Items</th>
<th align="left"><%=howoften0%></th>
<th align="left"><%=howoften1%></th>
<th align="left"><%=howoften2%></th>
<th align="left"><%=howoften3%></th>
<th align="left"><%=howoften4%></th>
<%=sOptions%>
</TABLE>
<INPUT TYPE='submit' value='Update'>
</form>
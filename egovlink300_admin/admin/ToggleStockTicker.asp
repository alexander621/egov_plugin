<%
Dim sSql, oCnn

sSql = "EXEC ToggleStockTicker " & Session("UserID")

Set oCnn = Server.CreateObject("ADODB.Connection")
oCnn.Open Application("DSN")
oCnn.Execute sSql,,adExecuteNoRecords
oCnn.Close
Set oCnn = Nothing

Session("ShowStockTicker") = Abs(clng(Session("ShowStockTicker")-1))

If Request.QueryString("redirect") <> "" Then
  Response.Redirect Request.QueryString("redirect")
Else
  Response.Write "<html><body onload=""history.back();""></body></html>"
End If
%>

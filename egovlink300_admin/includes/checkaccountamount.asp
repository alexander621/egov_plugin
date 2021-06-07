<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkaccountamount.asp
' AUTHOR: Steve Loar	
' CREATED: 02/05/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the citizen account has enough money to cover the queried amount
'
' MODIFICATION HISTORY
' 1.0   02/05/2007   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oAmount, sResponse

sSql = "SELECT isnull(accountbalance, 0.00) as accountbalance FROM egov_users WHERE userid = " & request("uid")

Set oAmount = Server.CreateObject("ADODB.Recordset")
oAmount.Open sSQL, Application("DSN"), 3, 1

If NOT oAmount.EOF Then
	If CDbl(oAmount("accountbalance")) >= CDbl(request("amt")) Then 
		sResponse = "OK"
	Else
		sResponse = "FAILED"
	End If 
Else
	sResponse = "FAILED"
End If 

oAmount.close
Set oAmount = Nothing 

response.write sResponse

%>
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: order_items.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module reorders the News Items.
'
' MODIFICATION HISTORY
' 1.0   10/31/06	Steve Loar - Initial version.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim iNewOrder, oOrder

	iNewOrder = clng(request("itemorder")) + clng(request("iDirection"))
 iNewsType = request("newstype")
	
	sSQL = "Select newsitemid, itemorder FROM egov_news_items where orgid = " & Session("OrgID") & " order by itemorder"
	If clng(request("iDirection")) = clng(-1) Then
		sSql = sSql & " DESC"
	End If
	
	Set oOrder = Server.CreateObject("ADODB.Recordset")
	oOrder.CursorLocation = 3
	oOrder.Open sSQL, Application("DSN"), 1, 3

	Do While Not oOrder.EOF

		If oOrder("itemorder") = clng(request("itemorder")) Then
			oOrder("itemorder") = iNewOrder
		Else  
			If clng(oOrder("itemorder")) = clng(iNewOrder) Then
				oOrder("itemorder") = request("itemorder")
			End If 
		End if
		oOrder.Update
		oOrder.MoveNext
	Loop 
	oOrder.close
	Set oOrder = Nothing 

	' REDIRECT TO News Items page
	response.redirect "list_items.asp?newstype=" & iNewsType

%>

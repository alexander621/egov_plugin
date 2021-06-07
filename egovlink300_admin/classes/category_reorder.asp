<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: category_reorder.asp
' AUTHOR: Steve Loar
' CREATED: 11/22/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module reorders the Class Categories.
'
' MODIFICATION HISTORY
' 1.0   11/22/06	Steve Loar - Initial version.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iNewOrder, oOrder

iNewOrder = clng(request("sequenceid")) + clng(request("iDirection"))

sSQL = "SELECT categoryid, sequenceid FROM EGOV_CLASS_CATEGORIES WHERE orgid = " & Session("OrgID") & " ORDER BY sequenceid"
If iDirection = "-1" Then
	sSql = sSql & " DESC"
End If

Set oOrder = Server.CreateObject("ADODB.Recordset")
oOrder.CursorLocation = 3
oOrder.Open sSQL, Application("DSN"), 1, 3

Do While Not oOrder.EOF

	If oOrder("sequenceid") = clng(request("sequenceid")) Then
		oOrder("sequenceid") = iNewOrder
	Else  
		If clng(oOrder("sequenceid")) = clng(iNewOrder) Then
			oOrder("sequenceid") = request("sequenceid")
		End If 
	End if
	oOrder.Update
	oOrder.MoveNext
Loop 

oOrder.close
Set oOrder = Nothing 

' REDIRECT TO Class Category page
response.redirect("category_mgmt.asp")

%>

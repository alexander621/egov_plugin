<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: delete_items.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module deletes the News Scroller Items.
'
' MODIFICATION HISTORY
' 1.0 10/31/06	Steve Loar - Initial Version Created
' 1.6 07/09/09 David Boyer - Added "newstype" to split News and News Scroller items.

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Retrieve the newstype
 if request("newstype") <> "" then
    lcl_newstype = UCASE(request("newstype"))
 else
    lcl_newstype = "SCROLLER"
 end if

	Dim sSql, oCmd

'Delete Question
	sSQL = "DELETE FROM egov_news_items WHERE newsitemid = " & CLng(request("newsitemid"))
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSQL
		.Execute
	End With
	Set oCmd = Nothing

	ReorderNewsItems

'Redirect to item list
	response.redirect "list_items.asp?success=SD&newstype=" & lcl_newstype

'------------------------------------------------------------------------------
Sub ReorderNewsItems()
	Dim sSql, oOrder, iOrder

	iOrder = clng(0)
	sSql = "Select newsitemid, itemorder FROM egov_news_items where orgid = " & session("orgid") & " Order by itemorder"
	Set oOrder = Server.CreateObject("ADODB.Recordset")
	oOrder.CursorLocation = 3
	oOrder.Open sSQL, Application("DSN"), 1, 3

	Do While Not oOrder.EOF
		iOrder = iOrder + clng(1)
		oOrder("itemorder") = iOrder
		oOrder.Update 
		oOrder.MoveNext
	Loop

	oOrder.Close
	Set oOrder = Nothing 

End Sub 
%>

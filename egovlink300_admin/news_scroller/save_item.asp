<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: save_items.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists the News Scroller Items.
'
' MODIFICATION HISTORY
' 1.0 10/31/06	Steve Loar - Initial Version Created
' 1.1	12/04/07	Steve Loar - Added Pub start and end dates
' 1.2	02/25/08	Steve Loar - New items are now added to the top of the list and the list renumbered 
' 1.3 07/09/09 David Boyer - Added "newstype" to split News and News Scroller items.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim oCmd, iItemOrderNo, sPubStartDate, sPubEndDate

 if request("newstype") <> "" then
    sNewsType    = "'" & dbsafe(UCASE(request("newstype"))) & "'"
    lcl_newstype = UCASE(request("newstype"))
 else
    sNewsType    = "'SCROLLER'"
    lcl_newstype = "SCROLLER"
 end if

 if request("newsitemid") <> "" then
    lcl_newsitemid = CLng(request("newsitemid"))
 else
    lcl_newsitemid = 0
 end if

 if request("itemtitle") <> "" then
    sItemTitle = "'" & dbsafe(request("itemtitle")) & "'"
 else
    sItemTitle = "NULL"
 end if

 if request("itemdate") <> "" then
    sItemDate = "'" & request("itemdate") & "'"
 else
    sItemDate = "NULL"
 end if

 if request("itemtext") <> "" then
    sItemText = "'" & dbsafe(request("itemtext")) & "'"
 else
    sItemText = "NULL"
 end if

 if request("itemlinkurl") <> "" then
    sItemLinkURL = "'" & dbsafe(request("itemlinkurl")) & "'"
 else
    sItemLinkURL = "NULL"
 end if

	if request("publicationstart") = "" then
  		sPubStartDate = "NULL"
	else
  		sPubStartDate = "'" &	dbsafe(request("publicationstart")) & "'"
	end if

	if request("publicationend") = "" then
  		sPubEndDate = "NULL"
	else
  		sPubEndDate = "'" & dbsafe(request("publicationend")) & "'"
	end if

 if request("itemdisplay") <> "on" then
    sItemDisplay = 0
 else
    sItemDisplay = 1
 end if
	
	if CLng(lcl_newsitemid) <> CLng(0) then

	'The news item exists, so update it
		'Set oCmd = Server.CreateObject("ADODB.Command")
		'With oCmd
		'	.ActiveConnection = Application("DSN")
		'	sCommand = "UPDATE egov_news_items SET "
  ' sCommand = sCommand & " itemtitle = '"       & dbsafe(request("itemtitle"))   & "', "
  ' sCommand = sCommand & " itemdate = '"        & request("itemdate")            & "', "
		'	sCommand = sCommand & " itemtext = '"        & dbsafe(request("itemtext"))    & "', "
  ' sCommand = sCommand & " itemlinkurl = '"     & dbsafe(request("itemlinkurl")) & "', "
		'	sCommand = sCommand & " publicationstart = " & sPubStartDate                  & ", "
  ' sCommand = sCommand & " publicationend = "   & sPubEndDate
		'	sCommand = sCommand & " WHERE newsitemid = " & request("newsitemid")
		'	.CommandText = sCommand
		'	.Execute
		'End With
		'Set oCmd = Nothing

			sSQL = "UPDATE egov_news_items SET "
   sSQL = sSQL & " itemtitle = "        & sItemTitle    & ", "
   sSQL = sSQL & " itemdate = "         & sItemDate     & ", "
			sSQL = sSQL & " itemtext = "         & sItemText     & ", "
   sSQL = sSQL & " itemlinkurl = "      & sItemLinkURL  & ", "
			sSQL = sSQL & " publicationstart = " & sPubStartDate & ", "
   sSQL = sSQL & " publicationend = "   & sPubEndDate   & ", "
   sSQL = sSQL & " itemdisplay = "      & sItemDisplay  & ", "
   sSQL = sSQL & " newstype = "         & sNewsType
			sSQL = sSQL & " WHERE newsitemid = " & request("newsitemid")

 	set oUpdateNewsItems = Server.CreateObject("ADODB.Recordset")
 	oUpdateNewsItems.Open sSQL, Application("DSN"), 0, 1

  set oUpdateNewsItems = nothing

  lcl_success = "SU"

	else		'New News Item
	'Get the next itemorder number 
		iItemOrderNo = 0  'Put this at the head of the list. Will renumber after the insert

	'Insert the news item
		'Set oCmd = Server.CreateObject("ADODB.Command")
		'With oCmd
		'	.ActiveConnection = Application("DSN")
		'	sCommand = "INSERT INTO egov_news_items ("
  ' sCommand = sCommand & "orgid, "
  ' sCommand = sCommand & "itemtitle, "
  ' sCommand = sCommand & "itemdate, "
  ' sCommand = sCommand & "itemtext, "
  ' sCommand = sCommand & "itemlinkurl, "
  ' sCommand = sCommand & "itemdisplay, "
  ' sCommand = sCommand & "itemorder, "
  ' sCommand = sCommand & "publicationstart, "
  ' sCommand = sCommand & "publicationend"
  ' sCommand = sCommand & ") VALUES ("
		'	sCommand = sCommand &       session("orgid")                           & ", "
  ' sCommand = sCommand & "'" & dbready_string(request("itemtitle"),100)   & "', "
  ' sCommand = sCommand & "'" & request("itemdate")                        & "', "
  ' sCommand = sCommand & "'" & dbready_string(request("itemtext"),400)    & "', "
  ' sCommand = sCommand & "'" & dbready_string(request("itemlinkurl"),500) & "', "
		'	sCommand = sCommand &       "0, "
  ' sCommand = sCommand &       iItemOrderNo                               & ", "
  ' sCommand = sCommand &       sPubStartDate                              & ", "
  ' sCommand = sCommand &       sPubEndDate                                & " )"
		'	.CommandText = sCommand
		'	.Execute
		'End With
		'Set oCmd = Nothing

			sSQL = "INSERT INTO egov_news_items ("
   sSQL = sSQL & "orgid, "
   sSQL = sSQL & "itemtitle, "
   sSQL = sSQL & "itemdate, "
   sSQL = sSQL & "itemtext, "
   sSQL = sSQL & "itemlinkurl, "
   sSQL = sSQL & "itemorder, "
   sSQL = sSQL & "publicationstart, "
   sSQL = sSQL & "publicationend, "
   sSQL = sSQL & "itemdisplay, "
   sSQL = ssQL & "newstype "
   sSQL = sSQL & ") VALUES ("
			sSQL = sSQL & session("orgid") & ", "
   sSQL = sSQL & sItemTitle       & ", "
   sSQL = sSQL & sItemDate        & ", "
   sSQL = sSQL & sItemText        & ", "
   sSQL = sSQL & sItemLinkURL     & ", "
   sSQL = sSQL & iItemOrderNo     & ", "
   sSQL = sSQL & sPubStartDate    & ", "
   sSQL = sSQL & sPubEndDate      & ", "
   sSQL = sSQL & sItemDisplay     & ", "
   sSQL = sSQL & sNewsType
   sSQL = sSQL & ")"

 	'set oInsertNewsItems = Server.CreateObject("ADODB.Recordset")
 	'oInsertNewsItems.Open sSQL, Application("DSN"), 0, 1
 	lcl_newsitemid = RunIdentityInsertStatement(sSQL)

 'Re-number the news items.
		RenumberNewsItems

  'oInsertNewsItems.close
  'set oInsertNewsItems = nothing

  lcl_success = "SA"

	end if
   
'Check to see if there is any aditional processing we will need to do.
 lcl_return_parameters = ""

 if request("sendTo_RSS") = "on" then
    lcl_return_parameters = lcl_return_parameters & "&sendTo_RSS=" & lcl_newsitemid
 end if

'Return to the maintenance screen
 lcl_return_url = "edit_item.asp"
 lcl_return_url = lcl_return_url & "?newstype="   & lcl_newstype
 lcl_return_url = lcl_return_url & "&newsitemid=" & lcl_newsitemid
 lcl_return_url = lcl_return_url & "&success="    & lcl_success
 lcl_return_url = lcl_return_url & lcl_return_parameters

 response.redirect lcl_return_url

'------------------------------------------------------------------------------
Function DBsafe( strDB )
	if Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	strDB = Replace(strDB, Chr(13), "<br />" )
	strDB = Replace(strDB, Chr(10), "" )
	DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
Function GetNextItemOrderNo
	Dim sSql, oOrder 

	sSQL = "Select max(itemorder) as maxOrder FROM egov_news_items WHERE OrgID = " & Session("orgid") 

	Set oOrder = Server.CreateObject("ADODB.Recordset")
	oOrder.Open sSQL, Application("DSN"), 0, 1
	
	if oOrder.EOF Then
		GetNextItemOrderNo = 1
	else
		if IsNull(oOrder("maxOrder")) Then 
			GetNextItemOrderNo = 1
		else 
			GetNextItemOrderNo = clng(oOrder("maxOrder")) + 1
		end if
	End If
	oOrder.close 
	Set oOrder = Nothing

End Function 

'------------------------------------------------------------------------------
Sub RenumberNewsItems( )
	Dim sSql, oRs, x

	x = 0 
	sSql = "SELECT newsitemid FROM egov_news_items WHERE OrgID = " & Session("orgid") & " ORDER BY itemorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		x = x + 1
		sSql = "UPDATE egov_news_items SET itemorder = " & x & " WHERE newsitemid = " & oRs("newsitemid")
		RunSQL sSql 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 

'------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub
%>

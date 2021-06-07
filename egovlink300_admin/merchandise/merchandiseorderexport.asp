<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiseorderexport.asp
' AUTHOR: Steve Loar
' CREATED: 05/06/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Extract of selected regatta teams exported to EXCEL
'
' MODIFICATION HISTORY
' 1.0   05/06/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sOrdersPicks, sRptTitle, sSql, oRs, dTotal

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=MerchandiseOrders_" & sDate & ".xls"

sOrdersPicks = request("orderpicks")

sRptTitle = vbcrlf & "<tr height=""30""><th>Merchandise Orders</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

sSql = "SELECT O.orderdate, O.paymentid, O.merchandiseorderid, ISNULL(O.taxamount,0.00) AS taxamount, ISNULL(O.shippingfee,0.00) AS shippingfee,  "
sSql = sSql & " O.orderamount, O.shiptoname, O.shiptoaddress, O.shiptocity, O.shiptostate, O.shiptozip, "
sSql = sSql & " U.userfname, U.userlname, U.useraddress, U.usercity, U.userstate, U.userzip,  "
sSql = sSql & " I.merchandise, I.quantity, I.merchandisecolor, I.isnocolor, I.merchandisesize, I.isnosize, I.displayorder, I.itemprice "
sSql = sSql & " FROM egov_merchandiseorders O, egov_users U, egov_merchandiseorderitems I "
sSql = sSql & " WHERE O.userid = U.userid AND O.merchandiseorderid IN " & sOrdersPicks & " AND O.merchandiseorderid = I.merchandiseorderid "
sSql = sSql & " ORDER BY O.paymentid, O.merchandiseorderid, I.merchandise, I.merchandisecolor, I.displayorder "

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	response.write vbcrlf & "<html><body>"
	response.write vbcrlf & "<table cellpadding=""4"" cellspacing=""0"" border=""1"">"
	response.write sRptTitle
	response.write vbcrlf & "<tr height=""30""><th>Order Date</th><th>Receipt #</th><th>Order #</th>"
	response.write "<th>Merchandise Total</th><th>Shipping</th><th>Sales Tax</th><th>Total</th>"
	response.write "<th>Buyer Name</th><th>Buyer Address</th><th>Buyer City</th><th>Buyer State</th><th>Buyer Zip</th>"
	response.write "<th>Ship To Name</th><th>Ship To Address</th><th>Ship To City</th><th>Ship To State</th><th>Ship To Zip</th>"
	response.write "<th>Merchandise</th><th>Color</th><th>Size</th><th>Quantity</th><th>Item Price</th><th>Item Total</th>"
	response.write "</tr>"   
	response.flush
	Do While Not oRs.EOF
		response.write vbcrlf & "<tr height=""20"">"
		response.write "<td align=""center"">" & FormatDateTime(oRs("orderdate"),2) & "</td>"
		response.write "<td align=""center"">" & oRs("paymentid") & "</td>"
		response.write "<td align=""center"">" & oRs("merchandiseorderid") & "</td>"
		response.write "<td align=""center"">&nbsp;" & FormatNumber(oRs("orderamount"),2) & "</td>"
		response.write "<td align=""center"">&nbsp;" & FormatNumber(oRs("shippingfee"),2) & "</td>"
		response.write "<td align=""center"">&nbsp;" & FormatNumber(oRs("taxamount"),2) & "</td>"
		dTotal = CDbl(oRs("orderamount")) + CDbl(oRs("shippingfee")) + CDbl(oRs("taxamount"))
		response.write "<td align=""center"">&nbsp;" & FormatNumber(dTotal,2) & "</td>"
		response.write "<td align=""center"">" & oRs("userfname") & " " & oRs("userlname") & "</td>"
		response.write "<td align=""center"">" & oRs("useraddress") & "</td>"
		response.write "<td align=""center"">" & oRs("usercity") & "</td>"
		response.write "<td align=""center"">" & oRs("userstate") & "</td>"
		response.write "<td align=""center"">" & oRs("userzip") & "</td>"
		response.write "<td align=""center"">" & oRs("shiptoname") & "</td>"
		response.write "<td align=""center"">" & oRs("shiptoaddress") & "</td>"
		response.write "<td align=""center"">" & oRs("shiptocity") & "</td>"
		response.write "<td align=""center"">" & oRs("shiptostate") & "</td>"
		response.write "<td align=""center"">" & oRs("shiptozip") & "</td>"
		response.write "<td align=""center"">" & oRs("merchandise") & "</td>"
		response.write "<td align=""center"">"
		If oRs("isnocolor") Then 
			response.write "&nbsp;"
		Else 
			response.write oRs("merchandisecolor")
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If oRs("isnosize") Then 
			response.write "&nbsp;"
		Else 
			response.write oRs("merchandisesize") 
		End If 
		response.write "</td>"
		response.write "<td align=""center"">" & oRs("quantity") & "</td>"
		response.write "<td align=""center"">&nbsp;" & FormatNumber(oRs("itemprice"),2) & "</td>"
		dTotal = CDbl(oRs("itemprice")) * CLng(oRs("quantity"))
		response.write "<td align=""center"">&nbsp;" & FormatNumber(dTotal,2) & "</td>"
		response.write "</tr>"
		response.flush
		oRs.MoveNext
	Loop 
	response.write vbcrlf & "</table></body></html>"
	response.flush
End If 

oRs.Close
Set oRs = Nothing 

%>

<!-- #include file="../includes/common.asp" //-->



%>

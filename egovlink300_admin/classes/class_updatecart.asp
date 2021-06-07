<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_updatecart.asp
' AUTHOR: Steve Loar
' CREATED: 03/28/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module updates the changes in the quantities in the cart.
'
' MODIFICATION HISTORY
' 1.0   03/28/06   Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim x, sSql, oQty, iClassId, iQuantity, iChangeQty, iTimeId, sBuyOrWait, oChild

	' Loop through the items in the cart
	For x = 1 To clng(request("totalitems"))
		'response.write request("cartid." & x) & "<br />"

		If IsNumeric(request("quantity." & x)) Then 
			' Ignore non-numeric quantities
			
			' Get the original qty, timeid and buyorwait
			sSql = "Select classid, quantity, classtimeid, buyorwait, isdropin from egov_class_cart where cartid = " & request("cartid." & x)
			Set oQty = Server.CreateObject("ADODB.Recordset")
			oQty.CursorLocation = 3
			oQty.Open sSQL, Application("DSN"), 1, 3

			If Not oQty.EOF Then 
				iClassId = oQty("classid")
				iQuantity = clng(oQty("quantity"))
				iTimeId = Clng(oQty("classtimeid"))
				sBuyOrWait = oQty("buyorwait")
				bIsDropIn = oQty("isdropin")

				If clng(request("quantity." & x)) > 0 Then 
					' if original qty is not equal to the current qty
					If clng(iQuantity) <> clng(request("quantity." & x)) Then 

						' The difference between the new and old quantities
						iChangeQty = clng(request("quantity." & x)) - iQuantity
						'response.write vbcrlf & "iChange = " &  iChangeQty & "<br />"
						
						If Not bIsDropIn Then 
							' Update egov_class_time to have correct enrollment counts
							UpdateClassTime iTimeId, iChangeQty, sBuyOrWait
						End If 

						' Update egov_class_cart
						oQty("quantity") = clng(request("quantity." & x))
					End If 
					oQty.Update
					oQty.close
					Set oQty = Nothing
					
					If Not bIsDropIn Then 
						' Update the series children to have the correct enrollment
						UpdateSeriesChildren iClassId, iChangeQty, sBuyOrWait
					End If 

				Else
					'iChangeQty = - iQuantity
					oQty.close
					Set oQty = Nothing
					If clng(request("quantity." & x)) = 0 Then 
						' if they changed the quantity to 0 remove it
						RemoveItemFromCart request("cartid." & x), iTimeId, sBuyOrWait, bIsDropIn
					End If 
				End If 

			End If 
		End If 

	Next 

	' Recalculate the prices in the cart
	ResetCartPrices

	If OrgHasFeature("discounts") then
		' Recalculate any discounts
		DetermineDiscounts
	
	End If 

	' Return to the cart page
	response.redirect "class_cart.asp"

%>

<!--#Include file="class_global_functions.asp"-->  

<!-- #include file="../includes/common.asp" //-->


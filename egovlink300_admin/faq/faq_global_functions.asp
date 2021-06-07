<%
'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub reorderFAQCategories(iOrgID,iFAQType)

	sSQL = "SELECT FAQCategoryID, displayorder "
 sSQL = sSQL & " FROM faq_categories "
 sSQL = sSQL & " WHERE OrgID = "         & iOrgID
 sSQL = sSQL & " AND UPPER(faqtype) = '" & iFAQType & "' "
	sSQL = sSQL & " ORDER BY displayorder"

	set oCatOrder = Server.CreateObject("ADODB.Recordset")
	oCatOrder.Open sSQL, Application("DSN"), 0, 1

 if not oCatOrder.eof then

    iRowCount = 0

    do while not oCatOrder.eof
    			iRowCount = iRowCount + 1

     		sSQL = "UPDATE faq_categories SET displayorder = " & iRowCount & " WHERE FAQCategoryID = " & oCatOrder("FAQCategoryID")
      	set oUpdateCatOrder = Server.CreateObject("ADODB.Recordset")
     		oUpdateCatOrder.Open sSQL, Application("DSN"), 3, 1

       set oUpdateCatOrder = nothing

     		oCatOrder.movenext
    loop
 end if

	oCatOrder.close
	set oCatOrder = nothing

end sub
%>
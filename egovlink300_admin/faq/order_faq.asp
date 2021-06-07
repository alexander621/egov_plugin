<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: order_faq.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module reorders the Frequently Asked Questions (FAQ).
'
' MODIFICATION HISTORY
' 1.?   09/12/06   Steve Loar - Changes for categories.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
subOrderFaqs request("ifaqid"),request("orgid"),UCASE(request("direction")),request("faqcategoryid"), request("faqtype")

'------------------------------------------------------------------------------
sub subOrderFaqs( iFaqID, iOrgID, sDirection, iFaqCategoryId, iFAQType )

	Dim oCmd, sSQL, sFaqCategoryId, oOrder, iNumberOfFaqs
	
 if clng(iFaqCategoryId) = clng(0) then
  		sFaqCategoryId = " IS NULL "
 else
  		sFaqCategoryId = " = " & iFaqCategoryId
 end if

 if iFAQType = "" then
    iFAQType = "FAQ"
 end if

	iSequence     = 0
	iNumberOfFaqs = 0

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM FAQ "
 sSQL = sSQL & " WHERE orgid = "     & iorgID
 sSQL = sSQL & " AND faqcategoryid " & sFaqCategoryId
 sSQL = sSQL & " AND UPPER(faqtype) = '" & UCASE(iFAQType) & "' "
 sSQL = sSQL & " ORDER BY Sequence, FaqID"

	set oOrder = Server.CreateObject("ADODB.Recordset")
	oOrder.Open sSQL, Application("DSN"), 0, 1

'REPLACE ANY NULL SEQUENCE WITH CURRENT SEQUENCE
 if not oOrder.eof then
    do while not oOrder.eof
       iSequence     = iSequence + 1
    			iNumberOfFaqs = iNumberOfFaqs + 1

    			if CLng(iFaqID) = CLng(oOrder("FaqID")) then
      				iCurrentSequence = iSequence
       end if

    			sSQL = "UPDATE FAQ SET Sequence = " & iSequence & " WHERE FaqID = " & oOrder("faqid")
    			RunSQL sSQL

    			oOrder.movenext
    loop
 end if

	oOrder.close
	set oOrder = nothing

'Process Question Move
 select case sDirection
    case "UP"
       iNewSequence = iCurrentSequence - 1

       if iNewSequence < 1 then
          iNewSequence = 1
       end if
    case "DOWN"
       iNewSequence = iCurrentSequence + 1

       if iNewSequence > iNumberOfFaqs then
       			iNewSequence = iNumberOfFaqs
       end if
    case "TOP"
       iNewSequence = 0
    case "BOTTOM"
       iNewSequence = iNumberOfFaqs + 1
 end select

	if iNewSequence <> iCurrentSequence then
    sSQL = "UPDATE FAQ SET sequence = "     & iCurrentSequence
    sSQL = sSQL & " WHERE orgid = "         & iOrgID
    sSQL = sSQL & " AND sequence = "        & iNewSequence
    sSQL = sSQL & " AND faqcategoryid "     & sFaqCategoryId
    sSQL = sSQL & " AND UPPER(faqtype) = '" & UCASE(iFAQType) & "'"
  		RunSQL sSQL

  		sSQL = "UPDATE FAQ SET sequence = "     & iNewSequence
    sSQL = sSQL & " WHERE orgid = "         & iOrgID
    sSQL = sSQL & " AND FaqID = "           & iFaqID
    sSQL = sSQL & " AND faqcategoryid "     & sFaqCategoryId
    sSQL = sSQL & " AND UPPER(faqtype) = '" & UCASE(iFAQType) & "'"
  		RunSQL sSQL
 end if

 response.redirect "list_faq.asp?faqtype=" & iFAQType & "&success=SR"

end sub

'------------------------------------------------------------------------------
sub RunSQL(sSql)
	Dim oCmd

	set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	set oCmd = nothing

end sub
%>

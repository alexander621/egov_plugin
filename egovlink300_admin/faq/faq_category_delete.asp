<!-- #include file="../includes/common.asp" //-->
<!-- #include file="faq_global_functions.asp" //-->
<%	
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: faq_category_delete.asp
' AUTHOR: Steve Loar
' CREATED: 09/11/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module deletes FAQ categories
'
' MODIFICATION HISTORY
' 1.0 09/11/06 Steve Loar - Original code
' 1.1 07/22/09 David Boyer - Added the "reorder"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim oCmd

 lcl_success = ""

 if request("FAQCategoryID") <> "" then
    if isnumeric(request("FAQCategoryID")) then

       if request("faqtype") <> "" then
          lcl_faqtype = UCASE(request("faqtype"))
       else
          lcl_faqtype = "FAQ"
       end if

      'Set all of the FAQs that are associated to this FAQ Category to NULL.
       sSQL = "UPDATE faq SET FAQCategoryId = NULL WHERE FAQCategoryId = " & request("FAQCategoryId")

      	set oDelFAQs = Server.CreateObject("ADODB.Recordset")
     		oDelFAQs.Open sSQL, Application("DSN"), 3, 1

      'Delete the FAQ Category.
       sSQL = "DELETE FROM faq_categories WHERE FAQCategoryID = " & request("FAQCategoryID")

      	set oDelFAQCat = Server.CreateObject("ADODB.Recordset")
     		oDelFAQCat.Open sSQL, Application("DSN"), 3, 1

       set oDelFAQs   = nothing
       set oDelFAQCat = nothing

       lcl_success = "SD"

      'Reorder the Categories
       reorderFAQCategories session("orgid"), lcl_faqtype
    end if
 end if

 response.redirect "faq_categories.asp?faqtype=" & lcl_faqtype & "&success=" & lcl_success
%>
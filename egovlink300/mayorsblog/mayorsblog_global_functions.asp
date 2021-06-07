<%
'------------------------------------------------------------------------------
sub displayArchives( ByVal p_orgid )
	Dim sSQL, oBlogDates

 'Retreive a distinct list of createdbydates from egov_mayorsblog.
  sSQL = "SELECT distinct DATEPART(mm,mb.createdbydate) AS blogMonth, DATEPART(yyyy,mb.createdbydate) as blogYear "
  sSQL = sSQL & " FROM egov_mayorsblog mb "
  sSQL = sSQL & " WHERE mb.isInactive = 0 "
  sSQL = sSQL & " AND mb.orgid = " & p_orgid
  sSQL = sSQL & " ORDER BY 2 DESC, 1 DESC "

 	set oBlogDates = Server.CreateObject("ADODB.Recordset")
  oBlogDates.Open sSQL, Application("DSN"), 3, 1

  if not oBlogDates.eof then
     do while not oBlogDates.eof

        response.write "<a href=""mayorsblog.asp?blogMonth=" & oBlogDates("blogMonth") & "&blogYear=" & oBlogDates("blogYear") & """>" & monthname(oBlogDates("blogMonth")) & " " & oBlogDates("blogYear") & "</a><br />" & vbcrlf

        oBlogDates.movenext
     loop
  end if

  oBlogDates.close
  set oBlogDates = nothing

end sub

'------------------------------------------------------------------------------
function buildBlogImg( ByVal iImage, ByVal iVirtualSiteName)
	Dim lcl_return
  lcl_return = ""

  if trim(iImage) <> "" then
        lcl_imagefilename = iImage

        if left(lcl_imagefilename,1) <> "/" then
           lcl_imagefilename = "/" & lcl_imagefilename
        end if

        lcl_return = Application("CommunityLink_DocUrl")
        lcl_return = lcl_return & "public_documents300/"
        lcl_return = lcl_return & iVirtualSiteName
        lcl_return = lcl_return & "/unpublished_documents"
        lcl_return = lcl_return & lcl_imagefilename

        'lcl_return = sEgovWebsiteURL
        'lcl_return = lcl_return & "/admin/custom/pub/"
        'lcl_return = lcl_return & iVirtualSiteName
        'lcl_return = lcl_return & "/unpublished_documents"
        'lcl_return = lcl_return & iImage
  end if

  buildBlogImg = lcl_return

end function
%>
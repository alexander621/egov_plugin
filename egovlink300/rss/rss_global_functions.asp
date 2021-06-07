<%
'------------------------------------------------------------------------------
sub buildRSSFeed(iFeedName)

 dim lcl_feedname
 lcl_feedname = ""
 lcl_print_pubdate = true

 if iFeedName <> "" then
    lcl_feedname = dbsafe(iFeedName)
    if instr(lcl_feedname, ",NOPUBDATE") > 0 then
        lcl_print_pubdate = false
        lcl_feedname = left(lcl_feedname, instr(lcl_feedname, ",NOPUBDATE") - 1)
    end if
 end if

'BEGIN: Build RSS Header ------------------------------------------------------
 lcl_header_pubdate  = formatPublicationDate(now())

 buildRSSHeader "", lcl_feedname, lcl_header_pubdate
'END: Build RSS Header --------------------------------------------------------

'BEGIN: Build RSS List --------------------------------------------------------
 sSQL = "SELECT r.rssid, r.rowid, r.title, r.description, r.rsslink, r.publicationdate, "
 sSQL = sSQL & " r.createdbyid, r.createdbyname, r.feedid, f.feedname, f.feature "
 sSQL = sSQL & " FROM egov_rss r, egov_rssfeeds f "
 sSQL = sSQL & " WHERE r.feedid = f.feedid "
 sSQL = sSQL & " AND f.isActive = 1 "
 sSQL = sSQL & " AND r.orgid = " & iorgid
 sSQL = sSQL & " AND UPPER(f.feedname) = '" & lcl_feedname & "' "
 sSQL = sSQL & " AND DATEDIFF(d,publicationdate,'" & date() & "') <= 14 "
 sSQL = sSQL & " ORDER BY r.publicationdate DESC "

 set oRSSFeeds = Server.CreateObject("ADODB.Recordset")
	oRSSFeeds.Open sSQL, Application("DSN"), 3, 1

 i = 0
 if not oRSSFeeds.eof then
    if orghasfeature(iorgid,oRSSFeeds("feature")) then
       do while not oRSSFeeds.eof

          lcl_rssTitle = formatXML(oRSSFeeds("title"))
          lcl_rssLink  = sEgovWebsiteURL & oRSSFeeds("rsslink")
          lcl_rssDesc  = formatXML(oRSSFeeds("description"))
          lcl_pubdate  = formatPublicationDate(oRSSFeeds("publicationdate"))

          i = i + 1

          response.write "        <item>" & vbcrlf
          if lcl_print_pubdate then
              response.write "            <title>"       & lcl_rssTitle & "[" & lcl_pubdate & "]</title>"       & vbcrlf
          else
              response.write "            <title>"       & lcl_rssTitle & "</title>"       & vbcrlf
          end if
          response.write "            <description>" & lcl_rssDesc  & "</description>" & vbcrlf
	  if instr(lcl_rssLink,"news") > 0 then
          	response.write "            <link>"        & lcl_rssLink & "?" & "id=" & oRSSFeeds("rowid")  & "</link>"        & vbcrlf
	  else
          	response.write "            <link>"        & lcl_rssLink & "</link>"        & vbcrlf
	  end if
          response.write "            <pubDate>"     & lcl_pubdate  & "</pubDate>"     & vbcrlf
          response.write "            <guid isPermaLink=""true"">" & lcl_rssLink & "/item" & oRSSFeeds("rowid") & month(oRSSFeeds("publicationdate")) & day(oRSSFeeds("publicationdate")) & year(oRSSFeeds("publicationdate")) & "</guid>" & vbcrlf
          response.write "        </item>" & vbcrlf

          oRSSFeeds.movenext
       loop
    end if
 end if

 oRSSFeeds.close
 set oRSSFeeds = nothing
'END: Build RSS List ----------------------------------------------------------

 response.write "    </channel>" & vbcrlf
 response.write "</rss>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub buildRSSHeader(iFeedID,iFeedName,iPubDate)

  dim lcl_feedname
  lcl_feedname = ""

  if iFeedName <> "" then
     lcl_feedname = dbsafe(iFeedName)
  end if

  sSQL = "SELECT isnull(t.orgtitle,f.title) AS title, f.description, f.feedurl, f.lastbuilddate, f.feature "
  sSQL = sSQL & " FROM egov_rssfeeds f "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_rssfeeds_orgtitles t ON f.feedid = t.feedid AND t.orgid = " & iorgid
  sSQL = sSQL & " WHERE f.isActive = 1 "

  if iFeedID <> "" then
     sSQL = sSQL & " AND f.feedid = " & iFeedID
  else
     sSQL = sSQL & " AND f.feedname = '" & lcl_feedname & "'"
  end if

  set oFeed = Server.CreateObject("ADODB.Recordset")
	 oFeed.Open sSQL, Application("DSN"), 3, 1

  if not oFeed.eof then
     if orghasfeature(iorgid,oFeed("feature")) then
        lcl_feedtitle     = sOrgName & " - " & oFeed("title")
        lcl_feeddesc      = sOrgName & " - " & oFeed("description")
        lcl_feedurl       = sEgovWebsiteURL & oFeed("feedurl")
        lcl_lastbuilddate = formatPublicationDate(oFeed("lastbuilddate"))
     else
        response.redirect sEgovWebsiteURL  'include_top_functions.asp
     end if
'  else
     'lcl_feedtitle     = sOrgName & " - " & iFeedName & " RSS News"
     'lcl_feeddesc      = ""
     'lcl_feedurl       = sEgovWebsiteURL
     'lcl_lastbuilddate = ""


'     response.redirect sEgovWebsiteURL  'include_top_functions.asp
  end if

  oFeed.close
  set oFeed = nothing

  response.ContentType="text/xml"
  response.write "<?xml version=""1.0"" encoding=""ISO-8859-1"" ?>" & vbcrlf
  response.write "<rss version=""2.0"">" & vbcrlf
  response.write "    <channel>" & vbcrlf

 'Required
  response.write "        <title>"       & lcl_feedtitle & "</title>" & vbcrlf
  response.write "        <link>"        & lcl_feedurl   & "</link>" & vbcrlf
  response.write "        <description>" & lcl_feeddesc  & "</description>" & vbcrlf

 'Optional
  response.write "        <category>E-Government Services/Government</category>" & vbrlf
  response.write "        <copyright>2004-" & year(date) & " electronic commerce link, inc. dba egovlink</copyright>" & vbcrlf
  response.write "        <language>en-US</language>" & vbcrlf
  response.write "        <docs>http://www.w3schools.com/rss/rss_intro.asp</docs>" & vbcrlf
  response.write "        <lastBuildDate>" & lcl_lastbuilddate & "</lastBuildDate>" & vbcrlf
  response.write "        <pubDate>" & iPubDate & "</pubDate>" & vbcrlf
  response.write "        <generator>EditPlus</generator>" & vbcrlf
  'response.write "        <webMaster>E-Gov Support &lt;egovsupport@eclink.com&gt;</webMaster>" & vbcrlf
  'response.write "        <image>" & vbcrlf
  'response.write "            <url>http://www.diehardpaintball.com/images/headers/main_logo_3.jpg</url>" & vbcrlf
  'response.write "            <title>E-Gov Services - FAQs RSS</title>" & vbcrlf
  'response.write "            <link>http://www.egovlink.com/eclink/faq.asp</link>" & vbcrlf
  'response.write "            <width>144</width>" & vbcrlf
  'response.write "            <height>75</height>" & vbcrlf
  'response.write "        </image>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function formatXML(iValue)
  lcl_return = ""

  if iValue <> "" then
     lcl_return = iValue

     lcl_return = replace(lcl_return,"<<AMP>>","&amp;")
     lcl_return = replace(lcl_return,"<p>","&lt;p&gt;")
     lcl_return = replace(lcl_return,"</p>","&lt;/p&gt;")
     lcl_return = replace(lcl_return,"<br>","&lt;br /&gt;")
     lcl_return = replace(lcl_return,"<br />","&lt;br /&gt;")
     lcl_return = replace(lcl_return,"<P>","&lt;p&gt;")
     lcl_return = replace(lcl_return,"</P>","&lt;/p&gt;")
     lcl_return = replace(lcl_return,"<BR>","&lt;br /&gt;")
     lcl_return = replace(lcl_return,"<BR />","&lt;br /&gt;")

     lcl_return = replace(lcl_return,"<","&lt;")
     lcl_return = replace(lcl_return,">","&gt;")
     'lcl_return = replace(lcl_return,"&","&amp;")
     lcl_return = replace(lcl_return,"’","&apos;")
     lcl_return = replace(lcl_return,"'","&apos;")
     lcl_return = replace(lcl_return,"‘","&apos;")
     lcl_return = replace(lcl_return,"""","&quot;")

  end if

  formatXML = lcl_return

end function

'------------------------------------------------------------------------------
function formatPublicationDate(iPubDate)
  lcl_return = ""

  if iPubDate <> "" then
     if isDate(iPubDate) then
       'Format the Publication Date - ie. Thu, 28 Apr 2006
        lcl_return = WeekdayName(Weekday(iPubDate),true)
        lcl_return = lcl_return & ", "
        lcl_return = lcl_return & day(iPubDate)
        lcl_return = lcl_return & " "
        lcl_return = lcl_return & monthname(month(iPubDate),true)
        lcl_return = lcl_return & " "
        lcl_return = lcl_return & year(iPubDate)
        lcl_return = lcl_return & " "
        lcl_return = lcl_return & hour(iPubDate)
        lcl_return = lcl_return & ":"
        lcl_return = lcl_return & minute(iPubDate)
        lcl_return = lcl_return & ":"
        lcl_return = lcl_return & second(iPubDate)
     end if
  end if

  formatPublicationDate = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,"'","''")
  end if

  dbsafe = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>

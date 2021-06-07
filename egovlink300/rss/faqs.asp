<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp"-->
<!-- #include file="../class/classOrganization.asp" -->
<!-- #include file="rss_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'-- Validate feeds at: http://www.feedvalidator.org ---------------------------
'------------------------------------------------------------------------------
'Retrieve all of the RSS Feeds.
 lcl_feedname = "FAQ"

 buildRSSFeed lcl_feedname
%>
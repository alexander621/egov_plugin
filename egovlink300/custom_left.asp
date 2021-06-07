<%

sPath = Server.MapPath("custom_html/custom_left_" & iOrgID & ".asp")

 response.write ReadFile(sPath)

%>
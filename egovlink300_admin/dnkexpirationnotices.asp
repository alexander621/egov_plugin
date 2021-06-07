<!-- #include file="../egovlink300_global/includes/inc_email.asp" //-->
<%

SendReminders 7
SendReminders 0


Sub SendReminders(days)

	    sSql = "SELECT useremail, u.orgid,donotknockregdate, subject, body " _
                        & " FROM egov_users u " _
                        & " INNER JOIN Organizations o ON o.OrgID = u.orgid " _
                        & " INNER JOIN DoNotKnockOrgMessages dnkom ON dnkom.orgid = o.OrgID and dnkom.daysmatch = " & days & " " _
                        & " WHERE  " _
                        & " donotknockregdate >= DATEADD(d," & days & ",DATEADD(YYYY,-1*o.donotknockexpiration, Convert(DateTime, DATEDIFF(DAY, 0, GETDATE()))))  " _
                        & " and donotknockregdate < DateAdd(d," & (days + 1) & ",DATEADD(YYYY,-1*o.donotknockexpiration, Convert(DateTime, DATEDIFF(DAY, 0, GETDATE())))) "
                        '& " donotknockregdate IS NOT NULL "

	    Set oRs = Server.CreateObject("ADODB.RecordSet")
	    oRs.Open sSql, Application("DSN"), 3, 1

	    Do While not oRs.EOF
	    	email = oRs("useremail")
		subject = oRs("subject")
		body = oRs("body")
		sendEmail "",email,"",subject,body,"",true 

	    	oRs.MoveNext
	    loop

                 
End Sub
%>

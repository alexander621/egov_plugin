using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Helpers;

public partial class handle_email_issues : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
		string sSQL = "";
		string iNewId = "";
		string emailAddress = "";
		string logemailAddress = "";
		string ActionTaken = "";
		string VD = "";
		string OrgID = "";
		string fromAddress = "";
		int removedCount = 0;
		string messageType = "";
		//UNTESTED CODE TO READ JSON FROM RESPONSE BODY
		string json = new System.IO.StreamReader(Request.InputStream).ReadToEnd();

		//CODE TO SIMULATE GETTING JSON STRING
		/*
		StringBuilder sb = new StringBuilder();
		using (StreamReader sr = new StreamReader("D:\\wwwroot\\egov\\email_handling\\json.txt")) 
		{
    		String line;
    		// Read and display lines from the file until the end of 
    		// the file is reached.
    		while ((line = sr.ReadLine()) != null) 
    		{
        		sb.AppendLine(line);
    		}
		}
		string json = sb.ToString();
		//END
		*/

		if (!String.IsNullOrEmpty(json))
		{
			//CODE TO LOG SNS MESSAGE AND ACTION TAKEN
			sSQL = "INSERT INTO zEmailHandlingLog (SNSMessage) VALUES('" + dbSafe(json) + "')";
        		SqlConnection sqlConn2 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        		SqlCommand sqlCommander2 = new SqlCommand();
        		sqlCommander2.Connection = sqlConn2;
	
        		sqlConn2.Open();
	
        		sqlCommander2.CommandText = sSQL + ";SELECT @@IDENTITY";
        		iNewId = Convert.ToString(sqlCommander2.ExecuteScalar());
		
        		sqlConn2.Close();
        		sqlCommander2.Dispose();
		}
		
		var obj = Json.Decode(json);

		bool valid = true;
		try
		{
			var message = Json.Decode(obj.Message);
		}
		catch
		{
			//Nothing
			valid = false;
		}

		if (valid)
		{
			var message = Json.Decode(obj.Message);

			//Get OrgID
			fromAddress = message.mail.source;
			if (fromAddress.IndexOf("_") > 0)
			{
				VD = fromAddress.Substring(fromAddress.IndexOf("_")+1, fromAddress.IndexOf("@") - (fromAddress.IndexOf("_")+1));

				sSQL = "SELECT OrgID FROM Organizations WHERE OrgVirtualSiteName = '" + VD + "'";
        			SqlConnection sqlConnOrg = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        			sqlConnOrg.Open();

        			SqlCommand myCommand = new SqlCommand(sSQL, sqlConnOrg);
        			SqlDataReader myReader;
        			myReader = myCommand.ExecuteReader();
	
        			if (myReader.HasRows)
        			{
            				while (myReader.Read())
            				{
						OrgID = myReader["OrgID"].ToString();
					}
				}
        			myReader.Close();
        			sqlConnOrg.Close();
        			myReader.Dispose();
        			sqlConnOrg.Dispose();

			}


			if (message.notificationType == "Bounce")
			{
				//else if (message.bounce.bounceType == "Transient" && (message.bounce.bounceSubType == "General") && (message.bounce.bouncedRecipients[0].status == "5.7.1" || message.bounce.bouncedRecipients[0].status == "5.3.0" || message.bounce.bouncedRecipients[0].status == "5.1.0" || message.bounce.bouncedRecipients[0].status == "4.4.7"))
				if (message.bounce.bounceType == "Permanent" && (message.bounce.bounceSubType == "General" || message.bounce.bounceSubType == "Suppressed")) // && (message.bounce.bouncedRecipients[0].status == "5.1.1" || message.bounce.bouncedRecipients[0].status == "5.3.0"))
				{
					emailAddress = message.bounce.bouncedRecipients[0].emailAddress;
					ActionTaken = "Removed";
				}
				else if (message.bounce.bounceType == "Transient" && (message.bounce.bounceSubType == "General" || message.bounce.bounceSubType == "ContentRejected"))
				{
					emailAddress = message.bounce.bouncedRecipients[0].emailAddress;
					ActionTaken = "Removed - Transient";
				}
				else if (message.bounce.bounceType == "Undetermined" && message.bounce.bounceSubType == "Undetermined")
				{
					emailAddress = message.bounce.bouncedRecipients[0].emailAddress;
					ActionTaken = "Removed - Undetermined";
				}
				else if (message.bounce.bounceSubType == "MailboxFull")
				{
					ActionTaken = "Removed - MailBoxFull";
					emailAddress = message.bounce.bouncedRecipients[0].emailAddress;
				}
				messageType = "Bounce";
			}
			else if (message.notificationType == "Complaint")
			{
				emailAddress = message.complaint.complainedRecipients[0].emailAddress;

				ActionTaken = "Removed";
				messageType = "Complaint";

			}

			//Get Address for all cases for logging purposes
			try
			{
				logemailAddress = message.complaint.complainedRecipients[0].emailAddress;

			}
			catch
			{
				try
				{
					logemailAddress = message.bounce.bouncedRecipients[0].emailAddress;
				}
				catch
				{
					logemailAddress = "UNAVAILABLE";
				}
			}


			//CODE TO REMOVE EMAIL ADDRESS
			if (!String.IsNullOrEmpty(emailAddress))
			{
				//Add to suppression list
				sSQL = "SELECT TOP 1 emailsuppressionid FROM emailsuppressionlist WHERE emailaddress = '" + emailAddress + "'";
 				SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        			sqlConn.Open();
			
        			SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        			SqlDataReader myReader = myCommand.ExecuteReader();
			
				int newSuppression = 0;
        			if (!myReader.HasRows)
        			{
					sSQL = "INSERT INTO emailsuppressionlist (emailaddress) VALUES('" + emailAddress + "')";
        				SqlConnection sqlConn8 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        				sqlConn8.Open();
        				SqlCommand sqlCommander8 = new SqlCommand();
        				sqlCommander8.Connection = sqlConn8;
        				sqlCommander8.CommandText = sSQL;
        				newSuppression = sqlCommander8.ExecuteNonQuery();
					sqlConn8.Close();
					sqlConn8.Dispose();
        				sqlCommander8.Dispose();
				}
        			myReader.Close();
        			sqlConn.Close();
        			myReader.Dispose();
        			sqlConn.Dispose();

	
	
				
				sSQL = "DELETE ";
  					sSQL += " FROM egov_class_distributionlist_to_user ";
    					sSQL += " WHERE EXISTS (SELECT userid FROM egov_users WHERE egov_class_distributionlist_to_user.userid = egov_users.userid AND useremail = '" + emailAddress + "')";
        			sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        			SqlCommand sqlCommander = new SqlCommand();
        			sqlCommander.Connection = sqlConn;
		
        			sqlConn.Open();
		
        			sqlCommander.CommandText = sSQL;
        			removedCount = sqlCommander.ExecuteNonQuery();
		
        			sqlConn.Close();
        			sqlCommander.Dispose();



				//REMOVE FROM ALL ACTION LINE ITEMS TOO
				sSQL = "SELECT u.UserID, action_autoid, u.OrgID, alr.status, eu.userid as uid ";
					sSQL += " FROM egov_actionline_requests alr ";
					sSQL += " INNER JOIN Users u ON u.OrgID = alr.OrgID ";
					sSQL += " INNER JOIN egov_users eu ON eu.userid = alr.userid ";
					sSQL += " WHERE u.Username = 'eclink' and eu.useremail = '" + emailAddress + "' ";
        			sqlConn.Open();
			
        			myCommand = new SqlCommand(sSQL, sqlConn);
        			myReader = myCommand.ExecuteReader();
			
				int clearUser = 0;
				int recClear = 0;
        			if (myReader.HasRows)
        			{

            				while (myReader.Read())
            				{
						removedCount++;


						sSQL = "UPDATE egov_users SET useremail = '' WHERE userid = '" + myReader["uid"].ToString() + "'";
        					SqlConnection sqlConn2 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        					sqlConn2.Open();
        					SqlCommand sqlCommander2 = new SqlCommand();
        					sqlCommander2.Connection = sqlConn2;
        					sqlCommander2.CommandText = sSQL;
        					clearUser = sqlCommander2.ExecuteNonQuery();
						sqlConn2.Close();
						sqlConn2.Dispose();
        					sqlCommander2.Dispose();
	
						sSQL = "INSERT INTO egov_action_responses (action_internalcomment, action_editdate, action_userid, action_autoid, action_orgid, action_status) ";
  							sSQL += " VALUES('Edit Contact Information<br />Email: \"" + emailAddress + "\" changed to \"\"<br /> Invalid Email Removed', GetDate(), '" + myReader["userid"].ToString() + "', '" + myReader["action_autoid"].ToString() + "', '" + myReader["orgid"].ToString() + "', '" + myReader["status"].ToString() + "')";
        					SqlConnection sqlConn3 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        					sqlConn3.Open();
        					SqlCommand sqlCommander3 = new SqlCommand();
        					sqlCommander3.Connection = sqlConn3;
        					sqlCommander3.CommandText = sSQL;
        					recClear = sqlCommander3.ExecuteNonQuery();
						sqlConn3.Close();
						sqlConn3.Dispose();
        					sqlCommander3.Dispose();

					}
				}
        			myReader.Close();
        			sqlConn.Close();
        			myReader.Dispose();
        			sqlConn.Dispose();

				//REMOVE FROM ALL ACTION LINE NOTIFICATIONS & ESCALLATIONS (LEAVE MAIN ASSIGNMENT FOR NOW)
				sSQL = "SELECT userid FROM users WHERE email = '" + emailAddress + "'";
 				sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        			sqlConn.Open();
			
        			myCommand = new SqlCommand(sSQL, sqlConn);
        			myReader = myCommand.ExecuteReader();
			
				int clrUser2 = 0;
				int clrUser3 = 0;
				int delnotifications = 0;
				int delescalations = 0;
        			if (myReader.HasRows)
        			{
            				while (myReader.Read())
            				{
						removedCount++;

						string uid = myReader["userid"].ToString();

						sSQL = "UPDATE egov_action_request_forms SET assigned_userID2 = 0 WHERE assigned_userID2 = '" + uid + "'";
        					SqlConnection sqlConn4 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        					sqlConn4.Open();
        					SqlCommand sqlCommander4 = new SqlCommand();
        					sqlCommander4.Connection = sqlConn4;
        					sqlCommander4.CommandText = sSQL;
        					clrUser2 = sqlCommander4.ExecuteNonQuery();
						sqlConn4.Close();
						sqlConn4.Dispose();
        					sqlCommander4.Dispose();

						sSQL = "UPDATE egov_action_request_forms SET assigned_userID3 = 0 WHERE assigned_userID3 = '" + uid + "'";
        					SqlConnection sqlConn5 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        					sqlConn5.Open();
        					SqlCommand sqlCommander5 = new SqlCommand();
        					sqlCommander5.Connection = sqlConn5;
        					sqlCommander5.CommandText = sSQL;
        					clrUser3 = sqlCommander5.ExecuteNonQuery();
						sqlConn5.Close();
						sqlConn5.Dispose();
        					sqlCommander5.Dispose();

						sSQL = "DELETE FROM egov_action_notifications WHERE sendto = '" + uid + "'";
        					SqlConnection sqlConn6 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        					sqlConn6.Open();
        					SqlCommand sqlCommander6 = new SqlCommand();
        					sqlCommander6.Connection = sqlConn6;
        					sqlCommander6.CommandText = sSQL;
        					delnotifications = sqlCommander6.ExecuteNonQuery();
						sqlConn6.Close();
						sqlConn6.Dispose();
        					sqlCommander6.Dispose();

						sSQL = "DELETE FROM egov_action_escalations WHERE escNotify = '" + uid + "'";
        					SqlConnection sqlConn7 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        					sqlConn7.Open();
        					SqlCommand sqlCommander7 = new SqlCommand();
        					sqlCommander7.Connection = sqlConn7;
        					sqlCommander7.CommandText = sSQL;
        					delescalations = sqlCommander7.ExecuteNonQuery();
						sqlConn7.Close();
						sqlConn7.Dispose();
        					sqlCommander7.Dispose();
					}
				}
        			myReader.Close();
        			sqlConn.Close();
        			myReader.Dispose();
        			sqlConn.Dispose();



	
				if (message.notificationType == "Bounce")
				{
					//NOTE TO THE USER THAT EMAIL ADDRESS IS INVALID AND MUST BE CHANGED?
				}
				else if (message.notificationType == "Complaint")
				{
					//NOTE TO EMAIL ADDRESS THAT THEY'VE BEEN REMOVED AND ACCOUNT CLOSED?
				}
				
			}
		}

		//Response.Write(emailAddress);



		if (removedCount == 0 && !String.IsNullOrEmpty(ActionTaken) && ActionTaken != "MailboxFull")
		{
			ActionTaken = "NO SUBSCRIPTIONS";

			SendFullEmail("no-reply@eclink.com", "devsupport@eclink.com", "NO SUBSCRIPTIONS AWS Notification Message", json);
		}
		if (String.IsNullOrEmpty(ActionTaken))
		{
			//CODE TO EMAIL DEV FOR SNS REQUEST UNHANDLED CASE
			ActionTaken = "Unhandled";

			SendFullEmail("no-reply@eclink.com", "devsupport@eclink.com", "Unhandled AWS Notification Message", json);
		}

		if (!String.IsNullOrEmpty(json))
		{
			sSQL = "UPDATE zEmailHandlingLog SET ";
			sSQL += " EmailAddress = '" + dbSafe(logemailAddress) + "',";
			sSQL += " ActionTaken = '" + dbSafe(ActionTaken) + "',";
			sSQL += " virtualdirectoryname = '" + dbSafe(VD) + "',";
			sSQL += " orgid = '" + dbSafe(OrgID) + "',";
			sSQL += " fromAddress = '" + dbSafe(fromAddress) + "',";
			sSQL += " removedCount = '" + removedCount + "',";
			sSQL += " MessageType = '" + messageType + "' ";
			sSQL += " WHERE rowid = " + iNewId;


        		SqlConnection sqlConn2 = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        		SqlCommand sqlCommander2 = new SqlCommand();
        		sqlCommander2.Connection = sqlConn2;
	
        		sqlConn2.Open();
	
        		sqlCommander2.CommandText = sSQL;
        		iNewId = Convert.ToString(sqlCommander2.ExecuteScalar());
		
        		sqlConn2.Close();
        		sqlCommander2.Dispose();
		}
    }

    public void SendFullEmail(string fromAddress, string toAddress, string Subject, string Body, string fromName = "", string ccAddress = "", string bCCAddress = "", string priority = "", Boolean isHTMLFormat = true)
    {
            if (isHTMLFormat)
            {
                Body = "<html><body>" + Body + "</body></html>";
            }

            MailMessage message = new MailMessage();
            message.From = new MailAddress(fromAddress, fromName);
            message.To.Add(toAddress);
            message.Subject = Subject;
            message.Body = Body;
            message.IsBodyHtml = isHTMLFormat;

            if (ccAddress != "")
                message.CC.Add(ccAddress);
            if (bCCAddress != "")
                message.Bcc.Add(bCCAddress);
            if (priority == "high")
                message.Priority = MailPriority.High;


            SmtpClient smtp = new SmtpClient(ConfigurationManager.AppSettings["MailServer"]);
            try
            {
                smtp.Send(message);
            }
            //catch (SmtpException e)
            catch
            {
                // NF LOG ERROR
                //return e.Message;
            }
    }

    public static string dbSafe(string _Value)
    {
        string sNewString;
        sNewString = _Value.Replace("'", "''");
        sNewString = sNewString.Replace("<", "&lt;");
        return sNewString;
    }
	
}



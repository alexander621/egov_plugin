<%@ Page Language="C#" AutoEventWireup="true" ValidateRequest="false" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Net.Mail" %>
<%@ Import Namespace="System.IO" %>

<script runat="server">

/// <summary>
///
/// Process Payment
/// Filename: ProcessPayment.aspx
/// Author: Steve Loar
/// Created: 12/14/2012
/// Copyright: 2012 ECLink
/// Description: Taken from the process_payment.asp script to handle .net class purchases.
///				 This processes the class payments. It should handle $0 tranactions,
///				 Full CC transactions, Full Account Credit transactions and Partial
///				 CC and Account Credit Transactions.
/// 
/// Modifications:
///
/// </summary>
/// 

    public struct CartItem
    {
        public string PurchaserUserId;
        public string CartId;
        public string ClassId;
        public string AttendeeUserId;
        public string FamilyMemberId;
        public string BuyOrWait;
        public string Status;
        public string Quantity;
        public string ClassTimeId;
        public string ItemTypeId;
        public string PriceTypeId;
        public double Amount;
        public string RosterGrade;
        public string RosterShirtSize;
        public string RosterPantsSize;
        public string RosterCoachType;
        public string RosterVolunteerCoachName;
        public string RosterVolunteerCoachDayPhone;
        public string RosterVolunteerCoachCellPhone;
        public string RosterVolunteerCoachEmail;
    }
    
    
    public struct PayeeInfo
    {
        public string FirstName;
        public string LastName;
        public string Email;
        public string CardNumber;
        public int ExpMonth;
        public int ExpYear;
        public Boolean RequireCVV;
        public string CVV;
        public string Address;
        public string City;
        public string State;
        public string Zip;
        public string Details; // activity numbers being purchased
    }

    protected void Page_Load( object sender, EventArgs e )
    {



        //HttpCookie sCookieUserID = new HttpCookie("useridx");
        HttpCookie sCookieUserID = new HttpCookie("userid");
        
        int processControlNumber = 0;
        Int32 sProcessingErrorID = 0;
        
        string sProcessingErrorMsg = "";
        string paymentId = "0";
        string PNREFValue = "";
        string respMsg = "";
        string authCode = "";
        string orderNumber = "";
        string SVA = "";
        double feeAmount = 0;
        PayeeInfo myPayeeInfo;
        string gatewayErrorId;
        string sProcessingPath = "";
        string sPaymentName = "";
        string displayCardNumber = "";

        // Main Thread ******************************************************
        // get the itemid that is the link to the cart
        string sessionId;
        sessionId = "";
        sPaymentName = "";
        //sessionId = Request["itemnumber"].ToString();
        //sPaymentName = Request["paymentname"].ToString();

        if (Request["itemnumber"] != null)
        {
            sessionId = Request["itemnumber"].ToString();
        }

        if (Request["paymentname"] != null)
        {
            sPaymentName = Request["paymentname"].ToString();
        }

        //Response.Write( "SessionId = " + sessionId + "<br /><br />" );

        // check if the cart has items and if not return to the cart
        Boolean cartHasItems = classes.classCartHasItems( sessionId );
        //Response.Write( "cartHasItems = " + cartHasItems.ToString( ) + "<br />" );

        // get the orgid
        string orgId = common.getOrgId( );
        //Response.Write( "orgId = " + orgId.ToString( ) + "<br />" );






        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PAYMENT PROCESSING STARTED" );
        //Response.Write( "processControlNumber = " + processControlNumber.ToString( ) + "<br /><br />" );

        if (cartHasItems == true)
        {
            // pull some org features needed
            Boolean displayCVV = common.orgHasFeature( orgId, "display cvv" );
            Boolean customRegistrationCraigco = common.orgHasFeature( orgId, "custom_registration_CraigCO" );
            Boolean allowsAccountPayments = common.orgHasFeature( orgId, "public account payments" );
            Boolean skipPaymentProcessing = common.orgHasFeature( orgId, "skippayment" );
            
            // set the purchaserid
            common.makePaymentLogEntry(ref processControlNumber, orgId, "public", "classes/events", "request.form.userid: " + Request.Form["userid"].ToString());
            int buyerUserId;
            int.TryParse( Request.Form["userid"].ToString( ), out buyerUserId );
            common.makePaymentLogEntry(ref processControlNumber, orgId, "public", "classes/events", "buyerUserId: " + buyerUserId.ToString());

            // get the item total amount
            double cartTotal = 0;
            cartTotal = classes.getCartTotalAmount( sessionId );
            //Response.Write( "cartTotal = " + cartTotal.ToString( "C" ) + "<br />" );
            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "cartTotal: " + cartTotal.ToString( "C" ) );

            // set the apply credit flag. THis is paying with the citizen account balance
            Boolean applyCredit = false;
            
            if (Request["applycredit"].ToString() == "yes")
            {
                applyCredit = true;
            }

            double chargeableTotal = 0;
            double accountCredit = 0;
            double currentAccountBalance = 0;

            // if apply credit 
            if (applyCredit && allowsAccountPayments)
            {
                // get the citizen's current account balance
                currentAccountBalance = common.getCitizenAccountBalance( buyerUserId.ToString( ) );
                accountCredit = currentAccountBalance;
                
                if (accountCredit < 0)
                {
                    accountCredit = 0;
                    applyCredit = false;
                }
                else
                {
                    if (accountCredit > cartTotal)
                        accountCredit = cartTotal;
                }
                if (accountCredit == 0)
                    applyCredit = false;

                if (applyCredit)
                    chargeableTotal = cartTotal - accountCredit;
                else
                {
                    chargeableTotal = cartTotal;
                    applyCredit = false;
                }
            }
            else
            {
                // they do not allow citizens to apply their account balance, or the citizen elected not to apply it
                chargeableTotal = cartTotal;
                applyCredit = false;
            }

            if (chargeableTotal > 0)
	    {

		//NEED TO VERIFY CAPTCHA
		string strResponse = Request["g-recaptcha-response"].ToString();
		string strIP = Request.ServerVariables["REMOTE_HOST"].ToString();
		string strSecret = "6LcVxxwUAAAAAGGp_29X6bpiJ8YsWeNXinuUz6sx";
		using (WebClient client = new WebClient())
   		{
	
       			byte[] response =
       			client.UploadValues("https://www.google.com/recaptcha/api/siteverify", new NameValueCollection()
       			{
           			{ "secret", strSecret },
           			{ "response", strResponse },
           			{ "remoteip", strIP }
       			});
	
       			string result = System.Text.Encoding.UTF8.GetString(response);
			if (result.IndexOf("\"success\": true") < 0)
			{
            			Response.Redirect( common.getOrgFullSite( orgId ) + "/rd_classes/class_cart.aspx?message=timeout" );
			}
   		}
	    }
	

            //Response.Write( "accountCredit = " + accountCredit.ToString( "C" ) + "<br />" );
            //Response.Write( "chargeableTotal = " + chargeableTotal.ToString( "C" ) + "<br /><br />" );
            //Response.Write("applyCredit = " + applyCredit.ToString() + "<br />");
            
            if (applyCredit)
            {
                common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "Account Credit to apply: " + accountCredit.ToString( "C" ) );
                common.makePaymentLogEntry(ref processControlNumber, orgId, "public", "classes/events", "Prior Account Balance: " + currentAccountBalance.ToString("C"));
            }
            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "chargeableTotal: " + chargeableTotal.ToString( "C" ) );

            if (chargeableTotal == 0)
            {
                common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "Non-Credit Card Purchase" );
                
                // This happens when the cart totals $0, -OR- when the account credit covers the entire purchase
                // This will include wait list only purchases
                myPayeeInfo.CardNumber = "";
                myPayeeInfo.RequireCVV = false;
                myPayeeInfo.CVV = "";
                myPayeeInfo.ExpMonth = 0;
                myPayeeInfo.ExpYear = 0;
                myPayeeInfo.Email = "";
                myPayeeInfo.FirstName = "";
                myPayeeInfo.LastName = "";
                myPayeeInfo.Address = "";
                myPayeeInfo.City = "";
                myPayeeInfo.State = "";
                myPayeeInfo.Zip = "";
                myPayeeInfo.Details = "";

                // pull out their name, email and address from their egov_users record
                getPayeeInfo( buyerUserId, ref myPayeeInfo );
                common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "Payee: " + myPayeeInfo.FirstName + " " + myPayeeInfo.LastName );

                myPayeeInfo.Details = getCartDetails( sessionId );
                
                // There is no payment processor to talk to, so skip to processing the transaction as successful
                paymentId = processSuccessfulTransaction( orgId, sessionId, buyerUserId, cartTotal, 0, accountCredit, "NULL", currentAccountBalance, "NULL", "NULL", "NULL", "NULL", "NULL", processControlNumber );
            }
            else
            {
                // some of this must go through a payment processor. 
                
                // fill in the payee info from the information they provided on the payment form
                myPayeeInfo.FirstName = Request["firstname"].ToString( );
                myPayeeInfo.LastName = Request["lastname"].ToString( );
                common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "Payee: " + myPayeeInfo.FirstName + " " + myPayeeInfo.LastName );
                myPayeeInfo.Email = Request["email"].ToString( );
                myPayeeInfo.CardNumber = Request["accountnumber"].ToString( );
                //int.TryParse( Request["month"].ToString( ), out myPayeeInfo.ExpMonth );
                myPayeeInfo.ExpMonth = int.Parse( Request["month"].ToString( ) );
                //int.TryParse( Request["year"].ToString( ), out myPayeeInfo.ExpYear );
                myPayeeInfo.ExpYear = int.Parse( Request["year"].ToString( ) );
                if (displayCVV)
                {
                    myPayeeInfo.RequireCVV = true;
                    myPayeeInfo.CVV = "" + Request["cvv2"].ToString( );
                }
                else
                {
                    myPayeeInfo.RequireCVV = false;
                    myPayeeInfo.CVV = "";
                }
                myPayeeInfo.Address = Request["streetaddress"].ToString( );
                myPayeeInfo.City = Request["city"].ToString( );
                myPayeeInfo.State = Request["state"].ToString( );
                myPayeeInfo.Zip = Request["zipcode"].ToString( );
                myPayeeInfo.Details = Request["details"].ToString( ); // the activity ids of what is being purchased
                                
                //  If the org has the skip payment feature (for demos only)
                if (skipPaymentProcessing)
                {
                    common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "TRANSACTION SUCCEEDED: Bypassed Authorization" );
                    //Response.Write( "TRANSACTION SUCCEEDED: Bypassed Authorization<br /><br />" );
                    
                    //  set some variables to static values so things look real, and we can easily spot these fakes in the data
                    PNREFValue = "V19F1D5C82TEST";
                    respMsg = "Approved";
                    authCode = "010101";

                    // log the static values
                    common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PNREF: " + PNREFValue.ToString( ) );
                    common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "RESPMSG: " + respMsg.ToString( ) );
                    common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "AUTHCODE: " + authCode.ToString( ) );
                    //Response.Write( "PNREFValue = " + PNREFValue.ToString( ) + "<br />" );
                    //Response.Write( "RESPMSG = " + respMsg.ToString( ) + "<br />" );
                    //Response.Write( "AUTHCODE = " + authCode.ToString( ) + "<br /><br />" );
                    
                    //  Create rows for the successful non-payment
                    paymentId = processSuccessfulTransaction( orgId, sessionId, buyerUserId, cartTotal, chargeableTotal, accountCredit, "NULL", currentAccountBalance, authCode, PNREFValue, respMsg, "NULL", "NULL", processControlNumber );
                }
                else
                {
                    // get the processing route for the org
                    string processingRoute = common.getPaymentProcessingRoute( orgId );
                    //Response.Write( "processingRoute = " + processingRoute.ToString( ) + "<br /><br />" );

                    if (processingRoute == "PayPalPayFlowPro")
                    {
                        // PayPal Pay Flow Pro processing required
                        string PayPalResponse;
                        int PayPalResultCode;

                        sProcessingPath = "paypal";
                        
                        PayPalResultCode = ProcessPayPalTransaction( myPayeeInfo, processControlNumber, orgId, chargeableTotal, out PayPalResponse );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "RESULT: " + PayPalResultCode.ToString( ) );

                        PNREFValue = common.getPaymentProcessorResponseValue( PayPalResponse, "PNREF" );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PNREF: " + PNREFValue.ToString( ) );

                        respMsg = common.getPaymentProcessorResponseValue( PayPalResponse, "RESPMSG" );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "RESPMSG: " + respMsg.ToString( ) );

                        //Response.Write( "PNREFValue = " + PNREFValue.ToString( ) + "<br />" );
                        //Response.Write( "RESPMSG = " + respMsg.ToString( ) + "<br />" );
                        
                        // for testing of failures, the result code and messages are set here. For PayPal, you can also test it by inputing different card numbers, amounts, and CVV codes
                        //PayPalResultCode = -1; // Communication Error
                        //respMsg = "Communication Error Test";
                        //PayPalResultCode = 1; // Password Changed
                        //respMsg = "Password Changed Error Test";
                        //PayPalResultCode = 12; // Some other error
                        //respMsg = "Some Other Error including declined cards Test";
                        
                        if (PayPalResultCode == 0)
                        {
                            // TRANSACTION SUCCEEDED
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "TRANSACTION SUCCEEDED" );

                            authCode = common.getPaymentProcessorResponseValue( PayPalResponse, "AUTHCODE" );
                            //Response.Write( "AUTHCODE = " + authCode.ToString( ) + "<br /><br />" );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "AUTHCODE: " + authCode.ToString( ) );
                            //processSuccessfulTransaction( string _OrgId, string _SessionId, int _BuyerUserId, double _CartTotal, double _ChargeableTotal, double _AccountCredit, string _FeeAmount, string _AuthCode, string _PNRef, string _RespMsg, string _OrderNumber, string _SVA, int _ProcessControlNumber )
                            paymentId = processSuccessfulTransaction( orgId, sessionId, buyerUserId, cartTotal, chargeableTotal, accountCredit, "NULL", currentAccountBalance, authCode, PNREFValue, respMsg, "NULL", "NULL", processControlNumber );
                        }
                        
                        if (PayPalResultCode < 0)
                        {
                            // Communication Error with PayPal
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "COMMUNICATION ERROR" );
                            // input into the error table
                            gatewayErrorId = common.savePaymentProcessingError( orgId, "PayPal", "classes/events", "process payment", "Communication Error - " + respMsg, chargeableTotal.ToString( "C" ) );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "gatewayErrorId: " + gatewayErrorId );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PAYMENT PROCESSING FINISHED." );
                            // take them to a failure page
                            sProcessingErrorMsg += "Your credit card purchase was not processed because of a network ";
                            sProcessingErrorMsg += "communication error.  Please try your transaction again later.";
                            sProcessingErrorMsg += "&nbsp;";
                            sProcessingErrorMsg += "Payment Reference Number: [" + PNREFValue + "] - ";
                            sProcessingErrorMsg += "&nbsp;";
                            sProcessingErrorMsg += "Description: (" + PayPalResultCode.ToString() + ") - " + respMsg;

                            sProcessingErrorID = common.saveProcessingPaymentErrorMsg(sProcessingPath,
                                                                                      Convert.ToInt32(orgId),
                                                                                      buyerUserId,
                                                                                      Convert.ToInt32(gatewayErrorId),
                                                                                      sProcessingErrorMsg,
                                                                                      PNREFValue,
                                                                                      Convert.ToString(PayPalResultCode),
                                                                                      respMsg,
                                                                                      chargeableTotal,
                                                                                      orderNumber,
                                                                                      SVA,
                                                                                      authCode,
                                                                                      sessionId,
                                                                                      sPaymentName,
                                                                                      myPayeeInfo.CardNumber,
                                                                                      myPayeeInfo.FirstName + " " + myPayeeInfo.LastName,
                                                                                      myPayeeInfo.Address,
                                                                                      myPayeeInfo.City,
                                                                                      myPayeeInfo.State,
                                                                                      myPayeeInfo.Zip);
                                                    
                            //Response.Redirect( "../payment_processors/processing_failure.aspx?ge=" + gatewayErrorId );
                            Response.Redirect("ProcessFailure.aspx?&p=" + sProcessingErrorID.ToString());
                        }
                        
                        if (PayPalResultCode > 0)
                        {
                            // Transaction Declined
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "TRANSACTION DECLINED" );
                            if (PayPalResultCode == 1)
                            {
                                common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PayPal Account Password has changed" );
                                // send emails out about password being changed
                                sendLoginFailedEmail( orgId );
                            }

                            // input into the error table
                            gatewayErrorId = common.savePaymentProcessingError( orgId, "PayPal", "classes/events", "process payment", "Transaction Declined (" + PayPalResultCode.ToString( ) + ") - " + respMsg, chargeableTotal.ToString( "C" ) );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "gatewayErrorId: " + gatewayErrorId );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PAYMENT PROCESSING FINISHED." );
                            // take them to a failure page
                            //Response.Write("LEFT OFF HERE!!!");
                            
                            //Response.Write("modify table...add all columns to table in ProcessDeclinedTransaction.  Put in errorMsg column the value of ""ProcessDeclinedTransaction"".  If this is the value then ProcessFailure.aspx knows to show all of the fields....If not then simply display the errorMsg column like it does now.");
                            
                            //Response.Redirect( "../payment_processors/processing_failure.aspx?ge=" + gatewayErrorId );
                            //Response.Redirect("ProcessFailure.aspx?ge=" + gatewayErrorId.ToString());

                            sProcessingErrorMsg = "processdeclinedtransaction";
                            sProcessingErrorID = common.saveProcessingPaymentErrorMsg(sProcessingPath,
                                                                                      Convert.ToInt32(orgId),
                                                                                      buyerUserId,
                                                                                      Convert.ToInt32(gatewayErrorId),
                                                                                      sProcessingErrorMsg,
                                                                                      PNREFValue,
                                                                                      Convert.ToString(PayPalResultCode),
                                                                                      respMsg,
                                                                                      chargeableTotal,
                                                                                      orderNumber,
                                                                                      SVA,
                                                                                      authCode,
                                                                                      sessionId,
                                                                                      sPaymentName,
                                                                                      myPayeeInfo.CardNumber,
                                                                                      myPayeeInfo.FirstName + " " + myPayeeInfo.LastName,
                                                                                      myPayeeInfo.Address,
                                                                                      myPayeeInfo.City,
                                                                                      myPayeeInfo.State,
                                                                                      myPayeeInfo.Zip);

                            //Response.Redirect( "../payment_processors/processing_failure.aspx?ge=" + gatewayErrorId );
                            Response.Redirect("ProcessFailure.aspx?&p=" + sProcessingErrorID.ToString());
                            
                        }
                    }
                    else
                    {
                        //  Point & Pay processing required
                        string PNP_Response;
                        string statusCode;
                        string errorMsg;

                        sProcessingPath = "PointAndPay";
                        
                        PNP_Response = ProcessPointAndPayTransaction( myPayeeInfo, processControlNumber, orgId, chargeableTotal );
                        //Response.Write( "PNP_Response = [" + PNP_Response.ToString( ) + "]<br />" );
                        
                        // get the status
                        statusCode = common.getPaymentProcessorResponseValue( PNP_Response, "status" );
                        //Response.Write( "statusCode = [" + statusCode + "]<br />" );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "status: " + statusCode );

                        errorMsg = common.getPaymentProcessorResponseValue( PNP_Response, "errors" );
                        //Response.Write( "errorMsg = [" + errorMsg + "]<br />" );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "errors: " + errorMsg );

                        SVA = common.getPaymentProcessorResponseValue( PNP_Response, "sva" );
                        //Response.Write( "SVA = [" + SVA + "]<br />" );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "sva: " + SVA );

                        orderNumber = common.getPaymentProcessorResponseValue( PNP_Response, "orderNumber" );
                        //Response.Write( "orderNumber = [" + orderNumber + "]<br />" );
                        common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "orderNumber: " + orderNumber );

                        // This is to test failed transactions
                        //statusCode = "Declined"; // Some error
                        //errorMsg = "Some Error including declined cards Test";

                        if (statusCode.ToLower( ) != "success")
                        {
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "TRANSACTION FAILED" );
                            // input into the error table
                            gatewayErrorId = common.savePaymentProcessingError( orgId, "PayPal", "classes/events", "process payment", "Transaction Failed - " + errorMsg, chargeableTotal.ToString( "C" ) );
                            //Response.Write( "gatewayErrorId = [" + gatewayErrorId + "]<br />" );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "gatewayErrorId: " + gatewayErrorId );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PAYMENT PROCESSING FINISHED." );
                            // take them to a failure page
                            //Response.Redirect( "../payment_processors/processing_failure.aspx?ge=" + gatewayErrorId );
                            //Response.Redirect("ProcessFailure.aspx?ge=" + gatewayErrorId.ToString());

                            sProcessingErrorMsg = "processdeclinedtransaction";
                            sProcessingErrorID = common.saveProcessingPaymentErrorMsg(sProcessingPath,
                                                                                      Convert.ToInt32(orgId),
                                                                                      buyerUserId,
                                                                                      Convert.ToInt32(gatewayErrorId),
                                                                                      sProcessingErrorMsg,
                                                                                      PNREFValue,
                                                                                      "declined",
                                                                                      respMsg,
                                                                                      chargeableTotal,
                                                                                      orderNumber,
                                                                                      SVA,
                                                                                      authCode,
                                                                                      sessionId,
                                                                                      sPaymentName,
                                                                                      myPayeeInfo.CardNumber,
                                                                                      myPayeeInfo.FirstName + " " + myPayeeInfo.LastName,
                                                                                      myPayeeInfo.Address,
                                                                                      myPayeeInfo.City,
                                                                                      myPayeeInfo.State,
                                                                                      myPayeeInfo.Zip);
                            
                            Response.Redirect("ProcessFailure.aspx?&p=" + sProcessingErrorID.ToString());
                            
                        }
                        else
                        {
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "TRANSACTION SUCCEEDED" );
                            
                            // get the fee amount
                            string feeValue;
                            feeValue = common.getPaymentProcessorResponseValue( PNP_Response, "fee" );
                            if (feeValue != "")
                                double.TryParse( feeValue, out feeAmount );
                            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "feeAmount: " + feeAmount.ToString( "F2" ) );
                            
                            paymentId = processSuccessfulTransaction( orgId, sessionId, buyerUserId, cartTotal, chargeableTotal, accountCredit, feeAmount.ToString( "F2" ), currentAccountBalance, "NULL", "NULL", "NULL", orderNumber, SVA, processControlNumber );
                        }
                        
                    }

                } // end if (skipPaymentProcessing or Process Payment required)

            } // end if (chargeableTotal == 0 or a chargeable total > 0), free things and account credit cover cost are the $0 things

            // process the cart here - make ledger entries and put them on the roster, etc. 
            ProcessCartItems( sessionId, paymentId, buyerUserId, processControlNumber, orgId );
            //Response.Write( "Cart Processed.<br /><br />" );

            // send out emails to city and citizen. We still need the cart at this point for the emails
            sendEmailNotifications( sessionId, buyerUserId.ToString( ), orgId, paymentId, cartTotal, chargeableTotal, accountCredit, authCode, PNREFValue, myPayeeInfo, orderNumber, SVA, feeAmount );
            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "Emails Sent." );
            //Response.Write( "Emails sent.<br /><br />" );

            // clear the cart now that is has been processed and emails have been sent
            clearTheCart( sessionId );
            //Response.Write( "Cart Cleared.<br /><br />" );

            //Response.Write( "PAYMENT PROCESSING FINISHED.<br /><br />" );
            common.makePaymentLogEntry( ref processControlNumber, orgId, "public", "classes/events", "PAYMENT PROCESSING FINISHED." );
            
            //NOTE: In case of error, processing AFTER payment received display payment information to user
            if (myPayeeInfo.CardNumber != "")
            {
                displayCardNumber = myPayeeInfo.CardNumber.Substring(12);
            }
            
            Response.Write("<div>");
            Response.Write("<h2>Payment Processing Failed!</h2>");
            Response.Write("<strong>Your credit card was charged, but our application failed to process your payment.</strong>");
            Response.Write("<table border=\"0\">");
            Response.Write(  "<tr>");
            Response.Write(      "<td>AUTHORIZATION NUMBER:</td>");
            Response.Write(      "<td>" + authCode + "</td>");
            Response.Write(  "</tr>");
            Response.Write(  "<tr>");
            Response.Write(      "<td>PAYMENT REFERENCE NUMBER:</td>");
            Response.Write(      "<td>" + PNREFValue + "</td>");
            Response.Write(  "</tr>");
            Response.Write(  "<tr>");
            Response.Write(      "<td>MESSAGE:</td>");
            Response.Write(      "<td>" + respMsg + "</td>");
            Response.Write(  "</tr>");
            Response.Write(  "<tr>");
            Response.Write(      "<td>AMOUNT:</td>");
            Response.Write(      "<td>" + chargeableTotal.ToString() + "</td>");
            Response.Write(  "</tr>");
            Response.Write(  "<tr>");
            Response.Write(      "<td>Credit Card Number:</td>");
            Response.Write(      "<td>xxxx-xxxx-xxxx-" + displayCardNumber + "</td>");
            Response.Write(  "</tr>");
            Response.Write("</table>");
            Response.Write("</div>");

            //Add the UserID cookie for the SECURE URL.
            sCookieUserID.Value = Convert.ToString(buyerUserId);
            sCookieUserID.Expires = DateTime.Now.AddHours(8);

            Response.Cookies.Add(sCookieUserID);
            
            Response.Redirect("class_receipt.aspx?paymentid=" + paymentId);
        }
        else
        {
            // no cart items found
            // take them to the cart page and tell them that the session timed out and they lost their items
            //Response.Write("redirect to class_cart.aspx");
            Response.Redirect( common.getOrgFullSite( orgId ) + "/rd_classes/class_cart.aspx?message=timeout" );
        }

    }

    // ProcessPayPalTransaction() ***************************************************
    public int ProcessPayPalTransaction( PayeeInfo _PayeeInfo, int _ProcessControlNumber, string _OrgId, double _ChargeableTotal, out string _PayPalResponse )
    {
        common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "PayPal PayFLow Pro Transaction" );
        int resultCode = -1;  // default this to the communication error
        string duplicate_Transmission = "Start";

        _PayPalResponse = "";
        
        
        //      build the parameter list for PayPal
        StringBuilder parameters = new StringBuilder( );
        parameters.Append( "cardNum=" + _PayeeInfo.CardNumber );

	parameters.Append( "&paymentcontrolnumber=" + _ProcessControlNumber);
	parameters.Append( "&orgid=" + _OrgId);
	parameters.Append( "&orgfeature=classes/events");


        string expYear = common.Right( _PayeeInfo.ExpYear.ToString( ), 2 );
        //common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "exp year: " + expYear );
        parameters.Append( "&cardExp=" + _PayeeInfo.ExpMonth.ToString( "00" ) + expYear ); // MMYY format
        //common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "cardExp: " + _PayeeInfo.ExpMonth.ToString( "00" ) + expYear );
        parameters.Append( "&sjname=" + _PayeeInfo.FirstName + " " + _PayeeInfo.LastName );
        if (_PayeeInfo.RequireCVV)
            parameters.Append( "&cvv2=" + _PayeeInfo.CVV );
        parameters.Append( "&amount=" + _ChargeableTotal.ToString( "F2" ) );
        parameters.Append( "&StreetAddress=" + _PayeeInfo.Address );
        parameters.Append( "&ZipCode=" + _PayeeInfo.Zip );
        parameters.Append( "&ordernumber=" );
        //if (_OrgId == "60")
        //    common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "parameters: " + parameters.ToString( ) );
        parameters.Append( "&comment1=Recreation Purchase" );
        common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "COMMENT1: Recreation Purchase" );
        //Response.Write( "Details = " + _PayeeInfo.Details + "<br /><br />" );
        string comment2 = common.cleanAndSizeForPayFlowPro( _PayeeInfo.Details );
        parameters.Append( "&comment2=" + comment2 );
        common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "COMMENT2: " + comment2 );
        //Response.Write( "parameters = " + parameters.ToString( ) + "<br /><br />" );
        //common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Parameters: " + parameters );

        // On ocassion, PayPal will return a duplicate set of codes for an old transaction. This check for a 'Duplicate' will handle that.
        while (duplicate_Transmission != "" && duplicate_Transmission != "-1")
        {
            //      make the call to paypalsend.asp
            _PayPalResponse = PostDataToProcessor( common.getOrgFullSite( _OrgId ) + "/payment_processors/paypalsend.asp", parameters.ToString( ) );
            //Response.Write( "_PayPalResponse = [" + _PayPalResponse.ToString( ) + "]<br /><br />" );
            duplicate_Transmission = common.getPaymentProcessorResponseValue( _PayPalResponse, "DUPLICATE" );
        }
        
        // pull out the success value returned
        string resultValue;
        resultValue = common.getPaymentProcessorResponseValue( _PayPalResponse, "RESULT" );
        int.TryParse( resultValue, out resultCode );

        return resultCode;
    }


    // ProcessPointAndPayTransaction() **********************************************
    public string ProcessPointAndPayTransaction( PayeeInfo _PayeeInfo, int _ProcessControlNumber, string _OrgId, double _ChargeableTotal )
    {
        string myResponse = "";

        common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Point&Pay Transaction" );

        StringBuilder parameters = new StringBuilder( );
        parameters.Append( "paymentcontrolnumber=" + _ProcessControlNumber.ToString( ) );
        parameters.Append( "&chargeaccountnumber=" + _PayeeInfo.CardNumber );
        string expYear = common.Right( _PayeeInfo.ExpYear.ToString( ), 2 );
        parameters.Append( "&chargeexpirationmmyy=" + _PayeeInfo.ExpMonth.ToString( "00" ) + expYear );
        parameters.Append( "&signerfirstname=" + _PayeeInfo.FirstName );
        parameters.Append( "&signerlastname=" + _PayeeInfo.LastName );
        if (_PayeeInfo.RequireCVV)
            parameters.Append( "&chargecvn=" + _PayeeInfo.CVV );
        parameters.Append( "&chargeamount=" + _ChargeableTotal.ToString( "F2" ) );
        parameters.Append( "&signeraddressline1=" + _PayeeInfo.Address );
        parameters.Append( "&signeraddresscity=" + _PayeeInfo.City );
        parameters.Append( "&signeraddressregioncode=" + _PayeeInfo.State );
        parameters.Append( "&signeraddresspostalcode=" + _PayeeInfo.Zip );

        string notes = common.cleanAndSizeNotesForPointAndPay( "Recreation Purchase - " + _PayeeInfo.Details );
        parameters.Append( "&notes=" + notes );
        //Response.Write( "parameters = " + parameters.ToString( ) + "<br /><br />" );
        
        // make the call to PNP Here
        myResponse = PostDataToProcessor( common.getOrgFullSite( _OrgId ) + "/payment_processors/pnpsend.aspx", parameters.ToString( ) );
        
        return myResponse;
    }


    private string PostDataToProcessor( string url, string postData )
    {
	ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
        HttpWebRequest request = null;
        Uri uri = new Uri( url );

        request = (HttpWebRequest) WebRequest.Create( uri );
        request.Method = "POST";
        request.ContentType = "application/x-www-form-urlencoded";

        UTF8Encoding encoding = new UTF8Encoding( );
        byte[] postBytes = encoding.GetBytes( postData );
        request.ContentLength = postBytes.Length;

        Stream requestStream = request.GetRequestStream( );
        requestStream.Write( postBytes, 0, postBytes.Length );
        requestStream.Close( );

        string result = string.Empty;
        HttpWebResponse response = (HttpWebResponse) request.GetResponse( );

        if (response.StatusCode != System.Net.HttpStatusCode.OK)
        {
            throw new Exception( "HTTP status code \"" + response.StatusCode.ToString( ) + "\" returned from server" );
        }

        StreamReader reader = new StreamReader( response.GetResponseStream( ), encoding );
        result = reader.ReadToEnd( );
        reader.Close( );
        response.Close( );

        return result;
    }

    
    public string processSuccessfulTransaction(string _OrgId, string _SessionId, int _BuyerUserId, double _CartTotal, double _ChargeableTotal, double _AccountCredit, string _FeeAmount, double _CurrentAccountBalance, string _AuthCode, string _PNRef, string _RespMsg, string _OrderNumber, string _SVA, int _ProcessControlNumber)
    {
        // this handles inputting payment information into the system
        int paymentLocationId = common.getPaymentLocationId( );
        //Response.Write( "paymentLocationId = " + paymentLocationId.ToString( ) + "<br /><br />" );

        int paymentTypeId;
        int paymentAccountId = 0;
        int journalEntryTypeId = common.getJournalEntryTypeId( "purchase" );

        //Response.Write( "journalEntryTypeId = " + journalEntryTypeId.ToString( ) + "<br /><br />" );
        string ledgerId = "0";
        string paymentInformationId = "0";
        string journalNotes = "Purchase from Public Website";
        Boolean hasBuyItems = CartHasBuyItems( _SessionId );

        if (!hasBuyItems)
            journalNotes += " - All items are for the waitlist.";

        // make the Journal Entry here
        string paymentId = common.makeTheJournalEntry(_OrgId.ToString( ), paymentLocationId.ToString( ), _BuyerUserId.ToString( ), _CartTotal, journalEntryTypeId.ToString( ), journalNotes, "0", "NULL");

        common.makePaymentLogEntry(ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Journal Entry Made: " + paymentId);
        
        //Response.Write( "paymentId = " + paymentId.ToString( ) + "<br />" );

        if (_CartTotal > 0)
        {
            if (_ChargeableTotal > 0)
            {
                paymentTypeId = common.getPaymentTypeId( _OrgId, "Credit Card" );
                //Response.Write( "paymentTypeId = " + paymentTypeId.ToString( ) + "<br />" );
                paymentAccountId = common.getPaymentAccountId( _OrgId, paymentTypeId.ToString( ) );
                //Response.Write( "paymentAccountId = " + paymentAccountId.ToString( ) + "<br />" );
                
                // make the ledger entry for the Credit Card payment
                ledgerId = classes.makeALedgerEntry( paymentId, _OrgId, "debit", paymentAccountId.ToString( ), _ChargeableTotal, "NULL", "+", "NULL", "1", paymentTypeId.ToString( ), "NULL", "NULL" );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Payment Ledger Entry (" + ledgerId + ") Made for CC payment: " + _ChargeableTotal.ToString( "C" ) );
                //Response.Write( "CC ledgerId = " + ledgerId.ToString( ) + "<br />" );

                // Fill in the verisign info 
                paymentInformationId = common.saveTransactionDetails( paymentId, ledgerId, paymentTypeId.ToString( ), _ChargeableTotal, "APPROVED", "NULL", paymentAccountId.ToString( ), _AuthCode, _PNRef, _RespMsg, _OrderNumber, _SVA, _FeeAmount );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Verisign Info (" + paymentInformationId.ToString( ) + ") entered for CC payment." );
                //Response.Write( "Verisign Info entered for CC payment: " + paymentInformationId.ToString( ) + "<br />" );
            }

            if (_AccountCredit > 0)
            {
                paymentTypeId = common.getPaymentTypeId( _OrgId, "Citizen Accounts" );
                paymentAccountId = _BuyerUserId;
                //double priorBalance = common.getCitizenAccountBalance( _BuyerUserId.ToString( ) );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Citizen Account Prior Balance: " + _CurrentAccountBalance.ToString( "C" ) );

                // Make a ledger entry for the Citizen Account Credit applied
                ledgerId = classes.makeALedgerEntry( paymentId, _OrgId, "debit", paymentAccountId.ToString( ), _AccountCredit, "NULL", "-", "NULL", "1", paymentTypeId.ToString( ), _CurrentAccountBalance.ToString( "F2" ), "NULL" );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Payment Ledger Entry (" + ledgerId + ") Made for Citizen Account payment: " + _AccountCredit.ToString( "C" ) );
                //Response.Write( "Citizen Account ledgerId = " + ledgerId.ToString( ) + "<br />" );

                //AdjustCitizenAccountBalance iPurchaserId, "debit", nAccountCredit
                double newAccountBalance = _CurrentAccountBalance - _AccountCredit;
                common.setCitizenAccountBalance( _BuyerUserId.ToString( ), newAccountBalance.ToString( "F2" ) );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Setting the Citizen Account Balance to: " + newAccountBalance.ToString( "C" ) );

                // Fill in the verisign info
                paymentInformationId = common.saveTransactionDetails( paymentId, ledgerId, paymentTypeId.ToString( ), _AccountCredit, "APPROVED", "NULL", paymentAccountId.ToString( ), "NULL", "NULL", "NULL", "NULL", "NULL", "NULL" );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Verisign Info (" + paymentInformationId.ToString( ) + ") entered for Citizen Account payment." );
                //Response.Write( "Verisign Info entered for Citizen Account payment: " + paymentInformationId.ToString( ) + "<br />" );
            }
        }
        else
        {
            // handle $0 transactions here.
            // if any items are for purchase then input payment information, otherwise skip this part (those would be waitlist only transactions).
            if (hasBuyItems)
            {
                // These are handled as Credit Card charges for $0.
                paymentTypeId = common.getPaymentTypeId( _OrgId, "Credit Card" );
                //Response.Write( "$0 paymentTypeId = " + paymentTypeId.ToString( ) + "<br />" );
                paymentAccountId = common.getPaymentAccountId( _OrgId, paymentTypeId.ToString( ) );
                //Response.Write( "$0 paymentAccountId = " + paymentAccountId.ToString( ) + "<br />" );

                // make the ledger entry for the Credit Card payment
                ledgerId = classes.makeALedgerEntry( paymentId, _OrgId, "debit", paymentAccountId.ToString( ), 0, "NULL", "+", "NULL", "1", paymentTypeId.ToString( ), "NULL", "NULL" );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Payment Ledger Entry (" + ledgerId + ") Made for $0 payment: " + _ChargeableTotal.ToString( "C" ) );
                //Response.Write( "$0 ledgerId = " + ledgerId.ToString( ) + "<br />" );

                // Fill in the verisign info 
                paymentInformationId = common.saveTransactionDetails( paymentId, ledgerId, paymentTypeId.ToString( ), 0, "APPROVED", "NULL", paymentAccountId.ToString( ), "NULL", "NULL", "NULL", "NULL", "NULL", "NULL" );
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Verisign Info (" + paymentInformationId.ToString( ) + ") entered for $0 payment." );
                //Response.Write( "Verisign Info entered for $0 payment: " + paymentInformationId.ToString( ) + "<br />" );
            }
        }

        return paymentId;
        
    }
    

    public void ProcessCartItems(string _SessionId, 
                                 string _PaymentId, 
                                 int _BuyerUserId, 
                                 int _ProcessControlNumber, 
                                 string _OrgId )
    {
        int familyMemberId;
        string classListId;
        Boolean isParentClass;
        string classTypeId;
        CartItem myCartItem;
        
        // the cart struct wants to be initialized or a complier error happens when it is passed. So some values are set here.
        myCartItem.PurchaserUserId = _BuyerUserId.ToString( );
        myCartItem.CartId = "0";
        myCartItem.ClassId = "0";
        myCartItem.AttendeeUserId = "0";
        myCartItem.FamilyMemberId = "0";
        myCartItem.BuyOrWait = "B";
        myCartItem.Status = "";
        myCartItem.Quantity = "0";
        myCartItem.ClassTimeId = "0";
        myCartItem.ItemTypeId = "0";
        myCartItem.PriceTypeId = "0";
        myCartItem.Amount = 0;
        
        // Pull the cart items
        string sql = "";
        string sSessionID = "";

        sSessionID = common.dbSafe(_SessionId);
        sSessionID = "'" + sSessionID + "'";

        sql  = "SELECT cartid, ";
        sql += " classid, ";
        sql += " userid, ";
        sql += " ISNULL(familymemberid,0) AS familymemberid, ";
        sql += " quantity, ";
        sql += " amount, ";
        sql += " optionid, ";
        sql += " classtimeid, ";
        sql += " pricetypeid, ";
        sql += " buyorwait, ";
        sql += " sessionid_csharp, ";
        sql += " orgid, ";
        sql += " isparent, ";
        sql += " classtypeid, ";
        sql += " itemtypeid, ";
        sql += " ISNULL(rostergrade,'') AS rostergrade, ";
        sql += " ISNULL(rostershirtsize,'') AS rostershirtsize, ";
        sql += " ISNULL(rosterpantssize,'') AS rosterpantssize, ";
        sql += " ISNULL(rostercoachtype,'') AS rostercoachtype, ";
        sql += " ISNULL(rostervolunteercoachname,'') AS rostervolunteercoachname, ";
        sql += " ISNULL(rostervolunteercoachdayphone,'') AS rostervolunteercoachdayphone, ";
        sql += " ISNULL(rostervolunteercoachcellphone,'') AS rostervolunteercoachcellphone, ";
        sql += " ISNULL(rostervolunteercoachemail,'') AS rostervolunteercoachemail ";
        sql += " FROM egov_class_cart ";
        sql += " WHERE sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            //  loop throught the items and fill in the cart structure 
            myCartItem.CartId      = myReader["cartid"].ToString( );
            myCartItem.ClassId     = myReader["classid"].ToString( );
            myCartItem.ClassTimeId = myReader["classtimeid"].ToString( );
            myCartItem.ItemTypeId  = myReader["itemtypeid"].ToString( );
            myCartItem.Quantity    = myReader["quantity"].ToString( );
            classTypeId            = myReader["classtypeid"].ToString( );
            isParentClass          = false;
            
            Boolean.TryParse( myReader["isparent"].ToString( ), out isParentClass );

            int.TryParse( myReader["familymemberid"].ToString( ), out familyMemberId );
            
            if (familyMemberId == 0)
            {
                myCartItem.AttendeeUserId = myReader["userid"].ToString( );
                myCartItem.FamilyMemberId = "NULL";
            }
            else
            {
                myCartItem.AttendeeUserId = classes.getFamilyMemberUserId( familyMemberId.ToString( ) );
                myCartItem.FamilyMemberId = familyMemberId.ToString( );
            }

            myCartItem.BuyOrWait = myReader["buyorwait"].ToString( );
            if (myCartItem.BuyOrWait == "B")
            {
                myCartItem.Status = "ACTIVE";
                myCartItem.Amount = 0;
                double.TryParse( myReader["amount"].ToString( ), out myCartItem.Amount );
            }
            else
            {
                myCartItem.Status = "WAITLIST";
                myCartItem.Amount = 0;
            }

            if (myReader["rostergrade"].ToString( ) == "")
                myCartItem.RosterGrade = "NULL";
            else
                myCartItem.RosterGrade = "'" + common.dbready_string( myReader["rostergrade"].ToString( ), 2 ) + "'";

            if (myReader["rostershirtsize"].ToString( ) == "")
                myCartItem.RosterShirtSize = "NULL";
            else
                myCartItem.RosterShirtSize = "'" + common.dbready_string( myReader["rostershirtsize"].ToString( ), 50 ) + "'";

            if (myReader["rosterpantssize"].ToString( ) == "")
                myCartItem.RosterPantsSize = "NULL";
            else
                myCartItem.RosterPantsSize = "'" + common.dbready_string( myReader["rosterpantssize"].ToString( ), 50 ) + "'";

            if (myReader["rostercoachtype"].ToString( ) == "")
                myCartItem.RosterCoachType = "NULL";
            else
                myCartItem.RosterCoachType = "'" + common.dbready_string( myReader["rostercoachtype"].ToString( ), 50 ) + "'";

            if (myReader["rostervolunteercoachname"].ToString( ) == "")
                myCartItem.RosterVolunteerCoachName = "NULL";
            else
                myCartItem.RosterVolunteerCoachName = "'" + common.dbready_string( myReader["rostervolunteercoachname"].ToString( ), 100 ) + "'";

            if (myReader["rostervolunteercoachdayphone"].ToString( ) == "")
                myCartItem.RosterVolunteerCoachDayPhone = "NULL";
            else
                myCartItem.RosterVolunteerCoachDayPhone = "'" + common.dbready_string( myReader["rostervolunteercoachdayphone"].ToString( ), 10 ) + "'";

            if (myReader["rostervolunteercoachcellphone"].ToString( ) == "")
                myCartItem.RosterVolunteerCoachCellPhone = "NULL";
            else
                myCartItem.RosterVolunteerCoachCellPhone = "'" + common.dbready_string( myReader["rostervolunteercoachcellphone"].ToString( ), 10 ) + "'";

            if (myReader["rostervolunteercoachemail"].ToString( ) == "")
                myCartItem.RosterVolunteerCoachEmail = "NULL";
            else
                myCartItem.RosterVolunteerCoachEmail = "'" + common.dbready_string( myReader["rostervolunteercoachemail"].ToString( ), 100 ) + "'";

            // add to the class list for the item
            classListId = addToClassList( _PaymentId, myCartItem );
            common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Added classlistid: " + classListId );
            //Response.Write( "Add to Class List completed: " + classListId + "<br />" );
            
            // CreateJournalItemStatus
            createJournalItemStatus( _PaymentId, myCartItem.ItemTypeId, classListId, myCartItem.Status, myCartItem.BuyOrWait );
            //Response.Write( "createJournalItemStatus completed<br />" );

            // AddClassLedgerRows
            classes.addClassLedgerEntries( _OrgId, _PaymentId, myCartItem.CartId, myCartItem.ItemTypeId, classListId, "credit", "+" );
            //Response.Write( "addClassLedgerEntries completed<br />" );

            // if class is a series parent then add to child classes
            if (isParentClass && classTypeId == "1")
            {
                common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Series Children Being Added" );
                AddToChildClassLists( _PaymentId, myCartItem, myCartItem.ClassId, _ProcessControlNumber, _OrgId );
            }

        }   //  End cart processing loop

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );
        
        common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Cart Processing finished." );
    }

    public string addToClassList( string _PaymentId, CartItem _CartItem )
    {
        string classListId = "0";

        string sql = "INSERT INTO egov_class_list ( userid, classid, status, quantity, classtimeid, familymemberid, amount, paymentid, ";
        sql += "attendeeuserid, rostergrade, rostershirtsize, rosterpantssize, rostercoachtype, rostervolunteercoachname, ";
        sql += "rostervolunteercoachdayphone, rostervolunteercoachcellphone, rostervolunteercoachemail ) VALUES ( ";
        sql += _CartItem.PurchaserUserId + ", " + _CartItem.ClassId + ", '" + _CartItem.Status + "', " + _CartItem.Quantity + ", ";
        sql += _CartItem.ClassTimeId + ", " + _CartItem.FamilyMemberId + ", " + _CartItem.Amount.ToString( "F2" ) + ", " + _PaymentId + ", ";
        sql += _CartItem.AttendeeUserId + ", " + _CartItem.RosterGrade + ", " + _CartItem.RosterShirtSize + ", " + _CartItem.RosterPantsSize + ", ";
        sql += _CartItem.RosterCoachType + ", " + _CartItem.RosterVolunteerCoachName + ", " + _CartItem.RosterVolunteerCoachDayPhone + ", ";
        sql += _CartItem.RosterVolunteerCoachCellPhone + ", " + _CartItem.RosterVolunteerCoachEmail + " )";

        classListId = common.RunInsertStatement( sql );

        return classListId;
    }

    public void createJournalItemStatus( string _PaymentId, string _ItemTypeId, string _ClassListId, string _Status, string _BuyOrWait)
    {
        string sql = "INSERT INTO egov_journal_item_status ( paymentid, itemtypeid, itemid, status, buyorwait ) Values ( ";
        sql += _PaymentId + ", " + _ItemTypeId + ", " + _ClassListId + ", '" + _Status + "', '" + _BuyOrWait + "')";

        common.RunSQLStatement( sql );
    }

    public void AddToChildClassLists( string _PaymentId, CartItem _CartItem, string _ParentClassId, int _ProcessControlNumber, string _OrgId )
    {
        string childClassListId;
        string sql = "SELECT C.classid, T.timeid FROM egov_class C, egov_class_time T ";
        sql += "WHERE C.classid = T.classid AND C.parentclassid = " + _ParentClassId;
        
        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            _CartItem.ClassId = myReader["classid"].ToString( );
            _CartItem.ClassTimeId = myReader["timeid"].ToString( );

            childClassListId = addToClassList( _PaymentId, _CartItem );
            common.makePaymentLogEntry( ref _ProcessControlNumber, _OrgId, "public", "classes/events", "Added child class. classlistid: " + childClassListId );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );
    }

    public void clearTheCart( string _SessionId )
    {
        // get the items in the cart
        string sql = "";
        string sSessionID = "";

        sSessionID = common.dbSafe(_SessionId);
        sSessionID = "'" + sSessionID + "'";

        sql  = "SELECT cartid ";
        sql += " FROM egov_class_cart ";
        sql += " WHERE sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            // delete the pricing information for each item
            sql = "DELETE FROM egov_class_cart_price WHERE cartid = " + myReader["cartid"].ToString( );
            common.RunSQLStatement( sql );

            // then delete the item from the cart
            sql = "DELETE FROM egov_class_cart WHERE cartid = " + myReader["cartid"].ToString( );
            common.RunSQLStatement( sql );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );
    }

    public void sendEmailNotifications(string _SessionId,
                                       string _BuyerUserId,
                                       string _OrgId,
                                       string _PaymentId,
                                       double _CartTotal,
                                       double _ChargeTotal,
                                       double _CitizenAccountAmount,
                                       string _AuthCode,
                                       string _PNREF,
                                       PayeeInfo _PayeeInfo,
                                       string _OrderNumber,
                                       string _SVA, 
                                       double _FeeAmount)
    {
        int orgId;
        int.TryParse( _OrgId, out orgId );
        string orgName = common.getOrgName( _OrgId );
        string adminEmailAddr;
        string buyerEmailAddr;
        string subject = "Thank you for your " + orgName + " E-Gov Purchase";
        string emailMessage;

        //Response.Write( "In sendEmailNotifications<br />" );
        // Get the email address of the admin assigned to receive these emails
        adminEmailAddr = classes.getAssignedAdminEmail( _OrgId );
        //Response.Write( "adminEmailAddr: " + adminEmailAddr + "<br />" );
        
        // Get the email address of the buyer
        buyerEmailAddr = common.getCitizenEmailAddress( _BuyerUserId );
        //Response.Write( "buyerEmailAddr: " + buyerEmailAddr + "<br />" );
        
        // Build the email message to the buyer
        emailMessage = buildEmailMessage( _SessionId, _OrgId, adminEmailAddr, "BUYER", _PaymentId, _CartTotal, _ChargeTotal, _CitizenAccountAmount, _FeeAmount, _PayeeInfo, _AuthCode, _PNREF, _OrderNumber, _SVA );
        //Response.Write( "emailMessage: " + emailMessage + "<br />" );

        // Send the BUYER email
        //Response.Write( "Sending Buyer email<br />" );
        common.sendMessage( orgId, "", "", buyerEmailAddr, subject, emailMessage, "normal", true, "", "" );

        subject = "Purchase: " + orgName + " (re: Items purchased from E-Gov website)";
        // Build the email message to the city
        emailMessage = buildEmailMessage( _SessionId, _OrgId, adminEmailAddr, "ADMIN", _PaymentId, _CartTotal, _ChargeTotal, _CitizenAccountAmount, _FeeAmount, _PayeeInfo, _AuthCode, _PNREF, _OrderNumber, _SVA );
        //Response.Write( "emailMessage: " + emailMessage + "<br />" );

        // Send the ADMIN email
        //Response.Write( "Sending Admin email<br />" );
        common.sendMessage( orgId, "", "", adminEmailAddr, subject, emailMessage, "normal", true, "", "" );
    }
    
    
    public string buildEmailMessage( string _SessionId, string _OrgId, string _AdminEmailAddr, string _MessageType, string _PaymentId, double _CartTotal, double _ChargeTotal, double _CitizenAccountAmount, double _FeeAmount, PayeeInfo _PayeeInfo, string _AuthCode, string _PNREF, string _OrderNumber, string _SVA )
    {
        StringBuilder emailBody = new StringBuilder( );

        emailBody.Append( "<p>This automated message was sent by the " + common.getOrgName( _OrgId ) + " E-Gov web site. Do not reply to this message.  " );
        emailBody.Append( "Contact " + _AdminEmailAddr + " for inquiries regarding this email.</p>" );

        if (_MessageType.ToUpper( ) == "BUYER")
            emailBody.Append( "<p>Thank you for making a purchase on " + DateTime.Today.ToString( "d" ) + ".</p>" );
        else
            emailBody.Append( "<p>Purchase was made on " + DateTime.Today.ToString( "d" ) + ".</p>" );

        emailBody.Append( "<p><strong>Recreation Purchase</strong></p>" );
        emailBody.Append( "<p><strong>Purchased Items:</strong><br />" );
        emailBody.Append( getCartItemValuesForEmail( _SessionId ) + "</p>" );

        if (_MessageType.ToUpper( ) == "BUYER")
        {
            emailBody.Append( "<p><strong>Link to view these Details:</strong><br />" );
            string virtualDirectory = "/" + common.getOrgInfo( _OrgId, "OrgVirtualSiteName" );
            //emailBody.Append( "<a href=\"http://www.egovlink.com" + virtualDirectory + "/classes/view_receipt.asp?ipaymentid=" + _PaymentId + "\">http://www.egovlink.com" + virtualDirectory + "/classes/view_receipt.asp?ipaymentid=" + _PaymentId + "</a></p>" );
            emailBody.Append("<a href=\"http://dev4.egovlink.com" + virtualDirectory + "/rd_classes/class_receipt.aspx?paymentid=" + _PaymentId + "\">http://dev4.egovlink.com" + virtualDirectory + "/rd_classes/class_receipt.aspx?paymentid=" + _PaymentId + "</a></p>");
        }
        
        // transaction information
        emailBody.Append( "<p><strong>Transaction Details</strong><br />" );
        emailBody.Append( "Item Total: " + _CartTotal.ToString( "C" ) + "<br />" );
        if (_CitizenAccountAmount > 0)
            emailBody.Append( "Account Credit Applied: " + _CitizenAccountAmount.ToString( "C" ) + "<br />" );
        if (_ChargeTotal > 0)
        {
            if (_FeeAmount > 0)
            {
                emailBody.Append( "Processing Fee: " + _FeeAmount.ToString( "C" ) + "<br />" );
                _ChargeTotal += _FeeAmount;
            }
            emailBody.Append( "Amount Charged: " + _ChargeTotal.ToString( "C" ) + "<br />" );
            if (_PNREF != "")
                emailBody.Append( "Payment Reference Number: " + _PNREF + "<br />" );
            if (_AuthCode != "")
                emailBody.Append( "Authorization Code: " + _AuthCode + "<br />" );
            if (_OrderNumber != "")
                emailBody.Append( "Order Number: " + _OrderNumber + "<br />" );
            if (_SVA != "")
                emailBody.Append( "SVA: " + _SVA + "<br />" );
        }
        emailBody.Append( "</p>" );
        
        // purchaser information
        emailBody.Append( "<p><strong>User Information</strong><br />" );
        if (_PayeeInfo.CardNumber != "")
            emailBody.Append( "Credit Card: XXXXXXXXXXXX" + common.Right( _PayeeInfo.CardNumber, 4 ) + "<br />" );
        emailBody.Append( "Name: " + _PayeeInfo.FirstName + " " + _PayeeInfo.LastName + "<br />" );
        emailBody.Append( "Address: " + _PayeeInfo.Address + "<br />" );
        emailBody.Append( "City: " + _PayeeInfo.City + "<br />" );
        emailBody.Append( "State: " + _PayeeInfo.State + "<br />" );
        emailBody.Append( "Zip: " + _PayeeInfo.Zip + "<br />" );
        emailBody.Append( "</p>" );
        
        // now show team roster info for Craig CO
        if (common.orgHasFeature( _OrgId, "custom_registration_CraigCO" ))
        {
            emailBody.Append( getTeamRegistrationDetails( _SessionId, _OrgId ) );
        }
        //Response.Write( "emailBody: " + emailBody.ToString( ) + "<br />" );
        
        return emailBody.ToString( );
    }


    public string getTeamRegistrationDetails( string _SessionId, string _OrgId )
    {
        StringBuilder teamDetails = new StringBuilder( );
        string shirtLabel = "T-Shirt";

        teamDetails.Append( "" );

        string sql = "";
        string sSessionID = "";

        sSessionID = common.dbSafe(_SessionId);
        sSessionID = "'" + sSessionID + "'";

        sql = "SELECT C.classname, ";
        sql += " U.userfname, ";
        sql += " U.userlname, ";
        sql += " CC.rostergrade, ";
        sql += " CC.rostershirtsize, ";
        sql += " CC.rosterpantssize, ";
        sql += " ISNULL(CC.rostercoachtype,'') AS rostercoachtype, ";
        sql += " ISNULL(CC.rostervolunteercoachname,'') AS rostervolunteercoachname, ";
        sql += " ISNULL(CC.rostervolunteercoachdayphone,'') AS rostervolunteercoachdayphone, ";
        sql += " ISNULL(CC.rostervolunteercoachcellphone,'') AS rostervolunteercoachcellphone, ";
        sql += " ISNULL(CC.rostervolunteercoachemail,'') AS rostervolunteercoachemail ";
        sql += " FROM egov_class_cart CC, ";
        sql +=      " egov_class C, ";
        sql +=      " egov_familymembers F, ";
        sql +=      " egov_users U ";
        sql += " WHERE CC.classid = C.classid ";
        sql += " AND CC.familymemberid = F.familymemberid ";
        sql += " AND F.userid = U.userid ";
        sql += " AND CC.rostergrade IS NOT NULL ";
        sql += " AND CC.sessionid_csharp = " + sSessionID;
        sql += " ORDER BY CC.dateadded";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            teamDetails.Append( "<p><strong>Team Registration Information</strong><br />" );
            while (myReader.Read( ))
            {
                teamDetails.Append( "Team: " + myReader["classname"].ToString( ) + "<br />" );
                teamDetails.Append( "Name: " + myReader["userfname"].ToString( ) + " " + myReader["userlname"].ToString( ) + "<br />" );
                teamDetails.Append( "Grade: " + myReader["rostergrade"].ToString( ) + "<br />" );

                if (common.orgHasDisplay( _OrgId, "class_teamregistration_tshirt_label" ))
                {
                    shirtLabel = common.getOrgDisplay( _OrgId, "class_teamregistration_tshirt_label" );
                    shirtLabel = (shirtLabel == "") ? "T-Shirt" : shirtLabel;
                }
                teamDetails.Append( shirtLabel + " Size: " + myReader["rostershirtsize"].ToString( ) + "<br />" );

                teamDetails.Append( "Pants Size: " + myReader["rosterpantssize"].ToString( ) + "<br />" );

                if (myReader["rostercoachtype"].ToString( ) != "")
                {
                    teamDetails.Append( "Knows someone or would like to be a volunteer: " + myReader["rostercoachtype"].ToString( ) + "<br />" );
                    teamDetails.Append( "Coach Name: " + myReader["rostervolunteercoachname"].ToString( ) + "<br />" );
                    if (myReader["rostervolunteercoachdayphone"].ToString( ) != "")
                        teamDetails.Append( "Day Phone:  " + myReader["rostervolunteercoachdayphone"].ToString( ) + "<br />" );
                    if (myReader["rostervolunteercoachcellphone"].ToString( ) != "")
                        teamDetails.Append( "Cell Phone: " + myReader["rostervolunteercoachcellphone"].ToString( ) + "<br />" );
                    if (myReader["rostervolunteercoachemail"].ToString( ) != "")
                        teamDetails.Append( "Email: " + myReader["rostervolunteercoachemail"].ToString( ) + "<br />" );
                }
                teamDetails.Append( "<br />" );
            }
            teamDetails.Append( "</p>" );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return teamDetails.ToString( );
    }

    public string getCartItemValuesForEmail( string _SessionId )
    {
        DateTime startDate;

        // build a string of things in the cart.
        StringBuilder cartItemNames = new StringBuilder( );
        
        string familyMemberUserId;
        string familyMemberName;
        string sSessionID = "";
        string sql = "";

        sSessionID = common.dbSafe(_SessionId);
        sSessionID = "'" + sSessionID + "'";

        sql = "SELECT CC.cartid, ";
        sql += " C.classname, ";
        sql += " C.startdate, ";
        sql += " CC.quantity, ";
        sql += " CC.buyorwait, ";
        sql += " ISNULL(CC.optionid,0) AS optionid, ";
        sql += " ISNULL(CC.familymemberid,0) as familymemberid, ";
        sql += " CC.itemtypeid ";
        sql += " FROM egov_class_cart CC, ";
        sql += " egov_class C ";
        sql += " WHERE CC.classid = C.classid ";
        sql += " AND CC.sessionid_csharp = " + sSessionID;
        sql += " ORDER BY CC.dateadded";
        
        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            if (cartItemNames.Length > 0)
                cartItemNames.Append( "<br />" );

            familyMemberUserId = myReader["familymemberid"].ToString( );
            familyMemberName = common.getFamilyMemberName( familyMemberUserId );
            cartItemNames.Append( familyMemberName + " was added to the " );

            if (myReader["buyorwait"].ToString( ).ToUpper( ) == "W")
                cartItemNames.Append( "wait list of " );
            else
                cartItemNames.Append( "list for " );

            DateTime.TryParse( myReader["startdate"].ToString( ), out startDate );

            cartItemNames.Append( myReader["classname"].ToString( ) + ". Start Date: " + startDate.ToString( "d" ) );

            // if this is a ticketed event, then we need the quantity of tickets purchased
            if (myReader["optionid"].ToString( ) == "2")
                cartItemNames.Append( " Qty: " + myReader["quantity"].ToString( ) );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return cartItemNames.ToString( );
    }


    public void getPayeeInfo( int _BuyerUserId, ref PayeeInfo _PayeeInfo )
    {
        string sql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, ";
        sql += "ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, ISNULL(useremail,'') AS useremail ";
        sql += "FROM egov_users WHERE userid = " + _BuyerUserId.ToString( );
        
        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            // set the stucture values here
            _PayeeInfo.FirstName = myReader["userfname"].ToString( );
            _PayeeInfo.LastName = myReader["userlname"].ToString( );
            _PayeeInfo.Address = myReader["useraddress"].ToString( );
            _PayeeInfo.City = myReader["usercity"].ToString( );
            _PayeeInfo.State = myReader["userstate"].ToString( );
            _PayeeInfo.Zip = myReader["userzip"].ToString( );
            _PayeeInfo.Email = myReader["useremail"].ToString( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );
    }

    public Boolean CartHasBuyItems( string _SessionId )
    {
        Boolean hasBuyItems = false;

        // if there are any buys, they will be at the top, and we only need one
        string sql = "";
        string sSessionID = "";

        sSessionID = common.dbSafe(_SessionId);
        sSessionID = "'" + sSessionID + "'";

        sql  = "SELECT TOP(1) buyorwait ";
        sql += " FROM egov_class_cart ";
        sql += " WHERE sessionid_csharp = " + sSessionID;
        sql += " ORDER BY buyorwait ASC";
        
        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            
            if (myReader["buyorwait"].ToString( ) == "B")
                hasBuyItems = true;
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return hasBuyItems;
    }


    public string getCartDetails( string _SessionId )
    {
        // the details are the Activity No of the items being purchased
        StringBuilder cartDetails = new StringBuilder( );

        cartDetails.Append( "" );

        string sql = "";
        string sSessionID = "";

        sSessionID = common.dbSafe(_SessionId);
        sSessionID = "'" + sSessionID + "'";
        
        sql  = "SELECT ISNULL(T.activityno,'') AS activityno ";
        sql += " FROM egov_class_cart C, ";
        sql += " egov_class_time T ";
        sql += " WHERE C.classtimeid = T.timeid ";
        sql += " AND C.sessionid_csharp = " + sSessionID;
        return sql;
/*
        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            if (cartDetails.Length > 0)
                cartDetails.Append( "," );
            
            cartDetails.Append( myReader["activityno"].ToString( ) );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return cartDetails.ToString( );
 */
    }


    public void sendLoginFailedEmail( string _OrgId )
    {
        int orgId;
        int.TryParse( _OrgId, out orgId );
        string orgName = common.getOrgName( _OrgId );
        string subject = "PayPal User Authentication Error Received";
        string body;

        body = "<p>A PayPal User Authentication error has been received for " + orgName + ". ";
        body += "<br />Someone has changed the PayPal account password without updating it in the E-Gov system.</p>";

        common.sendMessage( orgId, "noreply@eclink.com", "", "egovsupport@eclink.com", subject, body, "high", true, "", "" );

        if (orgId == 26)
        {
            common.sendMessage( orgId, "noreply@eclink.com", "", "mvander@ci.montgomery.oh.us", subject, body, "high", true, "", "" );
            common.sendMessage( orgId, "noreply@eclink.com", "", "edumont@ci.montgomery.oh.us", subject, body, "high", true, "", "" );
            common.sendMessage( orgId, "noreply@eclink.com", "", "cbridgewater@ci.montgomery.oh.us", subject, body, "high", true, "", "" );
        }

    }

</script>

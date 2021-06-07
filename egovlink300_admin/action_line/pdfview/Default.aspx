<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="PDFFormDemo._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>
        Fill Out this Form!
    </h2>
   
    <div class="accountInfo">
                        <fieldset class="register">
                            <legend>Account Information</legend>
                            <p>
                                <asp:Label ID="UserNameLabel" runat="server" >Full Name:</asp:Label>
                                <asp:TextBox ID="fullname" runat="server" CssClass="textEntry" Width="171px"></asp:TextBox>
                              
                            </p>
                            <p>
                                <asp:CheckBox ID="chkIsUpdate" runat="server" 
                                    Text="Is this an Update to a previous form?" TextAlign="Left" />
                                
                            </p>
                            <p>
                                <asp:Label ID="PasswordLabel" runat="server" >Name of Local Gov&#39;t officer:</asp:Label>
                                <asp:TextBox ID="officer" runat="server" CssClass="passwordEntry"></asp:TextBox>
                               
                            </p>
                            <p>
                              
                               
                                <asp:CheckBox ID="chkIncome" runat="server" Text="Does Officer receive Income?" 
                                    TextAlign="Left" />
                            </p>

                               <p>
                              
                               
                                <asp:CheckBox ID="chkYouIncome" runat="server" Text="Will you receive income?" 
                                       TextAlign="Left" />
                            </p>

                              <p>
                              
                               
                                <asp:CheckBox ID="chkCorpRelation" runat="server" Text="Does Govt Officer have any corporate relation to you?" />
                            </p>

                             <p>
                                <asp:Label ID="Label3" runat="server" >Describe any relation:</asp:Label>
                                <asp:TextBox ID="txtRelation" runat="server" CssClass="textEntry" 
                                     TextMode="MultiLine" Width="360px"></asp:TextBox>
                              
                            </p>
                              



                        </fieldset>
                        <p class="submitButton">
                            <asp:Button ID="btnPrint" runat="server" Text="PRINT FORM" 
                                onclick="btnPrint_Click"/>
                        </p>
                    </div>
</asp:Content>

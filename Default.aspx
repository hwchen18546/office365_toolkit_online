<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" EnableViewState="true" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" xmlns:mso="urn:schemas-microsoft-com:office:office" xmlns:msdt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <link rel="stylesheet" type="text/css" href="css/365Style.css"/>
    <title>Office365 Admin Web App</title>

<!--[if gte mso 9]><xml>
<mso:CustomDocumentProperties>
<mso:IsMyDocuments msdt:dt="string">1</mso:IsMyDocuments>
</mso:CustomDocumentProperties>
</xml><![endif]-->
</head>
<body>
    <form id="form1" runat="server">
            <div id="div_login">
                        <br/>
                         <img src="img/logo.png" height="86" width="300"/><br/>
                        <b><span style="font-size:20px;">Administrator Toolkit</span></b><br/><br/>
                        <b>Account<br/> </b>
                        <asp:TextBox ID="Account" runat="server" Width="200px"></asp:TextBox><br/>
                        <b>Password<br/> </b>
                        <asp:TextBox ID="Password" TextMode ="Password" runat="server" Width="201px"></asp:TextBox><br/>
                        <asp:Button ID="But_login" 
                            runat="server"
                            OnClick="But_login_Click"  
                            Text="Log in" Font-Size="Larger" Height="35" ForeColor="Blue" BackColor="#E7E7E7" BorderWidth="0" />
                        <asp:Button ID="But_out" 
                            runat="server"
                            OnClick="But_logout_Click"  
                            Text="Log out"  Font-Size="Larger" Height="35" ForeColor="Blue" BackColor="#E7E7E7" BorderWidth="0"/><br/>
            </div>
            <hr/>
            <div id="div_fun_btn"  align=center>
                        <asp:Button ID="SetUsr" 
                            runat="server"
                            Visible = "false"
                            Enabled = "false"
                            OnClick="SetUsr_Click"  
                            Text="Set-O365user"
                            CssClass="btn_fun"   />
                        <asp:Button ID="AddUsr" 
                            runat="server"
                            Visible = "false"
                            Enabled = "false"
                            OnClick="AddUsr_Click"  
                            Text="Add-O365user"
                            CssClass="btn_fun"   />     
                        <asp:Button ID="AddMultUsr" 
                            runat="server"
                            Visible = "false"
                            Enabled = "false"
                            OnClick="AddMultUsr_Click"  
                            Text="Add-MultiO365user"
                            CssClass="btn_fun"   />                               
                        <asp:Button ID="CheckLog" 
                            runat="server"
                            Visible = "false"
                            Enabled = "false"
                            OnClick="CheckLog_Click"  
                            Text="Check-Log"
                            CssClass="btn_fun"   />                                    
            <br/>
            </div>
            <hr/>
            <div id="div_set" runat="server"  class="DivGroup">
                    <!-- Get-Msoluser --> 
                    <asp:GridView ID="GridView1" runat="server" CssClass="GV_DataTable" AllowSorting="True" OnSorting="GridView1_Sorting" AutoGenerateColumns="False">
                        <Columns>
                            <asp:BoundField DataField="UserPrincipalName" HeaderText="UserPrincipalName" SortExpression="UserPrincipalName" />
                            <asp:BoundField DataField="DisplayName" HeaderText="DisplayName" SortExpression="DisplayName" />
                            <asp:BoundField DataField="FirstName" HeaderText="FirstName" SortExpression="FirstName" />
                            <asp:BoundField DataField="LastName" HeaderText="LastName" SortExpression="LastName" />
                            <asp:BoundField DataField="Department" HeaderText="Department" SortExpression="Department" />
                            <asp:BoundField DataField="UsageLocation" HeaderText="UsageLocation" SortExpression="UsageLocation" />
                            <asp:BoundField DataField="IsLicensed" HeaderText="IsLicensed" SortExpression="IsLicensed" />
                            <asp:TemplateField HeaderText="Function">
                                <ItemTemplate>
                                    <asp:Button ID="Btn_Detail" runat="server" Text="Detail Info" OnClick="Btn_Detail_Click" />
                                    <asp:Button ID="Btn_SetPwd" runat="server" Text="Set Password" OnClick="Btn_SetPwd_Click" />
                                    <asp:Button ID="Btn_Remove" runat="server" Text="Remove User" OnClick="Btn_Remove_Click" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
            </div>
            <div id="div_detail" runat="server"  class="DivGroup">
                
            </div>
            <div id="div_SetPwd" runat="server"  class="DivGroup">
                     <h1>Change Password</h1>
                    <table style="width:300px">
                        <tr><td>UserPrincipalName</td></tr>
                        <tr><td><asp:TextBox ID="UPN" runat="server" Width="200px"></asp:TextBox></td></tr>    
                        <tr><td>New Password</td></tr>
                        <tr><td><asp:TextBox ID="NewPwd" runat="server" Width="200px"></asp:TextBox></td></tr>    
                    </table>
                        <asp:Button ID="Btn_Pwd" runat="server" Text="Set" OnClick="Do_Pwd_Click" />                                                                            
            </div>
            <div id="div_Remove" runat="server"  class="DivGroup">  
                <asp:Label ID="Lab_Warn" runat="server" Text="Label"></asp:Label>  
                <asp:Label ID="Lab_Account" runat="server" Text="Label"></asp:Label>         
                <asp:Button ID="Btn_Yes" runat="server" Text="Yes" OnClick="Do_Remove_Click" />
                <asp:Button ID="Btn_No" runat="server" Text="No" OnClick="SetUsr_Click" />                  
            </div>
            <div id="div_add" runat="server"  class="DivGroup"  align=center>
                     <h2 style="color: blue">Create an account</h2>
                    <table style="width:300px">
                        <tr><td><b>1. UserPrincipalName</b></td></tr>
                        <tr><td><asp:TextBox ID="Add_upn" runat="server" Width="200px"></asp:TextBox></td></tr>
                        <tr><td><b>2.DisplayName</b></td></tr>
                        <tr><td><asp:TextBox ID="Add_display" runat="server" Width="200px"></asp:TextBox></td></tr>
                        <tr><td><b>3.FirstName</b></td></tr>
                        <tr><td><asp:TextBox ID="Add_first" runat="server" Width="200px"></asp:TextBox></td></tr>
                        <tr><td><b>4. LastName</b></td></tr>
                        <tr><td><asp:TextBox ID="Add_last" runat="server" Width="200px"></asp:TextBox></td></tr>
                        <tr><td><b>5. Password</b></td></tr>
                        <tr><td><asp:TextBox ID="Add_password" runat="server" Width="200px"></asp:TextBox></td></tr>
                        <tr><td><b>6. Select License</b></td></tr>
                        <tr><td>
                            <asp:DropDownList ID="dl" runat="server"></asp:DropDownList><br/>
                        </td></tr>
                    </table>                   
                    <asp:GridView ID="GridView2" runat="server" CssClass="GV_DataTable" AutoGenerateColumns="True"></asp:GridView>
                    <asp:Button ID="Btn_add" runat="server" Text="Add User" OnClick="Btn_add_Click" />
            </div>
            <div id="div_log" runat="server"  class="GV_DataTable">               
            </div>
            <div id="div_addmult" runat="server"  align=center  class="DivGroup">
                    <h2 style="color: blue">Creat Multi Office365 Users</h2>
                       <table style="width:300px">
                        <tr><td><h4>1. Download sample file</h4> </td></tr>
                        <tr><td><asp:Button
                            ID="btn_download"
                            Text="Download Sample file"
                            runat="server"
                            OnClick="btn_download_Click" BorderStyle="NotSet" />  </td></tr>
                        <tr><td><h4>2. Import users file</h4> </td></tr>
                        <tr><td><asp:FileUpload ID="fileupload1" runat="server" /> </td></tr>
                        <tr><td><br/><b>3. Select License</b></td></tr>
                        <tr><td>
                            <asp:DropDownList ID="dl2" runat="server"></asp:DropDownList><br/>
                        </td></tr>
                   </table><br/>
                   <asp:GridView ID="GridView3" runat="server" CssClass="GV_DataTable" AutoGenerateColumns="True"></asp:GridView>
                    <asp:Label
                        ID="Lab_up"
                        runat="server"
                        Font-Bold="True"
                        ForeColor="#000099"></asp:Label> <br/>
                    <asp:Button
                        ID="btn_upload"
                        Text="Add Users"
                        runat="server"
                        OnClick="btn_upload_Click" /> 
                    </div>
    </form>
                
</body>
</html>

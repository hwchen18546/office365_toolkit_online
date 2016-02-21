using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


/* StringBuilder */
using System.Text;
/* PSObject */
/* Add reference C:\Program Files (x86)\Reference Assemblies\Microsoft\WindowsPowerShell\3.0\System.Management.Automation.dll */
using System.Management.Automation;
/* Collection<PSObject> */
using System.Collections.ObjectModel;
/* DataTable */
using System.Data;
/* SqlConnection */
using System.Data.SqlClient;
/* Collection */
using System.Collections;
/* Upload Path */
using System.IO;

public partial class _Default : System.Web.UI.Page
{
/********* global var *********/
    StringBuilder ss = new StringBuilder();
    PowerShell ps = PowerShell.Create();
    Collection<PSObject> results;
    PSDataCollection<ErrorRecord> psErr;
    DataTable userList;

/********* Init *********/
    /* Page_Load */
    protected void Page_Load(object sender, EventArgs e)
    {
        div_add.Visible = false;
        div_addmult.Visible = false;
        div_set.Visible = false;
        div_log.Visible = false;
        div_detail.Visible = false;
        div_SetPwd.Visible = false;
        div_Remove.Visible = false;
    }

/********* Login *********/
    /* Login Button Click*/
    protected void But_login_Click(Object sender, EventArgs e)
    {
        ss.AppendLine("$password = ConvertTo-SecureString \"" + Password.Text + "\" -AsPlainText –Force");
        ss.AppendLine("$credential = New-Object  System.Management.Automation.PsCredential(\""
            + Account.Text + "\",$password)");
        ss.AppendLine("$cred = Get-Credential -cred $credential");
        ss.AppendLine("Import-Module MsOnline");
        ss.AppendLine("Connect-Msolservice -cred $cred");

        ps.AddScript(ss.ToString());
        // execute the script
        results = ps.Invoke();
        //Catch the exception when run PS script 
        psErr = ps.Streams.Error;
        if ((psErr != null) && (psErr.Count > 0))
        {
            Response.Write("We don't recognize this user ID or password");
        }
        else
        {
            /* Add into SQL*/
            SqlConnection Conn = new SqlConnection("server =localhost\\SQLEXPRESS; uid = vincechen; pwd = p@ssword; database=o365");
            Conn.Open();
            DateTime now = DateTime.Now;
            string queryString = "insert into o365_log(upn,login,ip) values(N' " + Account.Text + "',N'" + now.ToString() + "',N'" + GetIPAddress() + "' )";
            SqlCommand cmd = new SqlCommand(queryString, Conn);
            SqlDataReader dr = cmd.ExecuteReader();
            cmd.Cancel();
            dr.Close();
            Conn.Close();
            //
            /* Check Log  */
            CheckLog.Visible = false;
            CheckLog.Enabled = false;
            if (Account.Text == "vince.chen@5566.idv.tw")
            {
                CheckLog.Visible = true;
                CheckLog.Enabled = true;
            }
            //
            SetUsr.Visible = true;
            SetUsr.Enabled = true;
            AddUsr.Visible = true;
            AddUsr.Enabled = true;
            AddMultUsr.Visible = true;
            AddMultUsr.Enabled = true;
            Response.Write("Login success. Direct to dashboard...");
        }
    }

    protected void But_logout_Click(Object sender, EventArgs e)
    {
        SetUsr.Visible = false;
        SetUsr.Enabled = false;
        AddUsr.Visible = false;
        AddUsr.Enabled = false;
        AddMultUsr.Visible = false;
        AddMultUsr.Enabled = false;
        CheckLog.Visible = false;
        CheckLog.Enabled = false;
    }

/********* SetUsr *********/
    /* Admin Function Button Click - SetUsr */
    protected void SetUsr_Click(Object sender, EventArgs e)
    {
        GetMsoluser("UserPrincipalName", "Asc");
    }

    /* SetUsr中的欄位排列 button click */
    protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
    {

        if (e.SortDirection == SortDirection.Ascending)
        {
            // Asc query for e.SortExpression field;
            GetMsoluser(e.SortExpression, "Asc");
        }
        else
        {
            // Desc query for e.SortExpression field;
            GetMsoluser(e.SortExpression, "Desc");
        }
    }

    protected void Btn_Detail_Click(object sender, System.EventArgs e)
    {
        //Get the button that raised the event
        Button btn = (Button)sender;
        //Get the row that contains this button
        GridViewRow row = (GridViewRow)btn.NamingContainer;
        String Detail_Text = "";
        div_detail.Visible = true;
        if (row != null){
            int rowindex = row.RowIndex;
        }
        ss.AppendLine("Get-Msoluser -UserPrincipalName " + row.Cells[0].Text);
        ps.AddScript(ss.ToString());
        results = ps.Invoke();
        Detail_Text += "<table  align=center>";
        Detail_Text += " <h3  align=center>" + row.Cells[0].Text + "</h3>";
        foreach (PSObject obj in results){
            foreach (PSProperty prop in obj.Properties)
            {
                String propName = prop.Name;
                Object propValue = prop.Value;
                if (propValue != null){
                    /*
                    Detail_Text += "----[Begin: " + prop.Name + "]---</br>";
                    Detail_Text += "    TypeNameOfValue:  " + prop.TypeNameOfValue ;
                    Detail_Text += "    MemberType:       " + prop.MemberType.ToString();
                    Detail_Text += "</br>";
                    */
                    if (propValue is ICollection)
                    {
                        ICollection collection = (ICollection)propValue;
                        //Detail_Text += "    Multi-valued Property:";
                        foreach (object value in collection){
                            Detail_Text += "<tr><td>" + prop.Name + "</td><td>" + value.ToString() + "</td></tr>\n";
                        }
                    }
                    else{
                        Detail_Text += "<tr><td>" + prop.Name + "</td><td>" + propValue.ToString() + "</td></tr>\n";
                    }
                }
            }
        }
        Detail_Text += "</table>";
        div_detail.InnerHtml = Detail_Text;
    }
    protected void Btn_SetPwd_Click(object sender, System.EventArgs e)
    {
        //Get the button that raised the event
        Button btn = (Button)sender;
        //Get the row that contains this button
        GridViewRow row = (GridViewRow)btn.NamingContainer;
        if (row != null)
        {
            int rowindex = row.RowIndex;
        }
        div_SetPwd.Visible = true;
        UPN.Text = row.Cells[0].Text;
    }
    protected void Do_Pwd_Click(object sender, System.EventArgs e)
    {
        ss.AppendLine("Set-MsolUserPassword -UserPrincipalName \"" + UPN.Text + "\" -NewPassword \"" + NewPwd.Text + "\" -ForceChangePassword $false ");
        ps.AddScript(ss.ToString());
        results = ps.Invoke();
        psErr = ps.Streams.Error;
        if ((psErr != null) && (psErr.Count > 0))
        {
            foreach (ErrorRecord er in psErr)
            {
                Response.Write(er.Exception.ToString());
            }
        }
        else
        {
            Response.Write("Set-Password Success!!!");
        }
    }
    protected void Btn_Remove_Click(object sender, System.EventArgs e)
    {
        //Get the button that raised the event
        Button btn = (Button)sender;
        //Get the row that contains this button
        GridViewRow row = (GridViewRow)btn.NamingContainer;
        if (row != null)
        {
            int rowindex = row.RowIndex;
        }
        div_Remove.Visible = true;
        Lab_Warn.Text = "<b>Do you really want to remove</b>";
        Lab_Account.Text = row.Cells[0].Text;
    }
    protected void Do_Remove_Click(object sender, System.EventArgs e)
    {
        ss.AppendLine("Remove-MsolUser -UserPrincipalName \"" + Lab_Account.Text + "\" -force ");
        ps.AddScript(ss.ToString());
        results = ps.Invoke();
        psErr = ps.Streams.Error;
        if ((psErr != null) && (psErr.Count > 0))
        {
            foreach (ErrorRecord er in psErr)
            {
                Response.Write(er.Exception.ToString());
            }
        }
        else
        {
            Response.Write("Remove " + Lab_Account.Text + " Success!!!");
        }
    }

    /* SetUsr背後執行的ps與資料表DataTable建立, DataView設定 */
    protected void GetMsoluser(string field, string sort)
    {
        div_set.Visible = true;
        ss.AppendLine("Get-Msoluser -all");
        ps.AddScript(ss.ToString());
        results = ps.Invoke();

        userList = CreateUserTable();

        foreach (PSObject obj in results)
        {
            DataRow newUser = userList.NewRow();
            newUser["UserPrincipalName"] = obj.Properties["UserPrincipalName"].Value;
            newUser["DisplayName"] = obj.Properties["DisplayName"].Value;
            newUser["FirstName"] = obj.Properties["FirstName"].Value;
            newUser["LastName"] = obj.Properties["LastName"].Value;
            newUser["Department"] = obj.Properties["Department"].Value;
            newUser["UsageLocation"] = obj.Properties["UsageLocation"].Value;
            newUser["IsLicensed"] = obj.Properties["IsLicensed"].Value;
            userList.Rows.Add(newUser);
        }
        // Bind data list to grid view.
        this.GridView1.Visible = true;
        // DataTable中的DataView排列設定
        userList.DefaultView.Sort = field + " " + sort;
        GridView1.DataSource = userList;
        GridView1.DataBind();
    }

    /* SetUsr的usr的DataTable欄位宣告 */
    private DataTable CreateUserTable()
    {
        DataTable userListTemp = new DataTable();
        userListTemp.Columns.Add("UserPrincipalName");
        userListTemp.Columns.Add("DisplayName");
        userListTemp.Columns.Add("FirstName");
        userListTemp.Columns.Add("LastName");
        userListTemp.Columns.Add("Department");
        userListTemp.Columns.Add("UsageLocation");
        userListTemp.Columns.Add("IsLicensed");
        return userListTemp;
    }

/********* AddUsr *********/
    /* Admin Function Button Click - AddUsr */
    protected void AddUsr_Click(Object sender, EventArgs e)
    {
        div_add.Visible = true;
        ss.AppendLine("Get-MsolAccountSku");
        ps.AddScript(ss.ToString());
        results = ps.Invoke();
        DataTable license = new DataTable();
        license.Columns.Add("AccountSkuId");
        license.Columns.Add("ActiveUnits");
        license.Columns.Add("WarningUnits");
        license.Columns.Add("ConsumedUnits");

        dl.Font.Size = 12;
        dl.Items.Clear();
        foreach (PSObject obj in results)
        {
            DataRow newUser = license.NewRow();
            newUser["AccountSkuId"] = obj.Properties["AccountSkuId"].Value;
            newUser["ActiveUnits"] = obj.Properties["ActiveUnits"].Value;
            newUser["WarningUnits"] = obj.Properties["WarningUnits"].Value;
            newUser["ConsumedUnits"] = obj.Properties["ConsumedUnits"].Value;
            license.Rows.Add(newUser);

            dl.Items.Add(new ListItem(obj.Properties["AccountSkuId"].Value.ToString(), obj.Properties["AccountSkuId"].Value.ToString()));
        }
        // Bind data list to grid view.
        this.GridView2.Visible = true;
        GridView2.DataSource = license;
        GridView2.DataBind();       
    }

    /* AddUsr中的add button click */
    protected void Btn_add_Click(Object sender, EventArgs e)
    {
        div_add.Visible = true;
        ss.AppendLine("New-MsolUser -FirstName \"" + Add_first.Text + "\" -LastName \""
            + Add_last.Text + "\" -UserPrincipalName  \"" + Add_upn.Text + "\" -DisplayName  \""
            + Add_display.Text + "\" -Password  \"" + Add_password.Text + "\" -ForceChangePassword  $false");
        ss.AppendLine("Set-MsolUser -UserPrincipalName \"" + Add_upn.Text + "\" -UsageLocation TW;");
        ss.AppendLine("Set-MsolUserLicense -UserPrincipalName \"" + Add_upn.Text + "\" -AddLicenses \"" + dl.SelectedItem.Value.ToString() + "\"");
        ps.AddScript(ss.ToString());
        results = ps.Invoke();
        psErr = ps.Streams.Error;
        if ((psErr != null) && (psErr.Count > 0))
        {
            foreach (ErrorRecord er in psErr)
            {
                Response.Write(er.Exception.ToString());
            }
        }
        else {
            Response.Write("Add-MsolUser Success!!!");
        }

    }
    /********* AddMultUsr *********/
    protected void AddMultUsr_Click(Object sender, EventArgs e)
    {
        div_addmult.Visible = true;
        Lab_up.Text = "";
        ss.AppendLine("Get-MsolAccountSku");
        ps.AddScript(ss.ToString());
        results = ps.Invoke();
        DataTable license = new DataTable();
        license.Columns.Add("AccountSkuId");
        license.Columns.Add("ActiveUnits");
        license.Columns.Add("WarningUnits");
        license.Columns.Add("ConsumedUnits");

        dl2.Font.Size = 12;
        dl2.Items.Clear();
        foreach (PSObject obj in results)
        {
            DataRow newUser = license.NewRow();
            newUser["AccountSkuId"] = obj.Properties["AccountSkuId"].Value;
            newUser["ActiveUnits"] = obj.Properties["ActiveUnits"].Value;
            newUser["WarningUnits"] = obj.Properties["WarningUnits"].Value;
            newUser["ConsumedUnits"] = obj.Properties["ConsumedUnits"].Value;
            license.Rows.Add(newUser);

            dl2.Items.Add(new ListItem(obj.Properties["AccountSkuId"].Value.ToString(), obj.Properties["AccountSkuId"].Value.ToString()));
        }
        // Bind data list to grid view.
        this.GridView3.Visible = true;
        GridView3.DataSource = license;
        GridView3.DataBind();       
    }
    protected void btn_upload_Click(object sender, EventArgs e)
    {
        div_addmult.Visible = true;
        DateTime now = DateTime.Now;
        if (fileupload1.HasFile){
            try{
                if (fileupload1.PostedFile.ContentType == "application/vnd.ms-excel"){
                    if (fileupload1.PostedFile.ContentLength < 1024000){
                        //string filename = Path.GetFileName(fileupload1.FileName);
                        string filename = now.ToString() +".csv";
                        filename = filename.Replace(' ', '_');
                        filename = filename.Replace(':', '_');
                        filename = filename.Replace('/', '_');
                        //"IIS_IUSERS"add "write authority" in "CSV folder"
                        //fileupload1.SaveAs(Server.MapPath("~/CSV/") + filename);
                        fileupload1.SaveAs(Server.MapPath("~/CSV/") + filename);
                        //File uploaded successfully!
                        Lab_up.Text = "Import file...Please wait a moment";
                        ss.AppendLine("Import-Csv -Path  \"" + @"C:\inetpub\wwwroot\CSV\" 
                            + filename + "\" | ForEach-Object {"
                            + "New-MsolUser -FirstName $_.FirstName -LastName $_.LastName -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -Password $_.Password -ForceChangePassword $false -Department $_.Department;"
                            + "Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation TW;"
                            + "Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses  \"" + dl2.SelectedItem.Value.ToString()
                            + "\" ;} ");
                        ps.AddScript(ss.ToString());
                        results = ps.Invoke();
                        psErr = ps.Streams.Error;
                        if ((psErr != null) && (psErr.Count > 0))
                        {
                            foreach (ErrorRecord er in psErr)
                            {
                                Response.Write(er.Exception.ToString());
                            }
                        }
                        else
                        {
                            Lab_up.Text= "Add-MsolUser Success!!!";
                        }                           
                    }
                    else{
                        Lab_up.Text = "File maximum size is 1MB";
                    }
                }
                else{
                    Lab_up.Text = "Only CSV files are accepted!";
                }
            }
            catch (Exception exc){
                Lab_up.Text = "The file could not be imported. The following error occured: " + exc.Message;
            }
        }
    }
    protected void btn_download_Click(object sender, EventArgs e)
    {
        string fileName = @"C:\inetpub\wwwroot\CSV\sample_new_users.csv";
        Response.Clear();
        Response.AddHeader("content-disposition", "attachment;filename=sample_new_users.csv");
        Response.ContentType = @"application/vnd.ms-excel";
        System.IO.FileStream downloadFile =
            new System.IO.FileStream(fileName, System.IO.FileMode.Open);
        downloadFile.Close();
        Response.WriteFile(fileName);
        Response.Flush();
        Response.End();
    }

   /********* CheckLog *********/ 
    protected void CheckLog_Click(Object sender, EventArgs e)
    {
        div_log.Visible = true;
        String Log_Text = "";
        Log_Text += "<table  align=center>";
        Log_Text += " <h3  align=center> Login Record </h3>";
        Log_Text += "<tr><th>UserPrincipalName</th><th>Login Time</th><th>IP address</th></tr>";
        SqlConnection Conn = new SqlConnection("server =localhost\\SQLEXPRESS; uid = vincechen; pwd = p@ssword; database=o365");
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from o365_log", Conn);
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            Log_Text += "<tr><td>" + dr["upn"].ToString() + "</td><td>" + dr["login"].ToString() + "</td><td>" + dr["ip"].ToString() + "</td></tr>";
        }
        cmd.Cancel();
        dr.Close();
        Conn.Close();
        Log_Text += "</table>";
        div_log.InnerHtml = Log_Text;         
    }
    /* GetIPAddress */
    protected string GetIPAddress()
    {
        return Request.ServerVariables["HTTP_X_FORWARDED_FOR"] ?? Request.ServerVariables["REMOTE_ADDR"];
    }
}





    
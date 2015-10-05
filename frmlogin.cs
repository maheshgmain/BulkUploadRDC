using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Common;
using System.Data.SqlClient;
namespace UploadApplication
{
    public partial class frmlogin : Form
    {
        public frmlogin()
        {
            InitializeComponent();
        }

        private void btnsubmit_Click(object sender, EventArgs e)
        {
            frmmenu frm = new frmmenu(); this.Hide(); frm.Show();
            //if (SignIn(txtusername.Text.Trim().ToLower(), txtpassword.Text.Trim()) > 0)
            //{ 
            //    frmmenu frm = new frmmenu(); this.Hide(); frm.Show(); }
            //else {
            //    MessageBox.Show("Login failed.");
            //}
        }
        public Int32 SignIn(string UserName, string Password)
        {
            Encryption encryption = new Encryption();
            //OperationClass opration = new OperationClass();
           
            Int32 Returnvalue = 0;
            //opration.CreateWordFile();
            try
            {
                //getExpirydays(UserName);
                string MysaltValue, MyEncryptPassword, NewEncryptPassword = "";
                //dtPasswrod = _UserBal.GetSaltValue(txtUserName.Text.ToString().Trim(), txtPassword.Text.ToString().Trim());
                DataTable dtsaltvalue = GetDataTable(
                        string.Format(@"SELECT Username,[Password],Saltvalue,userID  FROM   tblUserMaster
	        WHERE  Username = '{0}' AND   status =1", UserName));
                if (dtsaltvalue.Rows.Count > 0)
                {
                    MysaltValue = dtsaltvalue.Rows[0]["Saltvalue"].ToString();
                    MyEncryptPassword = dtsaltvalue.Rows[0]["Password"].ToString();
                    NewEncryptPassword = encryption.CreatePasswordHash(Password, MysaltValue);

                    if (MyEncryptPassword == NewEncryptPassword)
                    {

                        DataTable dtuserdetails = GetDataTable(
                        string.Format(@"IF EXISTS ( SELECT Username,[Password],Saltvalue FROM tblUserMaster
	                              WHERE  Username = '{0}' AND [Password] = '{1}' AND Status = 1
	                )
                    begin
                        SELECT tbluserdetail.UserID,tbluserdetail.FirstName+' '+tbluserdetail.lastname as [Name],
                        tbluserdetail.GroupID, tblWorkGroupMaster.GroupName,tblUserMaster.loginstatus,tblUserMaster.userName
                        FROM   tbluserdetail  	INNER JOIN 
                          tblWorkGroupMaster		ON tbluserdetail.GroupId = tblWorkGroupMaster.GroupId 
			                        inner join
                          tblUserMaster			on tbluserdetail.userid = tblUserMaster.userid
                        WHERE tbluserdetail.userid = ( SELECT distinct Userid 
							                          FROM   tblusermaster 
							                          WHERE  Username = '{0}' )
                    end
                    ", UserName, NewEncryptPassword));

                        if (dtuserdetails.Rows.Count > 0)
                        {
                            int intLoginStatus = Convert.ToInt32(dtuserdetails.Rows[0]["LoginStatus"]);
                            if (intLoginStatus == 0)
                            {
                                Program.strUserName = Convert.ToString(dtuserdetails.Rows[0]["UserName"]);
                                Program.UserId = dtuserdetails.Rows[0]["UserID"].ToString();
                                Returnvalue = 50; //redirect to home page
                            }
                            else
                            {
                                Program.strUserName = Convert.ToString(dtuserdetails.Rows[0]["UserName"]);
                                Program.UserId = dtuserdetails.Rows[0]["UserID"].ToString();
                                Returnvalue = 100; //redirect to home page
                            }
                        }
                        else
                        {
                            Returnvalue = -3;//invalid user;
                        }

                    }
                    else
                    {
                        //ScriptManager.RegisterClientScriptBlock(Page, typeof(UpdatePanel), "msg", "alert('Invalid user name or password.')", true);
                        //return;
                        Returnvalue = -2;//invalid password;
                    }
                }
                else
                {
                    //ScriptManager.RegisterClientScriptBlock(Page, typeof(UpdatePanel), "msg", "alert('Invalid user name or password.')", true);
                    //return;
                    Returnvalue = -1;//invalid user name;
                }

            }
            catch (Exception ex)
            {
                Returnvalue = -3;
            }
            return Returnvalue;
        }
        public  DataTable GetDataTable(string sqlquery)
        {
            SqlConnection Sqlconn = new SqlConnection(ConfigurationSettings.AppSettings["SqlCon"].ToString());
            if (Sqlconn.State == ConnectionState.Closed)
                Sqlconn.Open();

            SqlCommand Sqlcomm = new SqlCommand(sqlquery, Sqlconn);

            SqlDataReader sdr = Sqlcomm.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(sdr);
            sdr.Close();

            return dt;
        }
    }
}

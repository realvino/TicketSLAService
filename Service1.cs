using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Text;

namespace WindowsServiceCS
{
    public partial class Service1 : ServiceBase
    {
        public string EmailId, password, smtpServer;
        public bool useSSL;
        public int Port, i = 0;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Simple Service stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            System.Timers.Timer time = new System.Timers.Timer();

            time.Start();

            time.Interval = 60000;

            time.Elapsed += SchedularCallback;
        }

        private void SchedularCallback(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {

                string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                //SqlConnection con=new SqlConnection(constr);
                DataTable dt = new DataTable();
                string query = "SELECT * FROM TicketRegistration WHERE IsSent='" + false + "'";

                using (SqlConnection con = new SqlConnection(constr))
                {

                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;                      
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }

                foreach (DataRow row in dt.Rows)
                {

                   // WriteToFile("Enter into the Ticket Reg_table");

                    string seniarty = string.Empty;
                    string status = string.Empty;
                    string Duration = string.Empty;
                    string TicketId = string.Empty;
                    string CreationTime = string.Empty;
                    string TicketNo = string.Empty;
                    string StartTime = string.Empty;
                    string ResumeCount = string.Empty;
                    int SLADuration = 0;
                    int ResumeCount1 = 0;
                    //int CurrentSLATime = 0;
                    bool Hold = false;
                    int TicketCount = 0;

                    seniarty = row["PriorityId"].ToString();
                    status = row["CurrentMileStoneId"].ToString();
                    StartTime = row["StartTime"].ToString();
                    ResumeCount = row["ResumeTimeCount"].ToString();
                    Hold = (bool)row["Hold"];
                   // WriteToFile("Enter into the SupportMilestoneSeniority");

                    SqlConnection conn1 = new SqlConnection(constr);
                    conn1.Open();
                    string MilestoneSeniority = "select * from SupportMilestoneSeniority where SeniorityId='" + Convert.ToInt32(row["PriorityId"]) + "' AND SupportMilestoneId ='" + Convert.ToInt32(row["CurrentMileStoneId"]) + "'";
                   // WriteToFile(" SupportMilestoneSeniority  " + MilestoneSeniority);

                    SqlCommand cmd = new SqlCommand(MilestoneSeniority, conn1);
                    SqlDataReader dr1 = cmd.ExecuteReader();
                    if (dr1.Read())
                    {
                        Duration = dr1["Duration"].ToString();
                        SLADuration = Int32.Parse(Duration) * 60*60;
                        TicketId = row["Id"].ToString();
                        TicketNo = row["TicketNo"].ToString();
                        
                    }
                    conn1.Close();
                    //WriteToFile("Duration   " + SLADuration);                 

                    //WriteToFile("Enter into the CommonActivity");

                    SqlConnection conn2 = new SqlConnection(constr);
                    conn2.Open();
                    string CommonActivity = "SELECT CreationTime FROM CommonActivity WHERE Id = (SELECT MAX(Id) FROM [dbo].[CommonActivity] where TicketId = " + TicketId + ")";
                    SqlCommand cmmd = new SqlCommand(CommonActivity, conn2);
                    SqlDataReader dr11 = cmmd.ExecuteReader();
                    if (dr11.Read())
                    {
                        CreationTime = dr11["CreationTime"].ToString();
                    }
                    conn2.Close();

                    DateTime Creation = Convert.ToDateTime(CreationTime).AddSeconds(SLADuration);
                    DateTime now = DateTime.Now;

                    TimeSpan varTime = (DateTime)now - (DateTime)Creation;
                    int CurrentSLATime = 0;
                    if (now > Creation)
                    {
                        CurrentSLATime = 0;
                    }
                    else
                    {
                        CurrentSLATime = (int)varTime.TotalMinutes;
                    }

                    //WriteToFile(" "+(DateTime)now +" "+ (DateTime)Creation + "Timedetails   " + CurrentSLATime);

                    if (StartTime == null || StartTime == "")
                    {
                        TicketCount = CurrentSLATime;
                    }
                    else if (StartTime != null && Hold == false)
                    {
                        DateTime StartTime1 = Convert.ToDateTime(StartTime);
                        ResumeCount1 = Int32.Parse(ResumeCount);
                        //WriteToFile("StartTime " + StartTime);
                        DateTime dt12 = Convert.ToDateTime(StartTime1);
                        dt12 = dt12.AddSeconds(ResumeCount1);
                        DateTime dt22 = DateTime.Now;

                        TimeSpan varTime1 = dt12 - dt22;

                        TicketCount = varTime1.Minutes * 60;

                       // WriteToFile("SLA Time " + dt12.ToString());
                       // WriteToFile("Current " + dt22.ToString());
                       // WriteToFile("Count " + varTime1.Minutes);
                       // WriteToFile("TicketCount " + TicketCount);


                    }
                    else
                    {

                        TicketCount = ResumeCount1;
                    }

                    if (TicketCount <= 0)
                    {
                        //WriteToFile("Mail Sending in Ticket ");

                        //WriteToFile("SLA Time CurrentSLATime" + CurrentSLATime);
                        //WriteToFile("SLA Time SLADuration" + SLADuration);

                        string Title = "WatchNet Notification Alert";
                        string SubTitle = string.Empty;
                        string TaskName = string.Empty;
                        string RequestFrom = string.Empty;
                        string RequestTo = string.Empty;
                        string RequestFromRoleName = string.Empty;
                        string ProjectName = string.Empty;
                        string RefNo = string.Empty;                    
                        string ReceiverMailId = string.Empty;
                        string CustomerCompany = string.Empty;
                        string SalesMail = string.Empty;

                        SubTitle = "SLA Failed : " + TicketNo;

                        SqlConnection con2 = new SqlConnection(constr);
                        con2.Open();
                        string qur1 = "select dbo.AbpUsers.UserName,dbo.AbpRoles.DisplayName as RoleName from dbo.AbpUserRoles inner join dbo.AbpUsers on (dbo.AbpUsers.Id=dbo.AbpUserRoles.UserId) inner join dbo.AbpRoles on (dbo.AbpRoles.Id=dbo.AbpUserRoles.RoleId) where UserId='" + row["CreatorUserId"] + "'";
                        SqlCommand cmd1 = new SqlCommand(qur1, con2);
                        SqlDataReader dr2 = cmd1.ExecuteReader();
                        if (dr2.Read())
                        {
                            RequestFrom = dr2["UserName"].ToString();
                            RequestFromRoleName = dr2["RoleName"].ToString();
                        }
                        con2.Close();

                        SqlConnection con3 = new SqlConnection(constr);
                        con3.Open();
                        string qur2 = " select AbpUsers.EmailAddress,AbpUsers.UserName from AbpUsers join AbpUserRoles  on AbpUsers.Id = AbpUserRoles.UserId  WHERE AbpUserRoles.RoleId = 8";
                        SqlCommand cmd2 = new SqlCommand(qur2, con3);
                        SqlDataReader dr3 = cmd2.ExecuteReader();
                        if (dr3.Read())
                        {
                            RequestTo = dr3["UserName"].ToString();
                            ReceiverMailId = dr3["EmailAddress"].ToString();
                        }
                        con3.Close();
                        SqlConnection con10 = new SqlConnection(constr);
                        con10.Open();
                        string qurcom = "select dbo.Companies.CompanyName as CompanyName,AbpUsers.EmailAddress as SalesEmail from dbo.TicketRegistration inner join dbo.Companies on (dbo.TicketRegistration.CompanyId=dbo.Companies.Id) inner join dbo.AbpUsers on (dbo.AbpUsers.Id=dbo.Companies.AccountManagerId) where dbo.TicketRegistration.Id='" + row["Id"] + "'";
                        SqlCommand cmds = new SqlCommand(qurcom, con10);
                        SqlDataReader dr10 = cmds.ExecuteReader();
                        if (dr10.Read())
                        {
                            CustomerCompany = dr10["CompanyName"].ToString();
                            SalesMail = dr10["SalesEmail"].ToString();
                        }
                        con10.Close();

                        SqlConnection con4 = new SqlConnection(constr);
                        con4.Open();
                        string qur3 = "select * from dbo.AbpSettings";
                        SqlCommand cmd3 = new SqlCommand(qur3, con4);
                        DataTable dt1 = new DataTable();
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd3))
                        {
                            sda.Fill(dt1);
                        }
                        foreach (DataRow row1 in dt1.Rows)
                        {
                            if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.Port")
                            {
                                Port = Convert.ToInt32(row1["Value"]);
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.UserName")
                            {
                                EmailId = row1["Value"].ToString();
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.Password")
                            {
                                password = row1["Value"].ToString();
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.EnableSsl")
                            {
                                useSSL = Convert.ToBoolean(row1["Value"]);
                            }
                            else if (row1["Name"].ToString() == "Abp.Net.Mail.Smtp.Host")
                            {
                                smtpServer = row1["Value"].ToString();
                            }
                        }

                        if(Port==0)
                        {
                            Port = 25;
                        }

                        con4.Close();

                        WriteToFile("Trying to send email to: " + ReceiverMailId);

                        using (MailMessage mm = new MailMessage(EmailId, ReceiverMailId))
                        {
                            mm.Subject = SubTitle;
                            
                            //var mailMessage = new StringBuilder();
                            //mailMessage.AppendLine("<b>" + "User Name" + "</b>: " + RequestTo + "<br />");
                            //mailMessage.AppendLine("<b>" + "Notification From" + "</b>: " + RequestFrom + "<br />");
                            //mailMessage.AppendLine("<b>" + "Notification From Role" + "</b>: " + RequestFromRoleName + "<br />");
                           
                            //mailMessage.AppendLine("<br />");

                            string mailbodycontent = "<b>Dear " + RequestTo + "</b><br /><br /><span>You are receiving this alert due to SLA Fail</span><br /><br /><span>Company Name : <b style='color:blue;font-weight: bolder;'>"+ CustomerCompany +"</b></span><br /><br /><span>Ticket No : <b style='color:blue;font-weight: bolder;'>"+ TicketNo +"</b></span><br /><br /><span>Kindly take necessary action.</span><br /><br />";
                            string body = string.Empty;
                            string path = Directory.GetCurrentDirectory();
                            using (StreamReader reader = new StreamReader("C:\\hostingemail\\WatchNetSLATemplate.html"))
                            {
                                body = reader.ReadToEnd();
                            }
                            body = body.Replace("{EMAIL_TITLE}", Title);
                            body = body.Replace("{EMAIL_SUB_TITLE}", SubTitle);
                            body = body.Replace("{Mail_Content}", mailbodycontent);
                            //body = body.Replace("{EMAIL_BODY}", mailMessage.ToString());
                            //body = body.Replace("{Date}",DateTime.Now.Year.ToString());

                            var im = "http://localhost:6240/";
                            var im1 = "http://localhost:6240/";
                            var ipath = "Common/Images/logo.jpg";
                            var ipath1 = "Common/Images/Teamworks%20logo1.png";

                            try
                            {
                                SqlConnection con11 = new SqlConnection(constr);
                                con11.Open();
                                string qur11 = "select Value from dbo.AbpSettings where Name='App.General.WebSiteRootAddress'";
                                SqlCommand cmd11 = new SqlCommand(qur11, con11);
                                SqlDataReader dr13 = cmd11.ExecuteReader();
                                if (dr13.Read())
                                {
                                    im = dr13["Value"].ToString();
                                    im1 = dr13["Value"].ToString();
                                }
                                con11.Close();
                                ////var image = context1.Database.SqlQuery<ret>("select Value from dbo.AbpSettings where Name='App.General.WebSiteRootAddress'").FirstOrDefaultAsync();
                                //if (qur11.Result.Value != null)
                                //{
                                //    im = image.Result.Value;
                                //    im1 = image.Result.Value;
                                //}
                            }
                            catch (Exception ex)
                            {

                            }

                            im = im + ipath;
                            body = body.Replace("{WatchNet_Logo}", im);
                            body = body.Replace("{TeamWorks_logo}", im1 + ipath1);

                            mm.Body = body;
                            mm.IsBodyHtml = true;
                            SmtpClient smtp = new SmtpClient();
                            smtp.Host = smtpServer;
                            smtp.EnableSsl = useSSL; 
                            System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                            credentials.UserName = EmailId;
                            credentials.Password = password;
                            smtp.UseDefaultCredentials = true;
                            smtp.Credentials = credentials;
                            smtp.Port = Port;
                            if(SalesMail != "")
                            {
                                mm.To.Add(SalesMail); 
                            }
                            smtp.Send(mm);
                            WriteToFile("Email sent successfully to: " + RequestTo + " " + ReceiverMailId + " " + SalesMail);

                            SqlConnection conFinal = new SqlConnection(constr);
                            conFinal.Open();
                            string qurfinal = "update TicketRegistration set IsSent='" + true + "'where Id='" + Convert.ToInt32(row["Id"]) + "' ";
                            SqlCommand cmdfinal = new SqlCommand(qurfinal, conFinal);
                            cmdfinal.ExecuteNonQuery();
                            conFinal.Close();
                        }
                    }

                }
                this.ScheduleService();
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("WatchNet Ticket Sla Test Service"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void WriteToFile(string text)
        {
            string path = "C:\\hostinglog\\TicketSlaTestService.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }    
    }
}

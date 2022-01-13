using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Xml;

namespace PSCService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer(); // name space(using System.Timers;)  
        public Service1()
        {
            this.CanHandleSessionChangeEvent = true;
            InitializeComponent();
            //Microsoft.Win32.SystemEvents.SessionSwitch += new Microsoft.Win32.SessionSwitchEventHandler(SystemEvents_SessionSwitch);
            //Console.ReadLine();
            CreatePCName();
        }
        protected override void OnStart(string[] args)
        {
            //SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
            Console.ReadLine();
        }
        string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["kintai"].ConnectionString;
        string CompanyID = System.Configuration.ConfigurationManager.AppSettings["CompanyID"];
        string isInstallPC = System.Configuration.ConfigurationManager.AppSettings["isInstallPC"];
        string MachineName = Environment.MachineName;

        protected override void OnStop()
        {
            //SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
        }
        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            //string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string path = @"C:\\WinLog\\";
            DateTime dateTime = DateTime.Now;
            string file = "sessionlog" + dateTime.Year + dateTime.Month + dateTime.Day+".txt";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            // get pc name
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string username = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];
            
            string MachineName = Environment.MachineName;

            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("{0} : Power mode changed = {1},userName = {2}", dateTime, changeDescription.Reason, MachineName));
            //sb.Append(string.Format("{0} : test = {1},userName = {2}", MachineName1, changeDescription.Reason, username));
            sb.Append(Environment.NewLine);
            File.AppendAllText(path + file, sb.ToString());
            sb.Clear();

            string session =Convert.ToString( changeDescription.Reason);

            //connect sql server data insert
            
          
            using (var conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = @"INSERT INTO workingtime(UserName,AuthDatetime,AuthDate,AuthTime,Status,CompanyID) 
                            VALUES(@param1,@param2,@param3,@param4,@param5,@param6)";

                    cmd.Parameters.AddWithValue("@param1", MachineName);
                    cmd.Parameters.AddWithValue("@param2", dateTime);
                    cmd.Parameters.AddWithValue("@param3", dateTime.Date);
                    cmd.Parameters.AddWithValue("@param4", dateTime.ToLongTimeString());
                    cmd.Parameters.AddWithValue("@param5", session);
                    cmd.Parameters.AddWithValue("@param6", CompanyID);

                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException e)
                    {
                        sb = new StringBuilder();
                        sb.Append(string.Format("{0} : error= {1}", dateTime, e.Message.ToString()));
                        sb.Append(Environment.NewLine);
                        File.AppendAllText(path + file, sb.ToString());
                        sb.Clear();
                    }
                }
            }
            base.OnSessionChange(changeDescription);
        }
        public void CreatePCName()
        {
            string path = @"C:\\WinLog\\";
            DateTime dateTime = DateTime.Now;
            string file = "sessionlog" + dateTime.Year + dateTime.Month + dateTime.Day + ".txt";
            //if flag isInstallPC = 0 then install
            //if (isInstallPC =="0")
            //{
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);

            xmlDoc.SelectSingleNode("//appSettings/add[@key='isInstallPC']").Attributes["value"].Value = "1";
            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            ConfigurationManager.RefreshSection("appSettings");
            if (!CheckPCName())
                {
                    // write sql db
                    using (var conn = new SqlConnection(connectionString))
                    {
                        using (SqlCommand cmd = new SqlCommand())
                        {
                            string EmployeeID = "";
                            cmd.Connection = conn;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = @"INSERT INTO [employee]([EmployeeID],[EmployeeName],[PCName],[CompanyID],[CreatedDate],UpdateTime) 
                    VALUES(@param1,@param2,@param3,@param4,@param5,@param6)";
                            cmd.Parameters.AddWithValue("@param1", EmployeeID);
                            cmd.Parameters.AddWithValue("@param2", MachineName);
                            cmd.Parameters.AddWithValue("@param3", MachineName);
                            cmd.Parameters.AddWithValue("@param4", CompanyID);
                            cmd.Parameters.AddWithValue("@param5", dateTime);
                            cmd.Parameters.AddWithValue("@param6", dateTime);
                            try
                            {
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();

                                StringBuilder sb = new StringBuilder();
                                sb = new StringBuilder();
                                sb.Append(string.Format("{0} : {1}", dateTime, "インストール済みです。"));
                                sb.Append(Environment.NewLine);
                                File.AppendAllText(path + file, sb.ToString());
                                sb.Clear();
                            }
                            catch (SqlException ex)
                            {
                                StringBuilder sb = new StringBuilder();
                                sb = new StringBuilder();
                                sb.Append(string.Format("{0} : error= {1}", dateTime, ex.Message.ToString()));
                                sb.Append(Environment.NewLine);
                                File.AppendAllText(path + file, sb.ToString());
                                sb.Clear();
                            }
                        }
                    }
                }
        }
        bool CheckPCName()
        {
            using (var conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = @"SELECT * FROM employee WHERE PCName = @param1 and CompanyID =@param2 ";
                    cmd.Parameters.AddWithValue("@param1", MachineName);
                    cmd.Parameters.AddWithValue("@param2", CompanyID);
                    try
                    {
                        conn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {

                            while (reader.Read())
                            {
                                return true;
                            }
                            return false;
                        }
                    }
                    catch (SqlException e)
                    {
                        return true;
                    }
                }
            }
        }
        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        }
        protected override void OnStart(string[] args)
        {
            //SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
            Console.ReadLine();
        }

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
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string username = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];

            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("{0} : Power mode changed = {1},userName = {2}", dateTime, changeDescription.Reason, username));
            sb.Append(Environment.NewLine);
            File.AppendAllText(path + file, sb.ToString());
            sb.Clear();

            string session =Convert.ToString( changeDescription.Reason);

            //connect sql server data insert
            var connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["kintai"].ConnectionString;
            using (var conn = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = @"INSERT INTO workingtime(UserName,AuthDatetime,AuthDate,AuthTime,Status) 
                            VALUES(@param1,@param2,@param3,@param4,@param5)";

                    cmd.Parameters.AddWithValue("@param1", username);
                    cmd.Parameters.AddWithValue("@param2", dateTime);
                    cmd.Parameters.AddWithValue("@param3", dateTime.Date);
                    cmd.Parameters.AddWithValue("@param4", dateTime.ToLongTimeString());
                    cmd.Parameters.AddWithValue("@param5", session);

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

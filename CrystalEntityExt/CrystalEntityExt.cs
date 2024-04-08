using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using LSEXT;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using System.Diagnostics;


namespace CrystalEntityExt
{
    [ComVisible(true)]
    [ProgId("NautilusExtensions.CrystalEntityExt")]
    public class CrystalEntityExt : IEntityExtension
    {

        #region private members
        private INautilusServiceProvider _sp;
        private OracleConnection _connection;
        private string sql = "";
        private OracleCommand cmd;
        private OracleDataReader reader;
        private string _connectionString;

        #endregion

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                _sp = Parameters["SERVICE_PROVIDER"];
                Connect();
                var records = Parameters["RECORDS"];
                var PlateId = records.Fields["SAMPLE_ID"].Value;

                //  MessageBox.Show("___" + PlateId);
                report(PlateId);
            }
            catch (Exception e)
            {
                MessageBox.Show("Err At Execute: " + e.Message);
            }
        }
        public void report(dynamic plateId)
        {
            ReportDocument CR = new ReportDocument();
            var crTableLoginInfo = new TableLogOnInfo();
            var crConnectionInfo = new ConnectionInfo();
            Tables CrTables;
            const string p2 = @"C:\Users\ashim\Desktop\DisplaySamples1.rpt";
            CR.Load(p2);
            crConnectionInfo.ServerName = "PATHOLAB";
            crConnectionInfo.UserID = "lims_sys";
            crConnectionInfo.Password = "lims_sys";
            CR.SetParameterValue("Sample_id", plateId.ToString());
            CrTables = CR.Database.Tables;
            foreach (Table CrTable in CrTables)
            {
                crTableLoginInfo = CrTable.LogOnInfo;
                crTableLoginInfo.ConnectionInfo = crConnectionInfo;
                CrTable.ApplyLogOnInfo(crTableLoginInfo);
            }
            ExportOptions crExportOptions;
            var crDiskFileDestinationOption = new DiskFileDestinationOptions();
            var crFormattypeOptions = new PdfRtfWordFormatOptions();
            crDiskFileDestinationOption.DiskFileName = @"c:\Plate1hila.pdf";
            crExportOptions = CR.ExportOptions;
            {
                crExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                crExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                crExportOptions.DestinationOptions = crDiskFileDestinationOption;
                crExportOptions.FormatOptions = crFormattypeOptions;

            }
            CR.Export();
            CR.Close();
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo(@"c:\Plate1hila.pdf");
            p.Start();
            // p.WaitForExit();
            //File.Delete(@"P:\ziv\Crystal\Plate1hila.pdf");
        }
        public void Connect()
        {
            try
            {
                INautilusDBConnection dbConnection;
                if (_sp != null)
                {
                    dbConnection = _sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                }
                else
                {
                    dbConnection = null;
                }
                if (dbConnection != null)
                {
                    // _username= dbConnection.GetUsername();
                    _connection = GetConnection(dbConnection);
                    //set oracleCommand's connection
                    cmd = _connection.CreateCommand();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Err At Connect: " + e.Message);
            }
        }
        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {
                    _connectionString = ntlsCon.GetADOConnectionString();
                    var splited = _connectionString.Split(';');
                    _connectionString = "";
                    for (int i = 1; i < splited.Count(); i++)
                    {
                        _connectionString += splited[i] + ';';
                    }

                    //create connection
                    connection = new OracleConnection(_connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                        MessageBox.Show("Inconsistent role Security : " + f.Message);
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                    MessageBox.Show("Err At GetConnection: " + e.Message);
                }
            }
            return connection;
        }
        private void WriteLog(string description, string applicaion, string function, string status)
        {
            string log_id = "";
            string log_name = "";
            string Description = description;
            string version = "1";
            string version_status = "A";
            string Application = applicaion;
            string Function = function;
            string created_on = "sysdate";//string.Format("{0:dd/MM/yyyy HH:mm:ss}", DateTime.Now);
            string operator_id = "";
            string Status = status;
            string machine = "";

            try
            {
                sql = "select lims.sq_U_LOG.nextval from dual";
                cmd.CommandText = sql;
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    log_id = reader["NEXTVAL"].ToString();
                    log_name = log_id;
                }
                reader.Close();
                cmd.Dispose();
                sql = "select  lims.lims_env.operator_id from dual ";
                cmd.CommandText = sql;
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    operator_id = reader["OPERATOR_ID"].ToString();
                }
                reader.Close();
                cmd.Dispose();
                sql = "select NAME from lims_sys.workstation where Workstation_ID IN ( SELECT Workstation_ID FROM Lims_Sys.Lims_Session WHERE Lims_Session.Session_ID IN ( SELECT Session_ID FROM Lims_Sys.Current_Session ) ) ";
                cmd.CommandText = sql;
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    machine = reader["NAME"].ToString();
                }
                reader.Close();
                cmd.Dispose();
                sql = "insert into lims_sys.U_LOG (u_log_id,name,description,version,version_status) values('" + log_id + "','" + log_name + "','" + Description + "','" + version + "','" + version_status + "')";
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                sql = "insert into lims_sys.U_LOG_USER (u_log_id,u_application,u_function, u_created_on, u_operator_id,u_status,u_machine) values('" + log_id + "','" + Application + "','" + Function + "'," + created_on + ",'" + operator_id + "','" + Status + "','" + machine + "')";
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error on WriteLog : " + e.Message);
            }
        }

    }
}

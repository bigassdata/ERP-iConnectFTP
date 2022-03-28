using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Timers;

namespace iConnectFTP
{
    public partial class FTPService : ServiceBase
    {
        #region module level declarations

        private const string CONN_STRING = "server=Erp-sql;database=BigAssFans;uid=WebService;password=MKhZx0_N";
        private const string PROGRAM_NAME = "IConnectFTP";
        private const string MODE = "Common"; // "Test" for testing, "Common" for production

        private Timer _t;
        private double _interval = 300000;
        private string _ediPush = "";
        private string _ediPull = "";
        private string _archivePath = "";
        private string _ftpAddress = "";
        private string _ftpPullFolder = "";
        private string _ftpPushFolder = "";
        private string _ftpUsername = "DEL12434";
        private string _ftpPassword = "IN10467";

        #endregion

        #region ServiceBase overrides

        public FTPService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            GetSettings();

            _t = new Timer();
            _t.Elapsed += new ElapsedEventHandler(OnTimer);
            _t.Interval = _interval;
            _t.Start();
        }

        protected override void OnStop()
        {
            _t.Stop();
            _t.Close();
            _t.Dispose();
        }

        #endregion

        #region Timer delegate

        private void OnTimer(object source, ElapsedEventArgs e)
        {
            try
            {
                _t.Stop();

                GetSettings();
                ProcessFiles();
            }
            catch(Exception ex)
            {
                base.EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
            finally
            {
                _t.Start();
            }
        }

        #endregion

        #region settings

        private void GetSettings()
        {
            string sql = $@"
                select ParamName, ParamDataType, ParamCharacter, ParamDate, ParamLogical, ParamInteger, ParamDecimal 
                from BASProgramSettings 
                where ProgramName = '{PROGRAM_NAME}' and ParamGroup = '{MODE}'";

            try
            {
                using (SqlConnection con = new SqlConnection(CONN_STRING))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                switch (r.GetString(r.GetOrdinal("ParamName")).ToLower())
                                {
                                    case "archivepath":
                                        _archivePath = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "edipullpath":
                                        _ediPull = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "edipushpath":
                                        _ediPush = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "ftpaddress":
                                        _ftpAddress = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "ftppullfolder":
                                        _ftpPullFolder = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "ftppushfolder":
                                        _ftpPushFolder = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "ftppassword":
                                        _ftpPassword = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "ftpusername":
                                        _ftpUsername = r.GetString(r.GetOrdinal("ParamCharacter"));
                                        break;
                                    case "postinterval":
                                        _interval = r.GetInt32(r.GetOrdinal("ParamInteger"));
                                        _t.Interval = _interval;
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                base.EventLog.WriteEntry("An error has ocurred while retrieving settings.\r\n\r\n" + ex.Message, EventLogEntryType.Error);
            }
        }

        #endregion

        #region file processing

        private void ProcessFiles()
        {
            base.EventLog.WriteEntry($"Variables used for Directories .\r\n\r\n pull:{_ediPull} push:{_ediPush} archivepath:{_archivePath}", EventLogEntryType.Information);
            
            DirectoryInfo ediPull = null;
            DirectoryInfo ediPush = new DirectoryInfo(_ediPush);
            DirectoryInfo archive = new DirectoryInfo(_archivePath);

            base.EventLog.WriteEntry($"File returned from current EDI Path:.\r\n\r\n {ediPush.EnumerateFiles("*", SearchOption.TopDirectoryOnly).ToList()}", EventLogEntryType.Information);
           

            if (ediPush.Exists)
            {
                if (!archive.Exists)
                    archive.Create();

                using (WebClient wc = new WebClient())
                {
                    wc.Credentials = new NetworkCredential(_ftpUsername, _ftpPassword);


                    ediPush.EnumerateFiles("*", SearchOption.TopDirectoryOnly).ToList().ForEach(f =>

                    {
                        bool failed = false;

                        try
                        {
                            base.EventLog.WriteEntry($"Starting Upload of Files to {_ftpAddress}\r\n\r\n", EventLogEntryType.Information);
                            base.EventLog.WriteEntry($"Combined path: {Path.Combine(_ftpAddress, _ftpPushFolder, f.Name)+ f.FullName}\r\n\r\n", EventLogEntryType.Information);
                            // Waiting to see if Files will need to have spaces removed, added ftp:// as it is needed to send off
                            wc.UploadFile(_ftpAddress + f.Name, f.FullName);
                        }
                        catch (Exception ex)
                        {
                            failed = true;

                            base.EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                        }

                        if (!failed)
                        {
                            base.EventLog.WriteEntry($"Starting Upload of Files....\r\n\r\n", EventLogEntryType.Information);
                            f.MoveTo(Path.Combine(_archivePath, GetFileName(f.Name)));
                        }
                    });
                }
            }
            else
            {
                base.EventLog.WriteEntry("No EDI push folder exists at path " + _ediPush + ". There are no files to process.", EventLogEntryType.Warning);
            }

            if (ediPull.Exists)
            {
                List<string> files = new List<string>();
                FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(Path.Combine(_ftpAddress, _ftpPullFolder));
                ftpRequest.Credentials = new NetworkCredential(_ftpUsername, _ftpPassword);
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;

                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();

                using (StreamReader sr = new StreamReader(ftpResponse.GetResponseStream()))
                {
                    string line = sr.ReadLine();

                    while (!string.IsNullOrEmpty(line))
                    {
                        files.Add(line);

                        line = sr.ReadLine();
                    }
                }

                using (WebClient wc = new WebClient())
                {
                    wc.Credentials = new NetworkCredential(_ftpUsername, _ftpPassword);

                    files.ForEach(f =>
                    {
                        try
                        {
                            wc.DownloadFile(Path.Combine(_ftpAddress, _ftpPullFolder, f), Path.Combine(_ediPull, f));
                        }
                        catch (Exception ex)
                        {
                            base.EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                        }
                    });
                }
            }
            else
            {
                base.EventLog.WriteEntry("No EDI pull folder exists at path " + _ediPull + ". There are no files to process.", EventLogEntryType.Warning);
            }
        }

        private string GetFileName(string fileName)
        {
            string result = fileName;
            List<string> files = Directory.EnumerateFiles(_archivePath, Path.GetFileNameWithoutExtension(fileName) + "*", SearchOption.TopDirectoryOnly).ToList();

            if (files.Count > 0)
            {
                int next = 1;
                List<string[]> arrays = files.Select(f => Path.GetFileNameWithoutExtension(f).Split(new string[] { " (" }, StringSplitOptions.None)).ToList();
                List<int> nums = arrays.Where(a => a.Length > 1).Select(a => int.Parse(a.Last().Replace(")", ""))).ToList();
                if (nums.Count > 0)
                    next = nums.Max() + 1;

                result = string.Format("{0} ({1}){2}.old", Path.GetFileNameWithoutExtension(fileName), next.ToString(), Path.GetExtension(fileName));
            }

            return result;
        }

        #endregion
    }
}
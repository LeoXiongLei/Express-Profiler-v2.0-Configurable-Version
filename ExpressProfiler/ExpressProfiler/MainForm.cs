//sample application for demonstrating Sql Server Profiling
//writen by Locky, 2009.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

using System.Diagnostics;
using System.Configuration;


namespace ExpressProfiler
{
    public partial class MainForm : Form
    {
        internal const string versionString = "Express Profiler v2.0-Configurable Version";

        private class PerfInfo
        {
            internal int m_count;
            internal readonly DateTime m_date = DateTime.Now;
        }

        public class PerfColumn
        {
            public string  Caption;
            public int Column;
            public int Width;
            public string Format;
            public HorizontalAlignment Alignment = HorizontalAlignment.Left;
        }

        private enum ProfilingStateEnum { psStopped, psProfiling, psPaused }
        private RawTraceReader m_Rdr;

        private readonly  YukonLexer m_Lex = new YukonLexer();
        private SqlConnection m_Conn;
        private readonly SqlCommand m_Cmd = new SqlCommand();
        private Thread m_Thr;
        private bool m_NeedStop = true;
        private ProfilingStateEnum m_ProfilingState ;
        private int m_EventCount;
        private readonly ProfilerEvent m_EventStarted = new ProfilerEvent();
        private readonly ProfilerEvent m_EventStopped = new ProfilerEvent();
        private readonly ProfilerEvent m_EventPaused = new ProfilerEvent();
        internal readonly List<ListViewItem> m_Cached = new List<ListViewItem>(1024);
        private readonly Dictionary<string,ListViewItem> m_itembysql = new Dictionary<string, ListViewItem>();
        
        private string m_servername = "";
        private string m_username = "";
        private string m_userpassword = "";
        private string m_userHostName = "";
        private string m_userDataBaseName = "";

        internal int lastpos = -1;
        internal string lastpattern = "";
        private ListViewNF lvEvents;
        Queue<ProfilerEvent> m_events = new Queue<ProfilerEvent>(10);
        private bool m_autostart;
        private bool dontUpdateSource;
        private Exception m_profilerexception;
        private readonly Queue<PerfInfo> m_perf = new Queue<PerfInfo>();
        private PerfInfo m_first, m_prev;
        internal TraceProperties.TraceSettings m_currentsettings;
        private readonly List<PerfColumn> m_columns = new List<PerfColumn>();
        internal bool matchCase = false;
        internal bool wholeWord = false;

        string strServerName = @"10.5.10.76\sql2008r2";
        string strUserName = @"team_cp";
        string strPassWord = @"95938";
        //string strServerName = "";
        //string strUserName = "";
        //string strPassWord = "";
        public MainForm()
        {
            InitializeComponent();
            tbStart.DefaultItem = tbRun;
            Text = versionString;
            edPassword.TextBox.PasswordChar = '*';
            m_servername = strServerName;// Properties.Settings.Default.ServerName;//ServerName
            m_username = strUserName;// Properties.Settings.Default.UserName;//UserName
            //
            m_userpassword = strPassWord;// Properties.Settings.Default.PassWord;//PassWord
            //
            //
            //m_servername = GetAppConfig("ServerName");
            //m_username = GetAppConfig("UserName");
            //m_userpassword = GetAppConfig("PassWord");
            //
            //Type type; 
            //string strtype = type.GUID.ToString("B");
            //string filename = System.Windows.Forms.Application.ExecutablePath + ".config";
            //XmlDocument doc = new XmlDocument();
            //doc.Load(filename); //加载配置文件

            //GetXmlDataInfo.AccountConfiguration accountConfig = (GetXmlDataInfo.AccountConfiguration)ConfigurationManager.GetSection("AccountConfiguration");
            //AccountConfiguration accountConfig = (AccountConfiguration)ConfigurationManager.GetSection("AccountConfiguration");
            //Console.WriteLine(accountConfig.Users.UserName);
            //
            m_currentsettings = GetDefaultSettings();
            ParseCommandLine();
            InitLV();
            edServer.Text = m_servername;
            edUser.Text = m_username;
            edPassword.Text = m_userpassword;
            edDatabaseName.Text = m_userDataBaseName;
            edHostName.Text = m_userHostName;

            tbAuth.SelectedIndex = String.IsNullOrEmpty(m_username)?0:1;
            if(m_autostart) RunProfiling(false);
            UpdateButtons();
            //var conn = Mysoft.CgZtbMap.Common.Data.MyDataBaseManager.GetDataBase().ConnectionString;
            //server=10.5.10.76\sql2008r2;uid=team_cp;pwd=95938;database=CzStandard_dotnet_erp306
            
            //加载数据配置HostUrl标识
            LoadHostCombox();
            cbbSPID.Items.Add("");//加载空进程(查询所有SPID)

            //默认加载第一个数据库信息节点
            LoadHostDefault();
        }

        public void LoadHostDefault()
        {
            GetXmlDataInfo.ConfigUserInfo configUser = GetXmlDataInfo.GetAppXmlDom_Default();
            if (configUser != null)
            {
                m_servername = configUser.ServerName;
                m_username = configUser.UserName;
                m_userpassword = configUser.PassWord;
                m_userHostName = configUser.HostName;
                m_userDataBaseName = configUser.DataBaseName;

                edServer.Text = m_servername;
                edUser.Text = m_username;
                edPassword.Text = m_userpassword;
                edDatabaseName.Text = m_userDataBaseName;
                edHostName.Text = m_userHostName;

                cbbHost.Text = configUser.HostUrl;
            }
            configUser = null;
        }

        //加载站点域名
        public void LoadHostCombox()
        {
            string strHostUrl, strServerName, strUserName, strPassWord, strDataBaseName, strHostName = "";
            string filename = System.Windows.Forms.Application.ExecutablePath + ".config";
            XmlDocument doc = new XmlDocument();
            doc.Load(filename); //加载配置文件

            XmlNodeList nodeList = doc.SelectNodes(string.Format("//configuration/ConfigUserInfo/users"));

            if (nodeList != null)
            {
                for (int i = 0;i<nodeList.Count ;i++ )
                {
                    XmlNode node = nodeList[i];
                    XmlElement childElement = (XmlElement)node;
                    strHostUrl = childElement.GetAttribute("HostUrl");
                    strServerName = childElement.GetAttribute("ServerName");
                    strUserName = childElement.GetAttribute("UserName");
                    strPassWord = childElement.GetAttribute("PassWord");
                    strDataBaseName = childElement.GetAttribute("DataBaseName");
                    strHostName = childElement.GetAttribute("HostName");

                    cbbHost.Items.Add(strHostUrl);   
                }
            }
            else{}
        }
        //public class AccountConfiguration : ConfigurationSection
        //{
        //    [ConfigurationProperty("users", IsRequired = true)]
        //    public AccountSectionElement Users
        //    {
        //        get { return (AccountSectionElement)this["users"]; }
        //    }
        //}

        //public class AccountSectionElement : ConfigurationElement
        //{
        //    [ConfigurationProperty("username", IsRequired = true)]
        //    public string UserName
        //    {
        //        get { return this["username"].ToString(); }
        //        set { this["username"] = value; }
        //    }

        //    [ConfigurationProperty("password", IsRequired = true)]
        //    public string Password
        //    {
        //        get { return this["password"].ToString(); }
        //        set { this["password"] = value; }
        //    }
        //}

        /// <summary>
        ///  读取appStrings配置节， 返回＊.exe.config文件中appSettings配置节的value项
        /// </summary>
        /// <param name="strKey"></param>
        /// <returns></returns>
        public static string GetAppConfig(string strKey)
        {
            foreach (string key in System.Configuration.ConfigurationManager.AppSettings)
            {
                if (key == strKey)
                {
                    return System.Configuration.ConfigurationManager.AppSettings[strKey];
                }
            }
            return null;
        }

        private TraceProperties.TraceSettings GetDefaultSettings()
        {
            try
            {
                XmlSerializer x = new XmlSerializer(typeof(TraceProperties.TraceSettings));
                using (StringReader sr = new StringReader(Properties.Settings.Default.TraceSettings))
                {
                    return (TraceProperties.TraceSettings)x.Deserialize(sr);   
                }
            }
            catch (Exception e)
            {
                //e.Message
            }
            return TraceProperties.TraceSettings.GetDefaultSettings();
        }


        private void ParseCommandLine()
        {
            string[] args = Environment.GetCommandLineArgs();
            int i = 1;
            while (i < args.Length)
            {
                string ep = i + 1 < args.Length ? args[i + 1] : "";
                switch (args[i].ToLower())
                {
                    case "-s":
                    case "-server":
                        m_servername = ep;
                        i++;
                        break;
                    case "-u":
                    case "-user":
                        m_username = ep;
                        i++;
                        break;
                    case "-p":
                    case "-password":
                        m_userpassword = ep;
                        i++;
                        break;
                    case "-m":
                    case "-maxevents":
                        int m;
                        if (!Int32.TryParse(ep, out m)) m = 1000;
                            m_currentsettings.Filters.MaximumEventCount = m;
                        break;
                    case "-d":
                    case "-duration":
                        int d;
                        if (Int32.TryParse(ep, out d))
                        {
                            m_currentsettings.Filters.DurationFilterCondition = TraceProperties.IntFilterCondition.GreaterThan;
                            m_currentsettings.Filters.Duration = d;
                        }

                        break;
                    case "-start":
                        m_autostart = true;
                        break;
                    case "-batchcompleted":
                        m_currentsettings.EventsColumns.BatchCompleted = true;
                        break;
                    case "-batchstarting":
                        m_currentsettings.EventsColumns.BatchStarting = true;
                        break;
                    case "-existingconnection":
                        m_currentsettings.EventsColumns.ExistingConnection = true;
                        break;
                    case "-loginlogout":
                        m_currentsettings.EventsColumns.LoginLogout = true;
                        break;
                    case "-rpccompleted":
                        m_currentsettings.EventsColumns.RPCCompleted = true;
                        break;
                    case "-rpcstarting":
                        m_currentsettings.EventsColumns.RPCStarting = true;
                        break;
                    case "-spstmtcompleted":
                        m_currentsettings.EventsColumns.SPStmtCompleted = true;
                        break;
                    case "-spstmtstarting":
                        m_currentsettings.EventsColumns.SPStmtStarting = true;
                        break;
                       
                }
                i++;
            }

            if (m_servername.Length == 0)
            {
                m_servername = @".\sqlexpress";
            }


        }
        private bool ValidStart()
        { 
            
            //必须输入服务器Server,用户名User,密码Password
            if (tbAuth.SelectedIndex == 0)
            {
                //WINDOWS模式
                if (edServer.Text == "")
                {
                    MessageBox.Show("请输入Server");
                    return false;
                }
            }
            else {
                //SQL模式
                if (edServer.Text == "" || edUser.Text == "" || edPassword.Text == "")
                {
                    MessageBox.Show("请输入Server,User,Password");
                    return false;
                }
            }
            
            return true;
        }
        private void tbStart_Click(object sender, EventArgs e)
        {
            if (!ValidStart())
            {
                return;
            }
            if (!TraceProperties.AtLeastOneEventSelected(m_currentsettings))
            {
                MessageBox.Show("You should select at least 1 event", "Starting trace", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RunProfiling(true);
            }
            {
                RunProfiling(false);
            }
        }

        private void UpdateButtons()
        {
            tbStart.Enabled = m_ProfilingState==ProfilingStateEnum.psStopped||m_ProfilingState==ProfilingStateEnum.psPaused;
            tbRun.Enabled = tbStart.Enabled;
            mnRun.Enabled = tbRun.Enabled;
            tbRunWithFilters.Enabled = ProfilingStateEnum.psStopped==m_ProfilingState;
            mnRunWithFilters.Enabled = tbRunWithFilters.Enabled;
            startTraceToolStripMenuItem.Enabled = tbStart.Enabled;
            tbStop.Enabled = m_ProfilingState==ProfilingStateEnum.psPaused||m_ProfilingState==ProfilingStateEnum.psProfiling;
            stopTraceToolStripMenuItem.Enabled = tbStop.Enabled;
            tbPause.Enabled = m_ProfilingState == ProfilingStateEnum.psProfiling;
            pauseTraceToolStripMenuItem.Enabled = tbPause.Enabled;
            timer1.Enabled = m_ProfilingState == ProfilingStateEnum.psProfiling;
            edServer.Enabled = m_ProfilingState == ProfilingStateEnum.psStopped;
            tbAuth.Enabled = m_ProfilingState == ProfilingStateEnum.psStopped;
            edUser.Enabled = edServer.Enabled&&(tbAuth.SelectedIndex==1);
            edPassword.Enabled = edServer.Enabled && (tbAuth.SelectedIndex == 1);
        }


        private void InitLV()
        {
            lvEvents = new ListViewNF
                           {
                               Dock = DockStyle.Fill,
                               Location = new System.Drawing.Point(0, 0),
                               Name = "lvEvents",
                               Size = new System.Drawing.Size(979, 297),
                               TabIndex = 0,
                               VirtualMode = true,
                               UseCompatibleStateImageBehavior = false,
                               BorderStyle = BorderStyle.None,
                               FullRowSelect = true,
                               View = View.Details,
                               GridLines = true,
                               HideSelection = false,
                               MultiSelect = true,
                               AllowColumnReorder = false
                           };
            lvEvents.RetrieveVirtualItem += lvEvents_RetrieveVirtualItem;
            lvEvents.KeyDown += lvEvents_KeyDown;
            lvEvents.ItemSelectionChanged += listView1_ItemSelectionChanged_1;
            lvEvents.ColumnClick += lvEvents_ColumnClick;
            lvEvents.SelectedIndexChanged += lvEvents_SelectedIndexChanged;
            lvEvents.VirtualItemsSelectionRangeChanged += LvEventsOnVirtualItemsSelectionRangeChanged;
            lvEvents.ContextMenuStrip = contextMenuStrip1;
            splitContainer1.Panel1.Controls.Add(lvEvents);
            InitColumns();
            InitGridColumns();
            
        }

        private void InitColumns()
        {
            m_columns.Clear();
            m_columns.Add(new PerfColumn{ Caption = "Event Class", Column = ProfilerEventColumns.EventClass,Width = 122});
            m_columns.Add(new PerfColumn { Caption = "Text Data", Column = ProfilerEventColumns.TextData, Width = 255});
            m_columns.Add(new PerfColumn { Caption = "Login Name", Column = ProfilerEventColumns.LoginName, Width = 79 });
            m_columns.Add(new PerfColumn { Caption = "CPU", Column = ProfilerEventColumns.CPU, Width = 82, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            m_columns.Add(new PerfColumn { Caption = "Reads", Column = ProfilerEventColumns.Reads, Width = 78, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            m_columns.Add(new PerfColumn { Caption = "Writes", Column = ProfilerEventColumns.Writes, Width = 78, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            m_columns.Add(new PerfColumn { Caption = "Duration, ms", Column = ProfilerEventColumns.Duration, Width = 82, Alignment = HorizontalAlignment.Right, Format = "#,0" });
            m_columns.Add(new PerfColumn { Caption = "SPID", Column = ProfilerEventColumns.SPID, Width = 50, Alignment = HorizontalAlignment.Right });

            if (m_currentsettings.EventsColumns.StartTime) m_columns.Add(new PerfColumn { Caption = "Start time", Column = ProfilerEventColumns.StartTime, Width = 140, Format = "yyyy-MM-dd hh:mm:ss.ffff" });
            if (m_currentsettings.EventsColumns.EndTime) m_columns.Add(new PerfColumn { Caption = "End time", Column = ProfilerEventColumns.EndTime, Width = 140, Format = "yyyy-MM-dd hh:mm:ss.ffff" });
            if (m_currentsettings.EventsColumns.DatabaseName) m_columns.Add(new PerfColumn { Caption = "DatabaseName", Column = ProfilerEventColumns.DatabaseName, Width = 70 });
            if (m_currentsettings.EventsColumns.ObjectName) m_columns.Add(new PerfColumn { Caption = "Object name", Column = ProfilerEventColumns.ObjectName, Width = 70 });
            if (m_currentsettings.EventsColumns.ApplicationName) m_columns.Add(new PerfColumn { Caption = "Application name", Column = ProfilerEventColumns.ApplicationName, Width = 70 });
            //if (m_currentsettings.EventsColumns.HostName) m_columns.Add(new PerfColumn { Caption = "HostName", Column = ProfilerEventColumns.HostName, Width = 70 });
            m_columns.Add(new PerfColumn { Caption = "HostName", Column = ProfilerEventColumns.HostName, Width = 70 });
            m_columns.Add(new PerfColumn { Caption = "#", Column = -1, Width = 53, Alignment = HorizontalAlignment.Right});
        }

        private void InitGridColumns()
        {
            InitColumns();
            lvEvents.BeginUpdate();
            try
            {
                lvEvents.Columns.Clear();
                foreach (PerfColumn pc in m_columns)
                {
                    var l = lvEvents.Columns.Add(pc.Caption, pc.Width);
                    l.TextAlign = pc.Alignment;//lvEvents.SelectedItems[0].SelectedItems[0] .Items[0].SubItems[0]
                }
            }
            finally
            {
                lvEvents.EndUpdate();
            }
        }

        private void LvEventsOnVirtualItemsSelectionRangeChanged(object sender, ListViewVirtualItemsSelectionRangeChangedEventArgs listViewVirtualItemsSelectionRangeChangedEventArgs)
        {
            UpdateSourceBox();
        }

        void lvEvents_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSourceBox();
        }

        void lvEvents_ColumnClick(object sender, ColumnClickEventArgs e)
        {
        }


        private string GetEventCaption(ProfilerEvent evt)
        {
            if (evt == m_EventStarted)
            {
                return "Trace started";
            }
            if (evt == m_EventPaused)
            {
                return "Trace paused";
            }
            if (evt == m_EventStopped)
            {
                return "Trace stopped";
            }
            return ProfilerEvents.Names[evt.EventClass];
        }

        private string GetFormattedValue(ProfilerEvent evt,int column,string format)
        {
            return ProfilerEventColumns.Duration == column ? (evt.Duration / 1000).ToString(format) : evt.GetFormattedData(column,format);
        }

        private void NewEventArrived(ProfilerEvent evt,bool last)
        {
            {
                ListViewItem current = (lvEvents.SelectedIndices.Count > 0) ? m_Cached[lvEvents.SelectedIndices[0]] : null;
                m_EventCount++;
                string caption = GetEventCaption(evt);
                ListViewItem lvi = new ListViewItem(caption);
                string []items = new string[m_columns.Count];
                for (int i = 1; i < m_columns.Count;i++ )
                {
                    PerfColumn pc = m_columns[i];
                    items[i - 1] = pc.Column == -1 ? m_EventCount.ToString("#,0") : GetFormattedValue(evt,pc.Column, pc.Format) ?? "";
                }
                lvi.SubItems.AddRange(items);
                lvi.Tag = evt;
                m_Cached.Add(lvi);
                if (last)
                {
                    lvEvents.VirtualListSize = m_Cached.Count;
                    lvEvents.SelectedIndices.Clear();
                    FocusLVI(tbScroll.Checked ? lvEvents.Items[m_Cached.Count - 1] : current, tbScroll.Checked);
                    lvEvents.Invalidate(lvi.Bounds);
                }
            }
        }

        internal void FocusLVI(ListViewItem lvi,bool ensure)
        {
            if (null != lvi)
            {
                lvi.Focused = true;
                lvi.Selected = true;
                listView1_ItemSelectionChanged_1(lvEvents, null);
                if (ensure)
                {
                    lvEvents.EnsureVisible(lvEvents.Items.IndexOf(lvi));
                }
            }
        }
        //是否启动固定语句执行后停止
        public static bool BIsStepExec = false;
        //固定执行语句数量
        public static int StepCounts = 0;
        //固定执行语句循环
        public static int StepLoop = 0;
        //用于定时回收(回收耗费性能)
        //public static int ClearTimeCount = 0;
        private void ProfilerThread(Object state)
        {
            MyInvoke mi = new MyInvoke(UpdateForm);
            MyInvoke2 mi2 = new MyInvoke2(AfterStepExec);
            try
            {
                while (!m_NeedStop && m_Rdr.TraceIsActive)
                {
                    ProfilerEvent evt = m_Rdr.Next();
                    if (evt != null)
                    {
                        //lbSPID.Text=evt.SPID.ToString();
                        this.BeginInvoke(mi, new Object[] { evt.SPID.ToString(), "test" });
                        lock (this)
                        {
                            m_events.Enqueue(evt);
                        }

                        //回收耗费性能，不建议使用。
                        //定时清理回收
                        //ClearTimeCount++;
                        //if (ClearTimeCount > 100)
                        //{
                        //    ClearMemory();
                        //    ClearTimeCount = 0;
                        //}

                        if (BIsStepExec)
                        {
                            StepLoop++;
                            if (StepLoop > StepCounts)
                            {
                                //置为初始
                                BIsStepExec = false;
                                StepCounts = 0;
                                StepLoop = 0;

                                this.BeginInvoke(mi2, new Object[] { "param1", "param2" });
                                return;
                            }
                        }

                    }
                }
            }
            catch (Exception e)
            {
                lock (this)
                {
                    if (!m_NeedStop && m_Rdr.TraceIsActive)
                    {
                        m_profilerexception = e;
                    }
                }
            }

            mi = null;
            //ClearMemory();
        }
        public delegate void MyInvoke(string str1, string str2);
        public void UpdateForm(string param1, string parm2)
        {
            //this.lbSPID.Text = param1;
            LoadSPIDCombox(param1);
            //内存显示
            GetProcessUsedMemory();
        }

        public delegate void MyInvoke2(string str1, string str2);
        public void AfterStepExec(string param1, string parm2)
        {
            StopProfiling();
        }

        //加载SQL会话
        public void LoadSPIDCombox(string strSPID)
        {
            if (!cbbSPID.Items.Contains(strSPID))
            {
                cbbSPID.Items.Add(strSPID);
            }
        }

        private  SqlConnection GetConnection()
        {
            return new SqlConnection
                       {
                           ConnectionString =
                           tbAuth.SelectedIndex==0?String.Format(@"Data Source = {0}; Initial Catalog = master; Integrated Security=SSPI;Application Name=Express Profiler",edServer.Text)
                           : String.Format(@"Data Source={0};Initial Catalog=master;User Id={1};Password='{2}';;Application Name=Express Profiler", edServer.Text, edUser.Text, edPassword.Text)
                       };
            //MyDataBaseManager.GetDataBase().ConnectionString:server=10.5.10.76\sql2008r2;uid=team_cp;pwd=95938;database=CzStandard_dotnet_erp306

        }

        private void StartProfiling()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                m_perf.Clear();
                m_first = null;
                m_prev = null;
                if (m_ProfilingState == ProfilingStateEnum.psPaused)
                {
                    ResumeProfiling();
                    return;
                }
                if (m_Conn != null && m_Conn.State == ConnectionState.Open)
                {
                    m_Conn.Close();
                }
                InitGridColumns();
                m_EventCount = 0;
                m_Conn = GetConnection();
                m_Conn.Open();
                m_Rdr = new RawTraceReader(m_Conn);

                m_Rdr.CreateTrace();
                if (true)
                {
                    if (m_currentsettings.EventsColumns.LoginLogout)
                    {
                        m_Rdr.SetEvent(ProfilerEvents.SecurityAudit.AuditLogin,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime
                            );
                        m_Rdr.SetEvent(ProfilerEvents.SecurityAudit.AuditLogout,
                                       ProfilerEventColumns.CPU,
                                       ProfilerEventColumns.Reads,
                                       ProfilerEventColumns.Writes,
                                       ProfilerEventColumns.Duration,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.ApplicationName
                            );
                    }

                    if (m_currentsettings.EventsColumns.ExistingConnection)
                    {
                        m_Rdr.SetEvent(ProfilerEvents.Sessions.ExistingConnection,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.ApplicationName
                            );
                    }
                    if (m_currentsettings.EventsColumns.BatchCompleted)
                    {
                        m_Rdr.SetEvent(ProfilerEvents.TSQL.SQLBatchCompleted,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.CPU,
                                       ProfilerEventColumns.Reads,
                                       ProfilerEventColumns.Writes,
                                       ProfilerEventColumns.Duration,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.DatabaseName,
                                       ProfilerEventColumns.ApplicationName
                                       , ProfilerEventColumns.HostName      //主机名
                            );
                    }
                    if (m_currentsettings.EventsColumns.BatchStarting)
                    {
                        m_Rdr.SetEvent(ProfilerEvents.TSQL.SQLBatchStarting,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.DatabaseName,
                                       ProfilerEventColumns.ApplicationName
                            );
                    }
                    if (m_currentsettings.EventsColumns.RPCStarting)
                    {
                        m_Rdr.SetEvent(ProfilerEvents.StoredProcedures.RPCStarting,
                                       ProfilerEventColumns.TextData,
                                       ProfilerEventColumns.LoginName,
                                       ProfilerEventColumns.SPID,
                                       ProfilerEventColumns.StartTime,
                                       ProfilerEventColumns.EndTime,
                                       ProfilerEventColumns.DatabaseName,
                                       ProfilerEventColumns.ObjectName,
                                       ProfilerEventColumns.ApplicationName
                            );
                    }

                }
                if (m_currentsettings.EventsColumns.RPCCompleted)
                {
                    m_Rdr.SetEvent(ProfilerEvents.StoredProcedures.RPCCompleted,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ObjectName
                                   , ProfilerEventColumns.ApplicationName
                                   , ProfilerEventColumns.HostName      //主机名
                        );
                }
                if (m_currentsettings.EventsColumns.SPStmtCompleted)
                {
                    m_Rdr.SetEvent(ProfilerEvents.StoredProcedures.SPStmtCompleted,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ObjectName
                                   , ProfilerEventColumns.ObjectID
                                   , ProfilerEventColumns.ApplicationName
                        );
                }
                if (m_currentsettings.EventsColumns.SPStmtStarting)
                {
                    m_Rdr.SetEvent(ProfilerEvents.StoredProcedures.SPStmtStarting,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ObjectName
                                   , ProfilerEventColumns.ObjectID
                                   , ProfilerEventColumns.ApplicationName
                        );
                }
                if (m_currentsettings.EventsColumns.UserErrorMessage)
                {
                    m_Rdr.SetEvent(ProfilerEvents.ErrorsAndWarnings.UserErrorMessage,
                                   ProfilerEventColumns.TextData,
                                   ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU,
                                   ProfilerEventColumns.SPID,
                                   ProfilerEventColumns.StartTime,
                                   ProfilerEventColumns.DatabaseName,
                                   ProfilerEventColumns.ApplicationName
                        );
                }
                if (m_currentsettings.EventsColumns.BlockedProcessPeport)
                {
                    m_Rdr.SetEvent(ProfilerEvents.ErrorsAndWarnings.Blockedprocessreport,
                                   ProfilerEventColumns.TextData,
                                   ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU,
                                   ProfilerEventColumns.SPID,
                                   ProfilerEventColumns.StartTime,
                                   ProfilerEventColumns.DatabaseName,
                                   ProfilerEventColumns.ApplicationName
                        );

                }

                if (m_currentsettings.EventsColumns.SQLStmtStarting)
                {
                    m_Rdr.SetEvent(ProfilerEvents.TSQL.SQLStmtStarting,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ApplicationName
                        );
                }
                if (m_currentsettings.EventsColumns.SQLStmtCompleted)
                {
                    m_Rdr.SetEvent(ProfilerEvents.TSQL.SQLStmtCompleted,
                                   ProfilerEventColumns.TextData, ProfilerEventColumns.LoginName,
                                   ProfilerEventColumns.CPU, ProfilerEventColumns.Reads,
                                   ProfilerEventColumns.Writes, ProfilerEventColumns.Duration,
                                   ProfilerEventColumns.SPID
                                   , ProfilerEventColumns.StartTime, ProfilerEventColumns.EndTime
                                   , ProfilerEventColumns.DatabaseName
                                   , ProfilerEventColumns.ApplicationName
                        );
                }

                if (null != m_currentsettings.Filters.Duration)
                {
                    SetIntFilter(m_currentsettings.Filters.Duration*1000,
                                 m_currentsettings.Filters.DurationFilterCondition, ProfilerEventColumns.Duration);
                }
                SetIntFilter(m_currentsettings.Filters.Reads, m_currentsettings.Filters.ReadsFilterCondition,ProfilerEventColumns.Reads);
                SetIntFilter(m_currentsettings.Filters.Writes, m_currentsettings.Filters.WritesFilterCondition,ProfilerEventColumns.Writes);
                SetIntFilter(m_currentsettings.Filters.CPU, m_currentsettings.Filters.CpuFilterCondition,ProfilerEventColumns.CPU);
                SetIntFilter(m_currentsettings.Filters.SPID, m_currentsettings.Filters.SPIDFilterCondition, ProfilerEventColumns.SPID);

                SetStringFilter(m_currentsettings.Filters.LoginName, m_currentsettings.Filters.LoginNameFilterCondition,ProfilerEventColumns.LoginName);
                //SetStringFilter(m_currentsettings.Filters.DatabaseName,m_currentsettings.Filters.DatabaseNameFilterCondition, ProfilerEventColumns.DatabaseName);
                SetStringFilter(m_currentsettings.Filters.TextData, m_currentsettings.Filters.TextDataFilterCondition,ProfilerEventColumns.TextData);
                SetStringFilter(m_currentsettings.Filters.ApplicationName, m_currentsettings.Filters.ApplicationNameFilterCondition, ProfilerEventColumns.ApplicationName);

                //, ProfilerEventColumns.HostName      //主机名
                //SetStringFilter("xionglei", m_currentsettings.Filters.DatabaseNameFilterCondition, ProfilerEventColumns.HostName);
                //SetStringFilter(m_currentsettings.Filters.HostName, m_currentsettings.Filters.HostNameFilterCondition, ProfilerEventColumns.HostName);
                if (!string.IsNullOrEmpty(m_userHostName))
                {
                    SetStringFilter(m_userHostName, TraceProperties.StringFilterCondition.Like, ProfilerEventColumns.HostName);
                }
                if (!string.IsNullOrEmpty(m_userDataBaseName))
                {
                    SetStringFilter(m_userDataBaseName, TraceProperties.StringFilterCondition.Like, ProfilerEventColumns.DatabaseName);
                }

                m_Cmd.Connection = m_Conn;
                m_Cmd.CommandTimeout = 0;
                m_Rdr.SetFilter(ProfilerEventColumns.ApplicationName, LogicalOperators.AND, ComparisonOperators.NotLike,
                                "Express Profiler");
                m_Cached.Clear();
                m_events.Clear();
                m_itembysql.Clear();
                lvEvents.VirtualListSize = 0;
                StartProfilerThread();
                m_servername = edServer.Text;
                m_username = edUser.Text;
                //Properties.Settings.Default.ServerName = m_servername;
                //Properties.Settings.Default.UserName = tbAuth.SelectedIndex == 0 ? "" : m_username;
                Properties.Settings.Default.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UpdateButtons();
                Cursor = Cursors.Default;
            }
        }


        private void SetIntFilter(int? value, TraceProperties.IntFilterCondition condition, int column)
        {
            int[] com = new[] { ComparisonOperators.Equal, ComparisonOperators.NotEqual, ComparisonOperators.GreaterThan, ComparisonOperators.LessThan};
            if ((null != value))
            {
                long? v = value;
                m_Rdr.SetFilter(column, LogicalOperators.AND, com[(int)condition], v);
            }
        }

        private void SetStringFilter(string value,TraceProperties.StringFilterCondition condition,int column)
        {
            if (!String.IsNullOrEmpty(value))
            {
                m_Rdr.SetFilter(column, LogicalOperators.AND
                    , condition == TraceProperties.StringFilterCondition.Like ? ComparisonOperators.Like : ComparisonOperators.NotLike
                    , value
                    );
            }

        }

        private void StartProfilerThread()
        { 
            if(m_Rdr!=null)
            {
                m_Rdr.Close();
            }
            m_Rdr.StartTrace();
            m_Thr = new Thread(ProfilerThread) {IsBackground = true, Priority = ThreadPriority.Lowest};
            m_NeedStop = false;
            m_ProfilingState = ProfilingStateEnum.psProfiling;
            NewEventArrived(m_EventStarted,true);
            m_Thr.Start();
        }

        private void ResumeProfiling()
        {
            StartProfilerThread();
            UpdateButtons();
        }

        private void tbStop_Click(object sender, EventArgs e)
        {
            StopProfiling();
            ClearMemory();
        }

        private void StopProfiling()
        {
            tbStop.Enabled = false;
            cbbHost.Enabled = true;
            edDatabaseName.Enabled = true;
            edHostName.Enabled = true;
            btnSaveConfig.Enabled = true;
            //进程只能在运行时使用
            cbbSPID.Enabled = false;
            btnCurrentSPID.Enabled = false;

            using (SqlConnection cn = GetConnection())
            {
                cn.Open();
                m_Rdr.StopTrace(cn);
                m_Rdr.CloseTrace(cn);
                cn.Close();
            }
            m_NeedStop = true;
            if (m_Thr.IsAlive)
            {
                m_Thr.Abort();
            }
            m_Thr.Join();
            m_ProfilingState = ProfilingStateEnum.psStopped;
            NewEventArrived(m_EventStopped,true);
            UpdateButtons();
        }

        private void listView1_ItemSelectionChanged_1(object sender, ListViewItemSelectionChangedEventArgs e)
        {

            UpdateSourceBox();
        }

        private void UpdateSourceBox()
        {
            if (dontUpdateSource) return;
            StringBuilder sb = new StringBuilder();
            foreach (int i in lvEvents.SelectedIndices)
            {
                ListViewItem lv = m_Cached[i];
                if (lv.SubItems[1].Text != "")
                {
                    sb.AppendFormat("{0}\r\ngo\r\n", lv.SubItems[1].Text);
                }
            }
            m_Lex.FillRichEdit(reTextData, sb.ToString());
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(m_ProfilingState==ProfilingStateEnum.psPaused||m_ProfilingState==ProfilingStateEnum.psProfiling)
            {
                if (MessageBox.Show("There are traces still running. Are you sure you want to close the application?","ExpressProfiler",MessageBoxButtons.YesNo,MessageBoxIcon.Question
                    ,MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    StopProfiling();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void lvEvents_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            e.Item = m_Cached[e.ItemIndex];
        }

        private void tbPause_Click(object sender, EventArgs e)
        {
            PauseProfiling();
        }

        private void PauseProfiling()
        {
            using (SqlConnection cn = GetConnection())
            {
                cn.Open();
                m_Rdr.StopTrace(cn);
                cn.Close();
            }
            m_ProfilingState = ProfilingStateEnum.psPaused;
            NewEventArrived(m_EventPaused,true);
            UpdateButtons();
        }


        internal void SelectAllEvents(bool select)
        {
            lock (m_Cached)
            {
                lvEvents.BeginUpdate();
                dontUpdateSource = true;
                try
                {

                    foreach (ListViewItem lv in m_Cached)
                    {
                        lv.Selected = select;
                    }
                }
                finally
                {
                    dontUpdateSource = false;
                    UpdateSourceBox();
                    lvEvents.EndUpdate();
                }
            }
        }

        private void lvEvents_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            Queue<ProfilerEvent> saved;
            Exception exc;
            lock (this)
            {
                saved = m_events;
                m_events = new Queue<ProfilerEvent>(10);
                exc = m_profilerexception;
                m_profilerexception = null;
            }
            if (null != exc)
            {
                using (ThreadExceptionDialog dlg = new ThreadExceptionDialog(exc))
                {
                    dlg.ShowDialog();
                }
            }
            lock (m_Cached)
            {
                while (0 != saved.Count)
                {
                    NewEventArrived(saved.Dequeue(), 0 == saved.Count);
                }
                if (m_Cached.Count > m_currentsettings.Filters.MaximumEventCount)
                {
                    while (m_Cached.Count > m_currentsettings.Filters.MaximumEventCount)
                    {
                        m_Cached.RemoveAt(0);
                    }
                    lvEvents.VirtualListSize = m_Cached.Count;
                    lvEvents.Invalidate();
                }

                if ((null == m_prev) || (DateTime.Now.Subtract(m_prev.m_date).TotalSeconds >= 1))
                {
                    PerfInfo curr = new PerfInfo {m_count = m_EventCount};
                    if (m_perf.Count >= 60)
                    {
                        m_first = m_perf.Dequeue();
                    }
                    if (null == m_first) m_first = curr;
                    if (null == m_prev) m_prev = curr;

                    DateTime now = DateTime.Now;
                    double d1 = now.Subtract(m_prev.m_date).TotalSeconds;
                    double d2 = now.Subtract(m_first.m_date).TotalSeconds;
                    slEPS.Text = String.Format("{0} / {1} EPS(last/avg for {2} second(s))",
                        (Math.Abs(d1 - 0) > 0.001 ? ((curr.m_count - m_prev.m_count)/d1).ToString("#,0.00") : ""),
                                 (Math.Abs(d2 - 0) > 0.001 ? ((curr.m_count - m_first.m_count) / d2).ToString("#,0.00") : ""), d2 .ToString("0"));

                    m_perf.Enqueue(curr);
                    m_prev = curr;
                }

            }
        }

        private void tbAuth_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateButtons();
        }

        private void ClearTrace()
        {
            lock (lvEvents)
            {
                m_Cached.Clear();
                m_itembysql.Clear();
                lvEvents.VirtualListSize = 0;
                listView1_ItemSelectionChanged_1(lvEvents, null);
                lvEvents.Invalidate();
            }
        }

        private void tbClear_Click(object sender, EventArgs e)
        {
            ClearTrace();
            ClearMemory();
            //清空SPID下拉加载进程
            cbbSPID.Items.Clear();
            cbbSPID.Items.Add("");
        }

        private void NewAttribute(XmlNode node, string name, string value)
        {
            XmlAttribute attr = node.OwnerDocument.CreateAttribute(name);
            attr.Value = value;
            node.Attributes.Append(attr);
        }
        private void NewAttribute(XmlNode node, string name, string value, string namespaceURI)
        {
            XmlAttribute attr = node.OwnerDocument.CreateAttribute("ss", name, namespaceURI);
            attr.Value = value;
            node.Attributes.Append(attr);
        }

        private void copyAllToClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(false);
        }

        private void CopyEventsToClipboard(bool copySelected)
        {
            XmlDocument doc = new XmlDocument();
            XmlNode root = doc.CreateElement("events");
            lock (m_Cached)
            {
                if (copySelected)
                {
                    foreach (int i in lvEvents.SelectedIndices)
                    {
                        CreateEventRow((ProfilerEvent)(m_Cached[i]).Tag, root);
                    }
                }
                else
                {
                    foreach (var i in m_Cached)
                    {
                        CreateEventRow((ProfilerEvent)i.Tag, root);
                    }
                }
            }
            doc.AppendChild(root);
            doc.PreserveWhitespace = true;
            using (StringWriter writer = new StringWriter())
            {
                XmlTextWriter textWriter = new XmlTextWriter(writer) {Formatting = Formatting.Indented};
                doc.Save(textWriter);
                Clipboard.SetText(writer.ToString());
            }
            MessageBox.Show("Event(s) data copied to clipboard", "Information", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void CreateEventRow(ProfilerEvent evt, XmlNode root)
        {
            XmlNode row = root.OwnerDocument.CreateElement("event");
            NewAttribute(row, "EventClass", evt.EventClass.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "CPU", evt.CPU.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "Reads", evt.Reads.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "Writes", evt.Writes.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "Duration", evt.Duration.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "SPID", evt.SPID.ToString(CultureInfo.InvariantCulture));
            NewAttribute(row, "LoginName", evt.LoginName);
            NewAttribute(row, "DatabaseName", evt.DatabaseName);
            NewAttribute(row, "ObjectName", evt.ObjectName);
            NewAttribute(row, "ApplicationName", evt.ApplicationName);
            row.InnerText = evt.TextData;
            root.AppendChild(row);
        }

        private void copySelectedToClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(true);
        }

        private void clearTraceWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearTrace();
        }

        private void extractAllEventsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(false);
        }

        private void extractSelectedEventsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyEventsToClipboard(true);
        }


        private void pauseTraceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PauseProfiling();
        }

        private void stopTraceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StopProfiling();
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DoFind();
        }

        private void DoFind()
        {
            if (m_ProfilingState == ProfilingStateEnum.psProfiling)
            {
                MessageBox.Show("You cannot find when trace is running", "ExpressProfiler", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }
            using (FindForm f = new FindForm(this))
            {
                f.ShowDialog();
            }
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lvEvents.Focused && (m_ProfilingState!=ProfilingStateEnum.psProfiling))
            {
                SelectAllEvents(true);
            }
            else
            if (reTextData.Focused)
            {
                reTextData.SelectAll();
            }
        }

        internal void PerformFind(bool forwards)
        {
            if(String.IsNullOrEmpty(lastpattern)) return;

            if (forwards)
            {
                for (int i = lastpos = lvEvents.Items.IndexOf(lvEvents.FocusedItem) + 1; i < m_Cached.Count; i++)
                {
                    if (FindText(i))
                    {
                        return;
                    }
                }
            }
            else
            {
                for (int i = lastpos = lvEvents.Items.IndexOf(lvEvents.FocusedItem) - 1; i > 0; i--)
                {
                    if (FindText(i))
                    {
                        return;
                    }
                }
            }
            MessageBox.Show(String.Format("Failed to find \"{0}\". Searched to the end of data. ", lastpattern), "ExpressProfiler", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool FindText(int i)
        {
            ListViewItem lvi = m_Cached[i];
            ProfilerEvent evt = (ProfilerEvent) lvi.Tag;
            string pattern = (wholeWord ? "\\b" + lastpattern + "\\b" : lastpattern);
            if (Regex.IsMatch(evt.TextData, pattern, (matchCase ? RegexOptions.None : RegexOptions.IgnoreCase)))
            {
                lvi.Focused = true;
                lastpos = i;
                SelectAllEvents(false);
                FocusLVI(lvi, true);
                return true;
            }

            return false;
        }

        private void findNextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (m_ProfilingState == ProfilingStateEnum.psProfiling)
            {
                MessageBox.Show("You cannot find when trace is running", "ExpressProfiler", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }
            PerformFind(true);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        internal void RunProfiling(bool showfilters)
        {
            if (showfilters)
            {
                TraceProperties.TraceSettings ts = m_currentsettings.GetCopy();
                using (TraceProperties frm = new TraceProperties())
                {
                    frm.SetSettings(ts);
                    if (DialogResult.OK != frm.ShowDialog()) return;
                    m_currentsettings = frm.m_currentsettings.GetCopy();
                }
            }
            //

            
            StartProfiling();

            cbbHost.Enabled = false;
            edDatabaseName.Enabled = false;
            edHostName.Enabled = false;
            btnSaveConfig.Enabled = false;
            //进程只能在运行时使用
            cbbSPID.Enabled = true;
            btnCurrentSPID.Enabled = true;
        }

        private void tbRunWithFilters_Click(object sender, EventArgs e)
        {
            RunProfiling(true);
        }

        private void copyToXlsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyForExcel();
        }

        private void CopyForExcel()
        {

            XmlDocument doc = new XmlDocument();
            XmlProcessingInstruction pi = doc.CreateProcessingInstruction("mso-application", "progid='Excel.Sheet'");
            doc.AppendChild(pi); 
            const string urn = "urn:schemas-microsoft-com:office:spreadsheet";
            XmlNode root = doc.CreateElement("ss","Workbook",urn);
            NewAttribute(root, "xmlns:ss", urn);
            doc.AppendChild(root);

            XmlNode styles = doc.CreateElement("ss","Styles", urn);
            root.AppendChild(styles);
            XmlNode style = doc.CreateElement("ss","Style", urn);
            styles.AppendChild(style);
            NewAttribute(style,"ID","s62",urn);
            XmlNode font = doc.CreateElement("ss","Font",urn);
            style.AppendChild(font);
            NewAttribute(font, "Bold", "1", urn);


            XmlNode worksheet = doc.CreateElement("ss", "Worksheet", urn);
            root.AppendChild(worksheet);
            NewAttribute(worksheet, "Name", "Sql Trace", urn);
            XmlNode table = doc.CreateElement("ss", "Table", urn);
            worksheet.AppendChild(table);
            NewAttribute(table, "ExpandedColumnCount",m_columns.Count.ToString(CultureInfo.InvariantCulture),urn);

            foreach (ColumnHeader lv in lvEvents.Columns)
            {
                XmlNode r = doc.CreateElement("ss","Column", urn);
                NewAttribute(r, "AutoFitWidth","0",urn);
                NewAttribute(r, "Width", lv.Width.ToString(CultureInfo.InvariantCulture), urn);
                table.AppendChild(r);
            }

            XmlNode row = doc.CreateElement("ss","Row", urn);
            table.AppendChild(row);
            foreach (ColumnHeader lv in lvEvents.Columns)
            {
                XmlNode cell = doc.CreateElement("ss","Cell", urn);
                row.AppendChild(cell);
                NewAttribute(cell, "StyleID","s62",urn);
                XmlNode data = doc.CreateElement("ss","Data", urn);
                cell.AppendChild(data);
                NewAttribute(data, "Type","String",urn);
                data.InnerText = lv.Text;
            }

            lock (m_Cached)
            {
                foreach (ListViewItem lvi in m_Cached)
                {
                    row = doc.CreateElement("ss", "Row", urn);
                    table.AppendChild(row);
                    for (int i = 0; i < m_columns.Count; i++)
                    {
                        PerfColumn pc = m_columns[i];
                        if(pc.Column!=-1)
                        {
                        XmlNode cell = doc.CreateElement("ss", "Cell", urn);
                        row.AppendChild(cell);
                        XmlNode data = doc.CreateElement("ss", "Data", urn);
                        cell.AppendChild(data);
                            string dataType;
                            switch (ProfilerEventColumns.ProfilerColumnDataTypes[pc.Column])
                            {
                                    case ProfilerColumnDataType.Int:
                                    case ProfilerColumnDataType.Long:
                                        dataType = "Number";
                                    break;
                                    case ProfilerColumnDataType.DateTime:
                                        dataType = "String";
                                    break;
                                default:
                                        dataType = "String";
                                    break;
                            }
                        if (ProfilerEventColumns.EventClass == pc.Column) dataType = "String";
                        NewAttribute(data, "Type", dataType, urn);
                        if (ProfilerEventColumns.EventClass == pc.Column)
                        {
                            data.InnerText = GetEventCaption(((ProfilerEvent) (lvi.Tag)));
                        }
                        else
                        {
                            data.InnerText = pc.Column == -1
                                                 ? ""
                                                 : GetFormattedValue((ProfilerEvent)(lvi.Tag),pc.Column,ProfilerEventColumns.ProfilerColumnDataTypes[pc.Column]==ProfilerColumnDataType.DateTime?pc.Format:"") ??
                                                   "";
                        }
                            }
                    }
                }
            }
            using (StringWriter writer = new StringWriter())
            {
                XmlTextWriter textWriter = new XmlTextWriter(writer) { Formatting = Formatting.Indented,Namespaces = true};
                doc.Save(textWriter);
                string xml = writer.ToString();
                MemoryStream xmlStream = new MemoryStream();
                xmlStream.Write(System.Text.Encoding.UTF8.GetBytes(xml), 0, xml.Length);
                Clipboard.SetData("XML Spreadsheet", xmlStream);
            }
            MessageBox.Show("Event(s) data copied to clipboard", "Information", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

        }

        private void mnAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show(versionString + "\nhttps://expressprofiler.codeplex.com/\n\nConfigurable Version is developed by XiongLei", "About", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void tbStayOnTop_Click(object sender, EventArgs e)
        {
            SetStayOnTop();
        }

        private void SetStayOnTop()
        {
            tbStayOnTop.Checked = !tbStayOnTop.Checked;
            this.TopMost = tbStayOnTop.Checked;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            SetTransparent();
        }

        private void SetTransparent()
        {
            tbTransparent.Checked = !tbTransparent.Checked;
            this.Opacity = tbTransparent.Checked ? 0.50 : 1;
        }

        private void stayOnTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetStayOnTop();
        }

        private void transparentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetTransparent();
        }

        private void deleteSelectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = lvEvents.SelectedIndices.Count-1; i >= 0; i--)
            {
                m_Cached.RemoveAt(lvEvents.SelectedIndices[i]);
            }
            lvEvents.VirtualListSize = m_Cached.Count;
            lvEvents.SelectedIndices.Clear();
        }

        private void keepSelectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = m_Cached.Count - 1; i >= 0; i--)
            {
                if (!lvEvents.SelectedIndices.Contains(i))
                {
                    m_Cached.RemoveAt(i);
                }
            }
            lvEvents.VirtualListSize = m_Cached.Count;
            lvEvents.SelectedIndices.Clear();
        }

        private void lbSPID_Click(object sender, EventArgs e)
        {
            if (!cbbHost.Enabled)
            {
                //跟踪开启时
                //关闭
                StopProfiling();
            }

            int curSPID;
            if (lbSPID.Text!="SPID" && lbSPID.Text!="")
            {
                curSPID=Convert.ToInt32(lbSPID.Text);
                //m_currentsettings.Filters.SPID = curSPID;
            }
            RunProfiling(false);
            cbbSPID.Items.Clear();
            cbbSPID.Items.Add("");
        }

        private void cbbHost_SelectedIndexChanged(object sender, EventArgs e)
        {
            //切换Host主机域名后，切换加载数据库连接
            //m_servername = configUser.ServerName;
            //m_username = configUser.UserName;
            //m_userpassword = configUser.PassWord;
            try
            {
                string strSelectedText = ((System.Windows.Forms.ComboBox)(sender)).Text;
                if (strSelectedText != "")
                {
                    GetXmlDataInfo.ConfigUserInfo configUser = GetXmlDataInfo.GetAppXmlDom(strSelectedText);
                    if (configUser != null)
                    {
                        m_servername = configUser.ServerName;
                        m_username = configUser.UserName;
                        m_userpassword = configUser.PassWord;
                        m_userHostName = configUser.HostName;
                        m_userDataBaseName = configUser.DataBaseName;

                        edServer.Text = m_servername;
                        edUser.Text = m_username;
                        edPassword.Text = m_userpassword;
                        edDatabaseName.Text = m_userDataBaseName;
                        edHostName.Text = m_userHostName;
                    }
                }
            }
            catch (Exception exp)
            { 
            
            }

        }

        private void edDatabaseName_TextChanged(object sender, EventArgs e)
        {
            //string strText = ((System.Windows.Forms.TextBox)(sender)).Text;
            m_userDataBaseName = edDatabaseName.Text;
        }

        private void edHostName_TextChanged(object sender, EventArgs e)
        {
            //string strText = ((System.Windows.Forms.TextBox)(sender)).Text;
            m_userHostName = edHostName.Text;
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            GetXmlDataInfo.HostUrl = cbbHost.Text;
            GetXmlDataInfo.ServerName = edServer.Text;
            GetXmlDataInfo.UserName = edUser.Text;
            GetXmlDataInfo.PassWord = edPassword.Text;
            GetXmlDataInfo.DataBaseName = edDatabaseName.Text;
            GetXmlDataInfo.HostName = edHostName.Text;

            if (GetXmlDataInfo.UpdateOrCreateAppSetting(cbbHost.Text))
            {
                MessageBox.Show("保存成功！");
            }
            else {
                MessageBox.Show("保存失败！");
            }

        }
        //切换到当前SPID进程
        private void CurrentSPID_Click(object sender, EventArgs e)
        {
            string curSPID = cbbSPID.Text;
            if (curSPID == "")
            {
                m_currentsettings.Filters.SPID = null;
            }
            else
            {
                m_currentsettings.Filters.SPID = Convert.ToInt32(curSPID);
            }

            StopProfiling();
            RunProfiling(false);
            
        }
        //切换到当前SPID进程
        private void btnCurrentSPID_Click(object sender, EventArgs e)
        {
            string curSPID = cbbSPID.Text;
            if (curSPID == "")
            {
                m_currentsettings.Filters.SPID = null;
            }
            else
            {
                m_currentsettings.Filters.SPID = Convert.ToInt32(curSPID);
            }

           
        }

        private void tsBtnTimerStart_Click(object sender, EventArgs e)
        {
            ////关闭
            //StopProfiling();
            //Thread.Sleep(3000);
            ////启动
            //RunProfiling(false);
            
            int nS = 0;

            try
            {
                nS = Convert.ToInt32(tsTxtStep.Text);
                if (nS > 0)
                {
                    StepsExecSQL(nS);
                }
            }
            catch (Exception exp)
            { 
            
            }
            
            
        }
        //固定命令执行
        public void StepsExecSQL(int nSteps)
        {
            if (!cbbHost.Enabled)
            { 
                //跟踪开启时
                //关闭
                StopProfiling();
            }
            //设定时间秒后执行
            int nS = 0;
            try
            {
                nS = Convert.ToInt32(tsTxtStepTime.Text);
                if (nS >= 0)
                {
                    //透明化启动
                    SetTransparent();
                    //this.Visible = false;
                    Thread.Sleep(nS*1000);
                    //透明化关闭
                    SetTransparent();
                    //this.Visible = true;
                }
            }
            catch (Exception exp){
                //return;
            }

            BIsStepExec = true;
            StepCounts = nSteps;
            StepLoop = 0;

            //启动
            RunProfiling(false);

        }
       
        //固定时间执行
        private void tsBtnTimerStart_Sec_Click(object sender, EventArgs e)
        {
            //设定时间秒后执行
            int nS = 0;
            try
            {
                nS = Convert.ToInt32(tsTxtStepTime.Text);
                if (nS >= 0)
                {
                    //透明化启动
                    SetTransparent();
                    //this.Visible = false;
                    Thread.Sleep(nS * 1000);
                    //透明化关闭
                    SetTransparent();
                    //this.Visible = true;
                }
            }
            catch (Exception exp)
            {
                //return;
            }
            //执行固定时间内
            nS = 0;
            try
            {
                nS = Convert.ToInt32(tsTxtTimer.Text);
                if (nS > 0)
                {
                    TimerExecSQL(nS);   
                }
            }
            catch (Exception exp) { return; }

        }

        public static int TimerSecond = 0;
        public static int TimerSecondLoop = 0;
        //设置时间执行
        public void TimerExecSQL(int nSeconds)
        {
            if (!cbbHost.Enabled)
            {
                //跟踪开启时
                //关闭
                StopProfiling();
            }
            //每隔1秒触发
            timer2.Enabled = true;
            timer2.Interval = 1000;
            TimerSecond = nSeconds;
            TimerSecondLoop = 0;

            //启动
            RunProfiling(false);
            
        } 

        private void timer2_Tick(object sender, EventArgs e)
        {
            //每隔1秒触发
            TimerSecondLoop++;

            if (TimerSecondLoop > TimerSecond)
            {
                //关闭
                StopProfiling();
                timer2.Enabled = false;
            }
        }

        #region 内存回收
        
        //public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);
        /// <summary>
        /// 释放内存
        /// </summary>
        public static void ClearMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            //{
            //    //App.SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            //    SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            //}
        }
        #endregion

        private void btnMemoryClear_Click(object sender, EventArgs e)
        {
            ClearMemory();
        }

        // 获得为该进程(程序)分配的内存. 做一个计时器,就可以时时查看程序占用系统资源
        public double GetProcessUsedMemory()
        {
            double usedMemory = 0;

            //usedMemory = Process.GetCurrentProcess().WorkingSet64 / 1024.0 / 1024.0;
            //专用内存
            //usedMemory = Process.GetCurrentProcess().PrivateMemorySize / 1024.0;
            //物理内存
            usedMemory = Process.GetCurrentProcess().WorkingSet64 / 1024.0;

            //Process[] MyProcess = Process.GetProcessesByName("ExpressProfiler.exe");
            

            lblMemory.Text = "内存：" + usedMemory.ToString();
            return usedMemory;

        }       





    }
}

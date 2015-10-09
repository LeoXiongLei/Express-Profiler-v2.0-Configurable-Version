using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Configuration;
using Microsoft.Win32;

namespace ExpressProfiler
{
    class GetXmlDataInfo : ConfigurationSection
    {

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

        [Serializable]
        public class ConfigUserInfo
        {
            public string HostUrl { get; set; }
            public string ServerName { get; set; }
            public string UserName { get; set; }
            public string PassWord { get; set; }
            public string DataBaseName { get; set; }
            public string HostName { get; set; }
        }
        
        public static string HostUrl { get; set; }
        public static string ServerName { get; set; }
        public static string UserName { get; set; }
        public static string PassWord { get; set; }
        public static string DataBaseName { get; set; }
        public static string HostName { get; set; }

        
        public static ConfigUserInfo GetAppXmlDom(string strHostUrl)
        {
            ConfigUserInfo u = new ConfigUserInfo();
            string filename = System.Windows.Forms.Application.ExecutablePath + ".config";
            XmlDocument doc = new XmlDocument();
            doc.Load(filename); //加载配置文件
            //XmlNodeList NodeList =  doc.SelectNodes("//configuration/ConfigUserInfo/users");

            XmlNode node = doc.SelectSingleNode(string.Format("//configuration/ConfigUserInfo/users[@HostUrl='{0}']", strHostUrl));
            if (node != null)
            {
                XmlElement childElement = (XmlElement)node;
                u.HostUrl = childElement.GetAttribute("HostUrl");
                u.ServerName = childElement.GetAttribute("ServerName");
                u.UserName = childElement.GetAttribute("UserName");
                u.PassWord = childElement.GetAttribute("PassWord");
                u.DataBaseName = childElement.GetAttribute("DataBaseName");
                u.HostName = childElement.GetAttribute("HostName");
                return u;
            }
            else {
                return null;
            }

            //node.SelectSingleNode("//ServerName")
            //XmlSerializer x = new XmlSerializer(typeof(ConfigUserInfo));
            //using (StringReader sr = new StringReader(node.InnerXml))
            //{
            //    ConfigUserInfo s = (ConfigUserInfo)x.Deserialize(sr);
            //}
            
            //ConfigUserInfo/users
        }

        //得到第一个节点数据库连接信息
        public static ConfigUserInfo GetAppXmlDom_Default()
        {
            ConfigUserInfo u = new ConfigUserInfo();
            string filename = System.Windows.Forms.Application.ExecutablePath + ".config";
            XmlDocument doc = new XmlDocument();
            doc.Load(filename); //加载配置文件

            XmlNode node = doc.SelectSingleNode(string.Format("//configuration/ConfigUserInfo/users"));
            if (node != null)
            {
                XmlElement childElement = (XmlElement)node;
                u.HostUrl = childElement.GetAttribute("HostUrl");
                u.ServerName = childElement.GetAttribute("ServerName");
                u.UserName = childElement.GetAttribute("UserName");
                u.PassWord = childElement.GetAttribute("PassWord");
                u.DataBaseName = childElement.GetAttribute("DataBaseName");
                u.HostName = childElement.GetAttribute("HostName");
                return u;
            }
            else
            {
                return null;
            }
        }

        public class AccountConfiguration : ConfigurationSection
        {
            [ConfigurationProperty("users", IsRequired = true)]
            public AccountSectionElement Users
            {
                get { return (AccountSectionElement)this["users"]; }
            }
        }

        public class AccountSectionElement : ConfigurationElement
        {
            public AccountSectionElement()
            {
            }

            [ConfigurationProperty("username", IsRequired = true)]
            public string UserName
            {
                get { return this["username"].ToString(); }
                set { this["username"] = value; }
            }

            [ConfigurationProperty("password", IsRequired = true)]
            public string Password
            {
                get { return this["password"].ToString(); }
                set { this["password"] = value; }
            }
        }
        //静态获取数据库连接信息
        public static bool GetAppXmlDomInfo(string strHostUrl)
        {
            string filename = System.Windows.Forms.Application.ExecutablePath + ".config";
            XmlDocument doc = new XmlDocument();
            doc.Load(filename); //加载配置文件

            XmlNode node = doc.SelectSingleNode(string.Format("//configuration/ConfigUserInfo/users[@HostUrl='{0}']", strHostUrl));
            if (node != null)
            {
                XmlElement childElement = (XmlElement)node;
                GetXmlDataInfo.HostUrl = childElement.GetAttribute("HostUrl");
                GetXmlDataInfo.ServerName = childElement.GetAttribute("ServerName");
                GetXmlDataInfo.UserName = childElement.GetAttribute("UserName");
                GetXmlDataInfo.PassWord = childElement.GetAttribute("PassWord");
                GetXmlDataInfo.DataBaseName = childElement.GetAttribute("DataBaseName");
                GetXmlDataInfo.HostName = childElement.GetAttribute("HostName");
                return true;
            }
            else {
                return false;
            }

        }





        //静态获取数据库连接信息
        public static bool UpdateOrCreateAppSetting(string strHostUrl)
        {
            string filename = System.Windows.Forms.Application.ExecutablePath + ".config";
            XmlDocument doc = new XmlDocument();
            doc.Load(filename); //加载配置文件

            XmlNode node = doc.SelectSingleNode("//configuration/ConfigUserInfo");   //得到[appSettings]节点

            XmlElement element = (XmlElement)node.SelectSingleNode(string.Format("//users[@HostUrl='{0}']", strHostUrl));
            //XmlNode node = doc.SelectSingleNode(string.Format("//configuration/ConfigUserInfo/users[@HostUrl='{0}']", strHostUrl));
            try
            {
                if (element != null)
                {
                    //存在则更新子节点Value
                    element.SetAttribute("ServerName", GetXmlDataInfo.ServerName);
                    element.SetAttribute("UserName", GetXmlDataInfo.UserName);
                    element.SetAttribute("PassWord", GetXmlDataInfo.PassWord);
                    element.SetAttribute("DataBaseName", GetXmlDataInfo.DataBaseName);
                    element.SetAttribute("HostName", GetXmlDataInfo.HostName);
                }
                else
                {
                    //不存在则新增子节点
                    XmlElement subElement = doc.CreateElement("users");

                    subElement.SetAttribute("HostUrl", GetXmlDataInfo.HostUrl);
                    subElement.SetAttribute("ServerName", GetXmlDataInfo.ServerName);
                    subElement.SetAttribute("UserName", GetXmlDataInfo.UserName);
                    subElement.SetAttribute("PassWord", GetXmlDataInfo.PassWord);
                    subElement.SetAttribute("DataBaseName", GetXmlDataInfo.DataBaseName);
                    subElement.SetAttribute("HostName", GetXmlDataInfo.HostName);

                    node.AppendChild(subElement);
                }


                //保存至配置文件(方式一)
                //using (XmlTextWriter xmlwriter = new XmlTextWriter(filename, null))
                //{
                //    xmlwriter.Formatting = Formatting.Indented;
                //    doc.WriteTo(xmlwriter);
                //    xmlwriter.Flush();
                //}

                doc.Save(filename);
            }
            catch (Exception e)
            {
                return false;
            }


            return true;
        }










        //注册表
        public static void GetRegDataBase(string strRegName)
        {
            var appName = strRegName;

            var regKey = Registry.LocalMachine.OpenSubKey(@"Software\mysoft\" + appName, false) ?? Registry.LocalMachine.OpenSubKey(@"Software\Wow6432Node\mysoft\" + appName, false);

            if (regKey == null)
            {
                //throw MyException.GetMyExceptionWrapper("未配置注册表信息");
            }
                
            GetXmlDataInfo.ServerName = regKey.GetValue("ServerName", string.Empty).ToString();
            //serverProt = regKey.GetValue("ServerProt", string.Empty).ToString() == string.Empty ? "1433" : regKey.GetValue("ServerProt", string.Empty).ToString();
            GetXmlDataInfo.DataBaseName = regKey.GetValue("DBName", string.Empty).ToString();
            GetXmlDataInfo.UserName = regKey.GetValue("UserName", string.Empty).ToString();
            GetXmlDataInfo.PassWord = regKey.GetValue("SaPassword", string.Empty).ToString();

            Type type = typeof(AccountSectionElement);
            string strtype = type.GUID.ToString("B");
            //RegisterBHO(type);
        }


        /*
        /// <summary>
        /// 创建数据访问类
        /// </summary>
        /// <returns>SQLServer数据库连接类</returns>
        public static MyDataBase GetDataBase()
        {
            if (string.IsNullOrEmpty(serverName))
            {
                // var appName = AppConfig.GetAppSettingValue("ApplicationName");
                var appName = License.RegName;
                if (string.IsNullOrEmpty(appName))
                {
                    throw MyException.GetMyExceptionWrapper("没有配置应用程序名：ApplicationName");
                }

                var regKey = Registry.LocalMachine.OpenSubKey(@"Software\mysoft\" + appName, false) ?? Registry.LocalMachine.OpenSubKey(@"Software\Wow6432Node\mysoft\" + appName, false);

                if (regKey == null)
                {
                    throw MyException.GetMyExceptionWrapper("未配置注册表信息");
                }

                serverName = regKey.GetValue("ServerName", string.Empty).ToString();
                serverProt = regKey.GetValue("ServerProt", string.Empty).ToString() == string.Empty ? "1433" : regKey.GetValue("ServerProt", string.Empty).ToString();
                databaseName = regKey.GetValue("DBName", string.Empty).ToString();
                userName = regKey.GetValue("UserName", string.Empty).ToString();
                password = Cryptography.Decode(regKey.GetValue("SaPassword", string.Empty).ToString());
            }

            return new MyDataBase(serverName, serverProt, databaseName, userName, password);
        }
       */

        public static string BHOKETNAME = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKETNAME, true);

            if (registryKey == null)
            {
                registryKey = Registry.LocalMachine.CreateSubKey(BHOKETNAME);
            }

            string guid = type.GUID.ToString("B");
            string guid2 = System.Guid.NewGuid().ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid2);

            if (ourKey == null)
            {
                ourKey = registryKey.CreateSubKey(guid);
            }


            ourKey.SetValue("Alright", 1);
            registryKey.Close();
            ourKey.Close();

        }

    }
}

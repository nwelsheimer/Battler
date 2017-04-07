using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
//using MySql.Data;
//using MySql.Data.MySqlClient;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

//TODO(JLT): Set up App-level Registry properties in Program.cs files that can override Sawyer-level
//TODO(JLT): set up the cascading Form.Text idea so only one taskbar button will show text of currently open form, and revert when form is closed.
namespace General
{
    public static class Global
    {
        #region Public Properties
        public static string
            SQLCON,
            validEmail = @"\w{1,}[@][\w\-]{1,}([.]([\w\-]{1,}))+$",
            validPhone = @"^(\d{3})?\d{7}",
            strAppSubKey;
        public static RegistryKey keyHKLM_AppSubKey;
        #endregion

        #region Private Properties
        private static RegistryKey keyHKLM_TopRegKey;
        private static string strHKLM_TopRegKey = @"Software\Sawyer";
        private static SqlConnection sqlConn = null; //= new SqlConnection(SQLCON);
        #endregion

        #region Public Data Functions
        public static bool ExecuteQuery(string query)
        {
            CheckConnection();
            SqlCommand com = new SqlCommand(query, sqlConn);
            com.Connection.Open();
            try
            {
                com.ExecuteNonQuery();
                com.Connection.Close();
                return true;
            }
            catch (SqlException sqlX)
            {
                MessageBox.Show("SQL Error:\n" + sqlX.Message + "\n\nSQL: " + query, "SQL Error", MessageBoxButtons.OK);
                com.Connection.Close();
                return false;
            }
            catch (Exception X)
            {
                MessageBox.Show("Some other error:\n" + X.Message);
                com.Connection.Close();
                return false;
            }
        }

        public static void BulkUpdate(DataTable update, string tempTableName, string tempTableQuery, string updateQuery)
        {
            
            using (SqlCommand command = new SqlCommand("", sqlConn))
            {
                try
                {
                    sqlConn.Open();

                    //Creating temp table on database
                    command.CommandText = tempTableQuery;
                    command.ExecuteNonQuery();

                    //Bulk insert into temp table
                    using (SqlBulkCopy bulkcopy = new SqlBulkCopy(sqlConn))
                    {
                        bulkcopy.BulkCopyTimeout = 660;
                        bulkcopy.DestinationTableName = tempTableName;
                        bulkcopy.WriteToServer(update);
                        bulkcopy.Close();
                    }

                    // Updating destination table, and dropping temp table
                    command.CommandTimeout = 300;
                    command.CommandText = updateQuery;
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    // Handle exception properly
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            
        }

        public static string ExecuteQueryGetId(string query)
        {
            CheckConnection();
            SqlCommand com = new SqlCommand("SET NOCOUNT ON; " + query + "; SELECT @@IDENTITY AS id SET NOCOUNT OFF", sqlConn);
            com.Connection.Open();
            try
            {
                SqlDataReader sdr = com.ExecuteReader();
                if (sdr.Read())
                {
                    string id = sdr["id"].ToString();
                    com.Connection.Close();
                    return id;
                }
                else throw new Exception("Failed to retrieve new ID.");
            }
            catch (SqlException sqlX)
            {
                com.Connection.Close();
                MessageBox.Show("SQL Error occurred:\n" + sqlX.Message + "\n\nSQL: " + query, "SQL Error", MessageBoxButtons.OK);
                return "";
            }
            catch (Exception X)
            {
                com.Connection.Close();
                MessageBox.Show("Unknown error occurred:\n" + X.Message, "General Error", MessageBoxButtons.OK);
                return "";
            }
        }

        public static DataSet GetData(string query)
        {
            CheckConnection();
            SqlDataAdapter da = new SqlDataAdapter(query, sqlConn);
            DataSet ds = new DataSet();
            try
            {
                da.Fill(ds);
            }
            catch (SqlException sqlX)
            {
                //com.Connection.Close();
                MessageBox.Show("SQL Error occurred:\n" + sqlX.Message + "\n\nSQL: " + query, "SQL Error", MessageBoxButtons.OK);
                return null;
            }
            catch (Exception X)
            {
                //com.Connection.Close();
                MessageBox.Show("Unknown error occurred:\n" + X.Message, "General Error", MessageBoxButtons.OK);
                return null;
            }
            return ds;
        }

        public static object GetDataSingleValue(string query, string column)
        {
            DataRow dr = GetData(query).Tables[0].Rows[0];
            return dr[column];
        }

        public static int GetRowCount(string query)
        {
            CheckConnection();
            try
            {
                DataSet ds = GetData(query);
                return ds.Tables.Count == 0 ? 0 :
                    ds.Tables[0].Rows.Count;
            }
            catch (SqlException sqlX)
            {
                MessageBox.Show("SQL Error:\n" + sqlX.Message + "\n\nSQL: " + query, "SQL Error", MessageBoxButtons.OK);
                return -1;
            }
            catch (Exception X)
            {
                MessageBox.Show("Some other error:\n" + X.Message);
                return -1;
            }
        }

        //public static DataSet getMysqlData(string query)
        //{
        //    DataSet mds = new DataSet();
        //    MySqlDataAdapter mda = new MySqlDataAdapter();
        //    string MyConString = "SERVER=74.204.4.6;" +
        //      "DATABASE=asterisk1;" +
        //      "UID=tim;" +
        //      "PASSWORD=1q2w3e4r;";
        //    MySqlConnection connection = new MySqlConnection(MyConString);
        //    MySqlCommand command = connection.CreateCommand();
        //    connection.Open();
        //    command.CommandText = query;
        //    mda.SelectCommand = command;
        //    mda.Fill(mds);
        //    return mds;
        //}

        public static void SetConnectionString(string server, string db)
        {
            try
            {
                //Scabbing on functionality to connect to the remote SBI server, while hopefully maintaining this code for future in house apps
                if (server == "SBI")
                {
                    SQLCON = db;
                }
                else
                {
                    SQLCON = "Data Source=" + server + ";Initial Catalog=" + db + ";Integrated Security=True;Connection Timeout=180";
                }
            }
            catch (Exception)
            {
                SQLCON = GetConnectionString();
            }
        }
        #endregion

        # region Registry Methods (Public & Private)
        # region registry() overloads
        /// <summary>
        /// Method for READING registry values from Top Level key
        /// </summary>
        /// <param name="setting">The registry key setting to READ</param>
        /// <returns>The registry key's value as OBJECT</returns>
        public static object registry(string setting)
        {
            return keyHKLM_TopRegKey.GetValue(setting, null);
            //RegistryKey key = reg_key();
            //return key.GetValue(setting, setting + " is invalid").ToString();
        }
        /// <summary>
        /// Method for READING registry values from Top Level key
        /// </summary>
        /// <param name="key">RegistryKey to read from</param>
        /// <param name="setting">The registry key setting to READ</param>
        /// <returns>The registry key's value as OBJECT</returns>
        public static object registry(RegistryKey key, string setting)
        {
            return key.GetValue(setting, null);
            //RegistryKey key = reg_key();
            //return key.GetValue(setting, setting + " is invalid").ToString();
        }
        /// <summary>
        /// Method for WRITING registry STRING values to Top Level key
        /// </summary>
        /// <param name="setting">The registry key setting to WRITE</param>
        /// <param name="data">The registry key's value</param>
        /// <returns>Returns boolean success/fail of the attempt.</returns>
        public static bool registry(string setting, string data)
        {
            try
            {
                keyHKLM_TopRegKey.SetValue(setting, data, RegistryValueKind.String);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            //RegistryKey key = reg_key();
            //key.SetValue(setting, data, RegistryValueKind.String);
            //return; // Even if it fails? Why return anything on a write except bool?
        }
        /// <summary>
        /// Method for WRITING registry STRING values to Top Level key
        /// </summary>
        /// <param name="key">RegistryKey to write to</param>
        /// <param name="setting">The registry key setting to WRITE</param>
        /// <param name="data">The registry key's value</param>
        /// <returns>Returns boolean success/fail of the attempt.</returns>
        public static bool registry(RegistryKey key, string setting, string data)
        {
            try
            {
                key.SetValue(setting, data, RegistryValueKind.String);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            //RegistryKey key = reg_key();
            //key.SetValue(setting, data, RegistryValueKind.String);
            //return; // Even if it fails? Why return anything on a write except bool?
        }
        /// <summary>
        /// Method for WRITING registry values to Top Level key with specific value type
        /// </summary>
        /// <param name="setting">The registry setting to WRITE</param>
        /// <param name="data">The setting's value</param>
        /// <param name="rvKind">RegistryValueKind (data type) of the value</param>
        /// <returns>Returns boolean success/fail of the attempt.</returns>
        public static bool registry(string setting, string data, RegistryValueKind rvKind)
        {
            try
            {
                keyHKLM_TopRegKey.SetValue(setting, data, rvKind);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            //RegistryKey key = reg_key();
            //key.SetValue(setting, data, RegistryValueKind.String);
            //return; // Even if it fails? Why return anything on a write except bool?
        }
        /// <summary>
        /// Method for WRITING registry values to any key with specific value type
        /// </summary>
        /// <param name="key">RegistryKey to be modified</param>
        /// <param name="setting">The setting to WRITE</param>
        /// <param name="data">The setting's value</param>
        /// <param name="rvKind">RegistryValueKind (data type) of the value</param>
        /// <returns>Returns boolean success/fail of the attempt.</returns>
        public static bool registry(RegistryKey key, string setting, string data, RegistryValueKind rvKind)
        {
            try
            {
                key.SetValue(setting, data, rvKind);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            //RegistryKey key = reg_key();
            //key.SetValue(setting, data, RegistryValueKind.String);
            //return; // Even if it fails? Why return anything on a write except bool?
        }
        # endregion

        public static void set_AppSubKey(string ask)
        {
            try
            {
                strAppSubKey = ask;
                //JLT// NOT NOT NOT NOT!!! A good place for this stuff.
                keyHKLM_TopRegKey = get_reg_key(strHKLM_TopRegKey, true);
                keyHKLM_AppSubKey = get_reg_key(strHKLM_TopRegKey + @"\" + ask, true);
                //JLT// These should be in the parent app's Program.cs to make sure the top-level is present.
                registry(keyHKLM_TopRegKey, "Server", "sql1.vision.local");
                registry(keyHKLM_TopRegKey, "Database", "MSVision");
                registry(keyHKLM_TopRegKey, "TSXDsn", "Sawyer");
                keyHKLM_TopRegKey = get_reg_key(strHKLM_TopRegKey, false);
                SQLCON = GetConnectionString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "set_AppSubKey");
            }
        }

        private static void set_key_permissions(string keyPath)
        {
            RegistryPermission regPermission = new RegistryPermission(RegistryPermissionAccess.AllAccess, keyPath);
            regPermission.Demand();
        }

        /// <summary>
        /// Creates the desired key with hard-coded settings...Should be expanded.
        /// </summary>
        /// <param name="key">String key path to be created</param>
        /// <returns>RegistryKey object created from string key</returns>
        public static RegistryKey create_reg_key(string keyPath)
        {
            RegistryKey k = Registry.LocalMachine.CreateSubKey(keyPath); //, RegistryKeyPermissionCheck.ReadWriteSubTree);
            set_key_permissions(keyPath);
            k.SetValue("Server", "sql1.vision.local");
            k.SetValue("Database", "MSVision");
            k.SetValue("TSXDsn", "Sawyer");
            return k;
        }

        # region get_reg_key() overloads
        /// <summary>
        /// Typical registry key getter
        /// </summary>
        /// <returns>HKLM key for CURRENT APPLICATION</returns>
        public static RegistryKey get_reg_key()
        {
            //JLT// Manage this better. 
            // If the string for the current app isn't set, find a substitute to make it work.
            if (String.IsNullOrEmpty(strAppSubKey))
            {
                MessageBox.Show("Application Registry Key (AppSubKey) string has not been set.");
                return null;
            }
            //Initial check for key existence, populate with defaults otherwise
            string keyPath = strHKLM_TopRegKey + @"\" + strAppSubKey;
            RegistryKey key = get_reg_key(keyPath, true);
            //JLT// This should be extracted, and for each application, should be extended/overridden with an app-specific set of registry values.
            // Also, would be better if central DB-related registry stuff wasn't stored per-app.
            // Need to organize that hierarchy stuff.
            if (key == null)
            {
                key = create_reg_key(keyPath);
            }
            return key;
        }
        /// <summary>
        /// Get some other registry key
        /// </summary>
        /// <param name="key">String for HKLM key</param>
        /// <param name="write">Boolean whether to grant write permission</param>
        /// <returns>HKLM Registry key if found or created</returns>
        public static RegistryKey get_reg_key(string keyPath, bool write)
        {
            RegistryKey k = Registry.LocalMachine.OpenSubKey(keyPath, write);
            if (k == null)
            {
                return write ? create_reg_key(keyPath) : null;
            }
            else
            {
                set_key_permissions(keyPath);
                return k;
            }
        }
        # endregion
        # endregion

        #region Other Public Functions
        public static bool isNumeric(string num)
        {
            try
            {
                int.Parse(num);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public static string Capitalize(this String s)
        {
            char[] Cbreaks = new char[] { ' ', '\'', '-' };
            return s.ToCharArray().Aggregate(String.Empty,
                (working, next) =>
                    working.Length == 0 && !Cbreaks.Contains<char>(next) ?
                        next.ToString().ToUpper() :
                        (Cbreaks.Contains<char>(working.Substring(working.Length - 1)[0]) ?
                            working + next.ToString().ToUpper() :
                            working + next.ToString()
                        )
            );
        }

        public static string Prepend(this String s, string pre)
        {
            return pre + s;
        }


        public static bool CrystalEngineCheck()
        { //TODO(JLT): Make this check correctly if the Runtime is installed
            RegistryKey CRKey = get_reg_key(@"SOFTWARE\Business Objects\10.5\Crystal Reports", false);
            return CRKey != null && System.IO.Directory.Exists(CRKey.GetValue("Path").ToString());
        }

        public static System.Security.SecureString SecureString(string pass)
        {
            System.Security.SecureString SSpass = new System.Security.SecureString();
            foreach (char c in pass)
            {
                SSpass.AppendChar(c);
            }
            SSpass.MakeReadOnly();
            return SSpass;
        }

        public static bool ValidateEmailPattern(string email)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(email.Trim(), validEmail);
        }

        public static bool ValidatePhonePattern(string phone)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(
                System.Text.RegularExpressions.Regex.Replace(phone.Trim(), @"[^\d]", ""),
                validPhone);
        }
        # endregion

        public static void GridNavigation(UltraGrid grid, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Up:
                    grid.PerformAction(UltraGridAction.ExitEditMode, false, false);
                    grid.PerformAction(UltraGridAction.AboveCell, false, false);
                    e.Handled = true;
                    grid.PerformAction(UltraGridAction.EnterEditMode, false, false);
                    break;
                case Keys.Down:
                    grid.PerformAction(UltraGridAction.ExitEditMode, false, false);
                    grid.PerformAction(UltraGridAction.BelowCell, false, false);
                    e.Handled = true;
                    grid.PerformAction(UltraGridAction.EnterEditMode, false, false);
                    break;
                case Keys.Enter:
                    grid.PerformAction(UltraGridAction.ExitEditMode, false, false);
                    grid.PerformAction(UltraGridAction.BelowCell, false, false);
                    e.Handled = true;
                    grid.PerformAction(UltraGridAction.EnterEditMode, false, false);
                    break;
            }
        }

        public static string GridLayout(UltraGrid grid, int mode, string settings="")
        {
            byte[] loa;
            System.IO.MemoryStream lo;

            switch (mode)
            {
                case 1://We are doing a save
                    lo = new System.IO.MemoryStream();
                    grid.DisplayLayout.Save(lo, PropertyCategories.All);
                    loa = lo.ToArray();
                    settings = System.Convert.ToBase64String(loa);
                    return settings;
                case 2://We are doing a load
                    loa = System.Convert.FromBase64String(settings);
                    lo = new System.IO.MemoryStream(loa);
                    lo.Seek(0, System.IO.SeekOrigin.Begin);
                    try { grid.DisplayLayout.Load(lo, PropertyCategories.All); } catch { }
                    return "SUCCESS";
            }
            return "FAILED";
        }

        public static char LetterToNum(int column)
        {
            Char colLetter = (Char)(65 + (column - 1));
            return colLetter;
        }
        #region Private Functions
        private static void CheckConnection()
        {
            if (sqlConn == null)
            {
                sqlConn = new SqlConnection(SQLCON);
            }
        }

        private static string GetConnectionString()
        {
            try
            {
                string server = registry("Server").ToString();
                string db = registry("Database").ToString();
                return "Data Source=" + server + ";Initial Catalog=" + db + ";Integrated Security=True;";
            }
            catch (Exception ex)
            {
                //run setup wizard
                return "ERROR " + ex.Message;
            }
        }
        #endregion
    }

    //JLT//
    /**
     * We're eventually going to need something that manages the registry with more flexibility and security control.
     * What it'll ultimately look like, i don't know. Need classes or books or something.
     * */
    public class SawyerRegistry //: Microsoft.Win32
    {
        private string topLevel = @"Software\Sawyer";
        private RegistryPermission regPerm;
        //internal struct SKs
        //{
        private string
            _sawyerHklmName,
            _sawyerHkcuName,
            _appHklmName,
            _appHkcuName;
        private RegistryKey
            _sawyerHKLM,
            _sawyerHKCU,
            _appHKLM,
            _appHKCU;
        internal string SawyerHklmName
        {
            get { return _sawyerHklmName; }
            private set { _sawyerHklmName = @value; }
        }
        internal string SawyerHkcuName
        {
            get { return _sawyerHkcuName; }
            private set { _sawyerHkcuName = @value; }
        }
        internal string AppHklmName
        {
            get { return _appHklmName; }
            private set { _appHklmName = @value; }
        }
        internal string AppHkcuName
        {
            get { return _appHkcuName; }
            private set { _appHkcuName = @value; }
        }
        internal RegistryKey SawyerHKLM
        {
            get { return _sawyerHKLM; }
            private set { _sawyerHKLM = value; }
        }
        internal RegistryKey SawyerHKCU
        {
            get { return _sawyerHKCU; }
            private set { _sawyerHKCU = value; }
        }
        internal RegistryKey AppHKLM
        {
            get { return _appHKLM; }
            private set { _appHKLM = value; }
        }
        internal RegistryKey AppHKCU
        {
            get { return _appHKCU; }
            private set { _appHKCU = value; }
        }
        //}
        //private SKs SubKeys = new SKs();

        internal SawyerRegistry() { }
        internal SawyerRegistry(string AppName)
        {
            SetAppKey(AppName);
        }
        internal void SetAppKey(string AppName)
        {
            try
            {
                //SubKeys.SawyerHklmName = @topLevel;
                //SubKeys.SawyerHkcuName = @topLevel;
                SawyerHklmName = @topLevel;
                SawyerHkcuName = @topLevel;
                AppHklmName = @topLevel + @"\" + AppName;
                AppHklmName = @topLevel + @"\" + AppName;
                string
                    hklm = @"HKEY_LOCAL_MACHINE\",
                    hkcu = @"HKEY_CURRENT_USER\";
                regPerm = new RegistryPermission(PermissionState.Unrestricted);
                regPerm.AddPathList(RegistryPermissionAccess.AllAccess,
                    hklm + topLevel + ";" +
                    hklm + topLevel + @"\" + AppName + ";" +
                    hkcu + topLevel + ";" +
                    hkcu + topLevel + @"\" + AppName);
                regPerm.Demand();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        internal RegistryKey CreateLocalMachineKey(string @keyPath)
        {
            return Registry.LocalMachine.CreateSubKey(keyPath);
        }
    }

    public class Desc : IComparer
    {
        public int Compare(object x, object y)
        {
            try
            {
                if ((int)x > (int)y) { return -1; }
                if ((int)x == (int)y) { return 0; }
                if ((int)x < (int)y) { return 1; }
                return 0;
            }
            catch (Exception)
            {
                try
                {
                    if (((string)x).CompareTo((string)y).Equals(1)) { return -1; }
                    if (((string)x).CompareTo((string)y).Equals(0)) { return 0; }
                    if (((string)x).CompareTo((string)y).Equals(-1)) { return 1; }
                    return 0;
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
    }
    public class Asc : IComparer
    {
        public int Compare(object x, object y)
        {
            try
            {
                return ((int)x).CompareTo((int)y);
            }
            catch (Exception)
            {
                try
                {
                    return ((string)x).CompareTo((string)y);
                }
                catch (Exception)
                {
                    return 0;
                }
            }
        }
    }

    //JLT// Just an idea:
    // Set up class, gets instanced when the app opens
    // Each form's creation and close adds/removes itself from the app's collection
    // When a child form doesn't have a TaskbarButton, it can access the parent's to set the text
    // The parent's text is stored here so when the child is closed, the taskbar can be set to the lowest-open form
    //public class TaskbarText
    //{
    //    private string[] texts;
    //    private Form[] forms;
    //    private Form TopForm;

    //    public TaskbarText(Form form)
    //    {
    //        forms[forms.Length] = form;
    //        texts[texts.Length] = form.Text;
    //    }

    //    public void AddForm(Form form)
    //    {
    //        if (!forms.Contains<Form>(form))
    //        {
    //            forms[forms.Length] = form;
    //            texts[texts.Length] = form.Text;
    //        }
    //    }

    //    public void SetFormText(Form form, string text)
    //    {
    //        this.AddForm(form);
    //        for (int i = 0; i < forms.Length; i++)
    //        {
    //            if (forms[i] == form)
    //            {
    //                texts[i] = text;
    //                break;
    //            }
    //        }
    //    }

    //    public void SetTaskbarText(Form f, string text)
    //    {
    //        while (!f.ShowInTaskbar)
    //        { f = f.Owner; }
    //        f.Text = text;
    //    }


    //}

}

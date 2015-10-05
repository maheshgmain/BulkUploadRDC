using System;
using System.Configuration;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
//using Atalasoft.Imaging;
//using Atalasoft.Imaging.Codec;
//using Atalasoft.Imaging.Codec.Tiff;
using System.Globalization;
using System.Security.Cryptography;

namespace UploadApplication
{
    public partial class Form2 : Form
    {
        SqlConnection Sqlconn = new SqlConnection(ConfigurationSettings.AppSettings["SqlCon"].ToString());

        bool ErrorOccured = false;
        int DocTypeID = 0;
        string TableName = "tblcustomerDatails", colname = "";
        DataTable dtExceldata = null;
        DataTable dtDocDetails = null;
        int w = 0;
        string ImageField = "", strDateFormat = "";
        public Form2()
        {
            InitializeComponent();

            //mandatory. Otherwise will throw an exception when calling ReportProgress method  
            backgroundWorker1.WorkerReportsProgress = true;

            //mandatory. Otherwise we would get an InvalidOperationException when trying to cancel the operation  
            backgroundWorker1.WorkerSupportsCancellation = true;

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            DateTime datImportedOn = new DateTime();
            string strDate = System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
            if (strDate.IndexOf("-") != -1)
                datImportedOn = DateTime.ParseExact(strDate, "dd-MM-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            else
                datImportedOn = DateTime.ParseExact(strDate, "dd/MM/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);



            if (datImportedOn.ToString("MM/dd/yyyy") == "03/15/2014")
            {
                MessageBox.Show("Application has exipered, please contact to administrator.");
                this.Close();
                return;
            }
            this.Size = new Size(754, 135);
            this.groupBox1.Size = new Size(729, 116);

            //DataTable dt = GetDataTable("select docLabel,DocId from tblDocmaster");
            //if (dt.Rows.Count > 0)
            //{
            //cmbDocumentType.DisplayMember = "docLabel";
            //cmbDocumentType.ValueMember = "DocId";
            //cmbDocumentType.DataSource = dt;

            dtDocDetails = GetDataTable(string.Format(@"SELECT character_maximum_length as 'Max Length'
,column_name as 'ColumnName', data_type as 'Data Type'
FROM information_schema.columns
WHERE table_name = 'tblcustomerDatails'"));
            LoadListBox(dtDocDetails, lstDatabaseColumn, "row");

            btnBrowse.Focus();
            //}
            //else
            //{
            //    MessageBox.Show("Create document type.");
            //    this.Close();
            //}


            //bind format of date in excel
            // CultureInfo ci = CultureInfo.InvariantCulture;
            CultureInfo ci = CultureInfo.GetCultureInfo("en-gb");
            CultureInfo ci1 = CultureInfo.GetCultureInfo("en-us");
           // CultureInfo ci2 = CultureInfo.GetCultureInfo("en-IN");

            string[] fmts = ci.DateTimeFormat.GetAllDateTimePatterns();
            string[] fmts1 = ci1.DateTimeFormat.GetAllDateTimePatterns();
            //string[] fmt2 = ci2.DateTimeFormat.GetAllDateTimePatterns();
            //foreach (string format2 in fmt2)
            //{
            //    comboBox2.Items.Add(format2);
            //}
            foreach (string format in fmts)
            {
                comboBox2.Items.Add(format);
            }
            foreach (string format1 in fmts1)
            {
                comboBox2.Items.Add(format1);
            }

            comboBox2.Items.Insert(0, "Select");
            comboBox2.SelectedIndex = 0;

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            OpenFileDialog oldBrowser = new OpenFileDialog();
            oldBrowser.Filter = "Excel File (*.xls)|*.xls | Excel File (*.xlsx)|*.xlsx";
            oldBrowser.ShowDialog();
            txtXLSFileName.Text = oldBrowser.FileName;
            btnGetXLSColumn.Focus();
            this.Cursor = Cursors.Default;
        }

        private void btnGetXLSColumn_Click(object sender, EventArgs e)
        {
            if (txtXLSFileName.Text.Length == 0)
            {
                MessageBox.Show("Select excel file to import data.");
                btnBrowse.Focus();
                return;
            }
            dtExceldata = GetExcelDataTable(Path.GetDirectoryName(txtXLSFileName.Text), Path.GetFileName(txtXLSFileName.Text));
            if (dtExceldata.Rows.Count > 0)
            {
                LoadListBox(dtExceldata, lstXLSColumn, "column");

                this.Size = new Size(754, 453);
                this.groupBox1.Size = new Size(729, 429);

                lbltotalcount.Text = Convert.ToString(dtExceldata.Rows.Count);

                progressBar1.Maximum = dtExceldata.Rows.Count + 1;
                btnGetXLSColumn.Enabled = false;
                btoStart.Enabled = true;
            }
            else
            {
                this.Size = new Size(754, 135);
                this.groupBox1.Size = new Size(729, 116);

                btnGetXLSColumn.Enabled = true;
                btoStart.Enabled = false;
            }
        }

        private void btnMoveRight_XlsColumn_Click(object sender, EventArgs e)
        {
            if (lstXLSColumn.SelectedItem != null)
            {
                listBox2.Items.Add(lstXLSColumn.SelectedItem);
                lstXLSColumn.Items.Remove(lstXLSColumn.SelectedItem);
            }
            else
            {
                MessageBox.Show("Select column name.");
            }
        }

        private void btnMoveLeft_XlsColumn_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
            {
                if (!lstXLSColumn.Items.Contains(listBox2.SelectedItem))
                    lstXLSColumn.Items.Add(listBox2.SelectedItem);

                listBox2.Items.Remove(listBox2.SelectedItem);
            }
            else { MessageBox.Show("Select column name."); }
        }

        private void btnMoveLeft_Database_Click(object sender, EventArgs e)
        {
            if (lstDatabaseColumn.SelectedItem != null)
            {
                listBox3.Items.Add(lstDatabaseColumn.SelectedItem);
                lstDatabaseColumn.Items.Remove(lstDatabaseColumn.SelectedItem);
            }
            else { MessageBox.Show("Select column name."); }
        }

        private void btnMoveRight_Database_Click(object sender, EventArgs e)
        {
            if (listBox3.SelectedItem != null)
            {
                if (!lstDatabaseColumn.Items.Contains(listBox3.SelectedItem))
                    lstDatabaseColumn.Items.Add(listBox3.SelectedItem);

                listBox3.Items.Remove(listBox3.SelectedItem);
            }
            else { MessageBox.Show("Select column name."); }
        }

        private void lstXLSColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstXLSColumn.SelectedItem != null)
            {
                btnMoveRight_XlsColumn.Enabled = true;
                btnMoveLeft_XlsColumn.Enabled = false;
            }
            else
            {
                btnMoveRight_XlsColumn.Enabled = false;
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
            {
                btnMoveRight_XlsColumn.Enabled = false;
                btnMoveLeft_XlsColumn.Enabled = true;
            }
            else
            {
                btnMoveLeft_XlsColumn.Enabled = false;
            }
        }

        private void lstDatabaseColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstDatabaseColumn.SelectedItem != null)
            {
                btnMoveLeft_Database.Enabled = true;
                btnMoveRight_Database.Enabled = false;
            }
            else
            {
                btnMoveLeft_Database.Enabled = false;
            }
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox3.SelectedItem != null)
            {
                btnMoveLeft_Database.Enabled = false;
                btnMoveRight_Database.Enabled = true;
            }
            else
            {
                btnMoveRight_Database.Enabled = false;
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtXLSFileName.Text = "";
            btnGetXLSColumn.Enabled = true;
            btoStart.Enabled = false;
            btoCancel.Enabled = false;

            this.Size = new Size(754, 135);
            this.groupBox1.Size = new Size(729, 116);

            //cleared
            ErrorOccured = false;
            // DocTypeID = 0;
            //  TableName = "";
            colname = "";
            dtExceldata = null;
            w = 0;
            ImageField = "";
            txtXLSFileName.Text = "";
            lblimportedcount.Text = "0";
            lbltotalcount.Text = "0";
            strDateFormat = "";
            progressBar1.Value = 0;
            lstXLSColumn.Items.Clear();
            lstDatabaseColumn.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            comboBox2.SelectedIndex = 0;
            //bind lstdatabase columns
            LoadListBox(dtDocDetails, lstDatabaseColumn, "row");

            btnBrowse.Focus();
        }

        private void btoStart_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count > 0 && listBox3.Items.Count > 0)
            {
                if (listBox2.Items.Count == listBox3.Items.Count)
                {
                    //if (comboBox1.SelectedIndex != 0)
                    //{
                    btoStart.Enabled = false;
                    btoCancel.Enabled = true;
                    backgroundWorker1.RunWorkerAsync();
                    //}
                    //else
                    //{
                    //    MessageBox.Show("Please select image field Name");
                    //    comboBox1.Focus();
                    //}
                }
                else
                {
                    MessageBox.Show("Mismatch in column mapping");
                    listBox2.Focus();
                }
            }
            else
            {
                MessageBox.Show("Please map the columns.");
                listBox2.Focus();
            }
        }

        private void btoCancel_Click(object sender, EventArgs e)
        {
            //notify background worker we want to cancel the operation.  
            //this code doesn't actually cancel or kill the thread that is executing the job.  
            btoStart.Enabled = true;
            btoCancel.Enabled = false;
            backgroundWorker1.CancelAsync();
        }



        private Int32 FramCount(string ImagePath)
        {
            Int32 frmCount;
            try
            {
                // frmCount = RegisteredDecoders.GetImageInfo(ImagePath).FrameCount;
                using (FileStream fs = File.OpenRead(ImagePath))
                {
                    Stream retStream = new MemoryStream();
                    using (System.Drawing.Image img = System.Drawing.Image.FromStream(fs, true, true))
                    {
                        frmCount = img.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);

                    }
                    return frmCount;
                }
                //return frmCount;
            }
            catch
            {
                return -1;
            }
        }
        //public static Int32 FramCount()
        //{
        //    Int32 frmCount;
        //    if (Path.GetExtension(HttpContext.Current.Session["FilePath"].ToString().ToLower()) != ".pdf")
        //        using (FileStream fs = File.OpenRead(HttpContext.Current.Session["FilePath"].ToString()))
        //        {
        //            Stream retStream = new MemoryStream();
        //            using (System.Drawing.Image img = System.Drawing.Image.FromStream(fs, true, true))
        //            {
        //                frmCount = img.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);

        //            }
        //            return frmCount;
        //        }
        //    else
        //    {
        //        OperationClass operation = new OperationClass();
        //        if (operation.convertPDF_TIFF(HttpContext.Current.Session["FilePath"].ToString()))
        //        {
        //            string FileName = HttpContext.Current.Session["FilePath"].ToString() + ".tif";
        //            using (FileStream fs = File.OpenRead(FileName))
        //            {
        //                Stream retStream = new MemoryStream();
        //                using (System.Drawing.Image img = System.Drawing.Image.FromStream(fs, true, true))
        //                {
        //                    frmCount = img.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);

        //                }
        //                return frmCount;
        //            }
        //        }
        //        else { return -1; }
        //    }

        //}
        //private void InsertImageDetails(string strFileName, string strFilePath, byte[] imgFileImage,
        //                Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo,
        //                Int32 intFileStatus, string strInitialPath, string strIpAddress, string TableName,
        //    Int32 PageCount, string columnName, string ColumnValue, string strFullFilePath)
        //{

        //    string strDate = System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
        //    if (strDate.IndexOf("-") != -1)
        //        datImportedOn = DateTime.ParseExact(strDate, "dd-MM-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
        //    else
        //        datImportedOn = DateTime.ParseExact(strDate, "dd/MM/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);


        //    ErrorOccured = false;
        //    //insert data into tblfile

        //    //string Folder = ConfigurationManager.AppSettings["Repository"].ToString() + TableName;
        //    //string strtime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + DateTime.Now.Millisecond.ToString("000");
        //    //if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
        //    //{
        //    //    strFilePath = "~\\Repository\\Encrypt\\" + TableName + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc";
        //    //    //  ViewState["Folder"] = Folder;
        //    //    if (!Directory.Exists(Folder))
        //    //    {
        //    //        Directory.CreateDirecory(Folder);
        //    //    }
        //    //    imgFileImage = new byte[1];
        //    //}

        //    //string DocInsert = InsertDocument(strFileName, strFilePath, imgFileImage, intDocId, strImportedBy, datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);

        //    string DocInsert = "";
        //    if (DocInsert == "true")
        //    {
        //        if (ColumnValue != "")
        //        {
        //            //get id of last record inserted
        //            DataTable maxidds = GetDataTable("select MAX(Id) from tblFile");
        //            string maxid = maxidds.Rows[0][0].ToString();

        //            //insert into index table
        //            string qury = "insert into " + TableName + " (File_Id,ImportedBy,ImportedOn," + columnName + " ) values(" + maxid + ",'Capture','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'," + ColumnValue + " )";
        //            if (dataUpdate(qury) == true)
        //            {
        //                //insert into tblfileactions table
        //                updatedata("Insert into TblFileActions(FileId,DocId,FileName,TransProcess,UserName,TransDate) values(" + maxid + "," + intDocId + ",'" + strFileName + "','Uploaded','Capture','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "')");

        //                //save image
        //                if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
        //                    ImageEncrypt(strFullFilePath, Folder + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc", "encrypt123");
        //            }
        //            else
        //            {
        //                ErrorOccured = true;
        //                //if error occured then delete inserted record from tblfile
        //                updatedata("delete from tblfile where id=" + maxid + ")");
        //                createlog("Error occured while insert in " + TableName + " document table : " + qury);
        //            }
        //        }
        //    }
        //}

        public void createlog(string msg)
        {
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(txtXLSFileName.Text) + "\\ImportLog"))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(txtXLSFileName.Text) + "\\ImportLog");
                }
                string logformat = System.DateTime.Now.ToString("ddMMyyyy");
                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(txtXLSFileName.Text) + "\\ImportLog\\" + "Importlogfile_" + logformat + ".txt", true);

                sw.WriteLine(msg);
                sw.Close();

            }
            catch
            {

            }

        }

        public string InsertDocument(string strFileName, string strFilePath, byte[] imgFileImage, Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo, int intFileStatus, string strInitialPath, string strIpAddress, Int32 PageCount)
        {
            string temp; string strQuery = "";
            try
            {
                if (Sqlconn.State == ConnectionState.Closed)
                    Sqlconn.Open();

                strQuery = "insert into tblFile (FileName,FilePath,fileImage,DocId,ImportedBy,ImportedOn,FileSize,VersionNo,FileStatus,Initialpath,IpAddress,PageCount) values(@FileName,@FilePath,@fileImage,@DocId,@ImportedBy,@ImportedOn,@FileSize,@VersionNo,@FileStatus,@InitialPath,@IpAddress,@PageCount)";
                // strQuery = "insert into tblFile (FileName,FilePath,fileImage,DocId,ImportedBy,FileSize,VersionNo,FileStatus,Initialpath,IpAddress,PageCount) values(@FileName,@FilePath,@fileImage,@DocId,@ImportedBy,@FileSize,@VersionNo,@FileStatus,@InitialPath,@IpAddress,@PageCount)";
                SqlCommand sqlCmd = new SqlCommand();
                SqlParameter prmFileName, prmFilePath, prmFileImage, prmDocId, prmImportedBy, prmImportedOn, prmFileSize, prmVersionNo, prmFileStatus, prmInitalPath, prmIpAddress, prmPageCount;

                sqlCmd.Connection = Sqlconn;
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.CommandText = strQuery;

                prmFileName = new SqlParameter();
                prmFileName.ParameterName = "@FileName";
                prmFileName.SqlDbType = SqlDbType.VarChar;
                prmFileName.Direction = ParameterDirection.Input;
                prmFileName.Value = strFileName.Trim();
                sqlCmd.Parameters.Add(prmFileName);

                prmFilePath = new SqlParameter();
                prmFilePath.ParameterName = "@FilePath";
                prmFilePath.SqlDbType = SqlDbType.VarChar;
                prmFilePath.Direction = ParameterDirection.Input;
                prmFilePath.Value = strFilePath.Trim();
                sqlCmd.Parameters.Add(prmFilePath);

                prmFileImage = new SqlParameter();
                prmFileImage.ParameterName = "@fileImage";
                //prmFileImage.SqlDbType = SqlDbType.Image;
                prmFileImage.Direction = ParameterDirection.Input;
                prmFileImage.Value = imgFileImage;
                sqlCmd.Parameters.Add(prmFileImage);

                prmDocId = new SqlParameter();
                prmDocId.ParameterName = "@DocId";
                prmDocId.SqlDbType = SqlDbType.Int;
                prmDocId.Direction = ParameterDirection.Input;
                prmDocId.Value = intDocId;
                sqlCmd.Parameters.Add(prmDocId);

                prmImportedBy = new SqlParameter();
                prmImportedBy.ParameterName = "@ImportedBy";
                prmImportedBy.SqlDbType = SqlDbType.VarChar;
                prmImportedBy.Direction = ParameterDirection.Input;
                prmImportedBy.Value = strImportedBy.Trim();
                sqlCmd.Parameters.Add(prmImportedBy);

                prmImportedOn = new SqlParameter();
                prmImportedOn.ParameterName = "@ImportedOn";
                prmImportedOn.SqlDbType = SqlDbType.DateTime;
                prmImportedOn.Direction = ParameterDirection.Input;
                prmImportedOn.Value = datImportedOn;

                sqlCmd.Parameters.Add(prmImportedOn);

                prmFileSize = new SqlParameter();
                prmFileSize.ParameterName = "@FileSize";
                prmFileSize.SqlDbType = SqlDbType.VarChar;
                prmFileSize.Direction = ParameterDirection.Input;
                prmFileSize.Value = strFileSize.Trim();
                sqlCmd.Parameters.Add(prmFileSize);

                prmVersionNo = new SqlParameter();
                prmVersionNo.ParameterName = "@VersionNo";
                prmVersionNo.SqlDbType = SqlDbType.VarChar;
                prmVersionNo.Direction = ParameterDirection.Input;
                prmVersionNo.Value = intVersionNo;
                sqlCmd.Parameters.Add(prmVersionNo);

                prmFileStatus = new SqlParameter();
                prmFileStatus.ParameterName = "@FileStatus";
                prmFileStatus.SqlDbType = SqlDbType.Int;
                prmFileStatus.Direction = ParameterDirection.Input;
                prmFileStatus.Value = intFileStatus;
                sqlCmd.Parameters.Add(prmFileStatus);

                prmInitalPath = new SqlParameter();
                prmInitalPath.ParameterName = "@InitialPath";
                prmInitalPath.SqlDbType = SqlDbType.VarChar;
                prmInitalPath.Direction = ParameterDirection.Input;
                prmInitalPath.Value = strInitialPath;
                sqlCmd.Parameters.Add(prmInitalPath);

                prmIpAddress = new SqlParameter();
                prmIpAddress.ParameterName = "@IpAddress";
                prmIpAddress.SqlDbType = SqlDbType.VarChar;
                prmIpAddress.Direction = ParameterDirection.Input;
                prmIpAddress.Value = strIpAddress;
                sqlCmd.Parameters.Add(prmIpAddress);

                prmPageCount = new SqlParameter();
                prmPageCount.ParameterName = "@PageCount";
                prmPageCount.SqlDbType = SqlDbType.Int;
                prmPageCount.Direction = ParameterDirection.Input;
                prmPageCount.Value = PageCount;
                sqlCmd.Parameters.Add(prmPageCount);
                int i = sqlCmd.ExecuteNonQuery();
                temp = "true";
            }
            catch (System.Exception ex)
            {
                ErrorOccured = true;
                createlog("Error occured while insert in tblfile table : " + strQuery);
                createlog("Error Details : " + ex.Message);
                temp = ex.Message;
            }
            finally
            {
                Sqlconn.Close();
            }
            return temp;
        }

        private DataTable GetExcelDataTable(string Path, string FileName)
        {
            dtExceldata = new DataTable();

            // string _ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + System.IO.Path.Combine(Path, FileName) + ";" + "Extended Properties=Excel 8.0;";
            string _ConnectionString = null;

            if (System.IO.Path.GetExtension(FileName) == ".xls")
                _ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + System.IO.Path.Combine(Path, FileName) + ";" + "Extended Properties=Excel 8.0;";
            else
                _ConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;Persist Security Info=False", System.IO.Path.Combine(Path, FileName));

            OleDbConnection objConn = new OleDbConnection(_ConnectionString);
            objConn.Open();

            DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt.Rows.Count > 0)
            {
                OleDbCommand ocmd = new OleDbCommand("SELECT * FROM [" + dt.Rows[0]["TABLE_NAME"].ToString() + "]", objConn);
                OleDbDataReader odr = ocmd.ExecuteReader();
                dtExceldata.Load(odr);

                if (dtExceldata.Rows.Count > 0)
                {

                    foreach (DataRow row in dtExceldata.Rows)
                    {
                        String valuesarr = String.Empty;
                        List<object> lst = row.ItemArray.ToList();
                        foreach (Object s in lst)
                        {
                            valuesarr += s.ToString();
                        }

                        if (String.IsNullOrEmpty(valuesarr))
                        {
                            //dtExceldata.Rows.RemoveAt(i);
                            row.Delete();

                        }
                    } dtExceldata.AcceptChanges();
                }
            }
            else
                MessageBox.Show("Excel not read, please try again.");

            return dtExceldata;
        }

        public void updatedata(string qury)
        {
            try
            {

                if (Sqlconn.State == ConnectionState.Closed)
                {
                    Sqlconn.Open();
                }

                SqlCommand sqlCmd = new SqlCommand(qury, Sqlconn);
                sqlCmd.CommandTimeout = 0;
                sqlCmd.ExecuteNonQuery();



            }
            catch (Exception ex)
            {
                ErrorOccured = true;
                createlog("Error occured while insert in tblfileAction table : " + qury);
                string strEx = ex.Message;
            }
            finally
            {

                Sqlconn.Close();


            }

        }

        public Boolean dataUpdate(string qury)
        {
            Boolean temp;
            try
            {

                if (Sqlconn.State == ConnectionState.Closed)
                {
                    Sqlconn.Open();

                }
                SqlCommand sqlCmd = new SqlCommand(qury, Sqlconn);
                sqlCmd.ExecuteNonQuery();
                temp = true;


            }
            catch (Exception ex)
            {
                temp = false;
            }

            finally
            {

                Sqlconn.Close();

            }

            return temp;
        }

        public string RetureExcuteScaler(string qury)
        {
            string retStr = "";
            try
            {

                if (Sqlconn.State == ConnectionState.Closed)
                {
                    Sqlconn.Open();

                }
                SqlCommand sqlCmd = new SqlCommand(qury, Sqlconn);
                retStr = Convert.ToString(sqlCmd.ExecuteScalar());

            }
            catch (Exception ex)
            {
                retStr = "";
            }

            finally
            {

                Sqlconn.Close();

            }

            return retStr;
        }

        private DataTable GetDataTable(string sqlquery)
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

        private void LoadListBox(DataTable dt, ListBox ListBox, string AddColumnOrRow)
        {
            ListBox.Items.Clear();

            if (AddColumnOrRow == "row")
            {
                foreach (DataRow row in dt.Rows)
                {
                    ListBox.Items.Add(row[1].ToString());
                }
            }
            else
            {
                comboBox1.Items.Clear();
                foreach (DataColumn col in dt.Columns)
                {
                    ListBox.Items.Add(col.ColumnName);
                    comboBox1.Items.Add(col.ColumnName);
                }
                comboBox1.Items.Insert(0, "Select");
                comboBox1.SelectedIndex = 0;
            }
        }

        private void SetColumnName()
        {
            colname = "";
            //DataTable dtDatabaseTable = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where docId =" + Convert.ToString(DocTypeID));
            //if (dtDatabaseTable.Rows.Count > 0)
            //{
            //    foreach (object o in listBox3.Items)
            //    {
            //        foreach (DataRow drDatabase in dtDatabaseTable.Rows)
            //        {
            //            if (drDatabase["label"].ToString() == o.ToString())
            //            {
            //                if (colname == "")
            //                    colname = drDatabase["FieldName"].ToString();
            //                else
            //                    colname = colname + "," + drDatabase["FieldName"].ToString();
            //            }
            //        }
            //    }
            //}
            foreach (object o in listBox3.Items)
            {
                foreach (DataRow column in dtDocDetails.Rows)
                {
                    if (column["ColumnName"].ToString() == o.ToString())
                        if (colname == "")
                            colname = "[" + column["ColumnName"].ToString() + "]";
                        else colname = colname + "," + "[" + column["ColumnName"].ToString() + "]";
                }
            }
        }

        #region BackgroundWorker Events
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage; //update progress bar  

            DateTime time = Convert.ToDateTime(e.UserState); //get additional information about progress  

            lblimportedcount.Text = (w + 1).ToString();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;

            e.Result = "";

            string colValue = ""; string strrec = "";
            //string strImportedBy = "Capture";
            //DateTime datImportedOn = System.DateTime.Now;
            //// DataTable dtDatabaseTable; DataTable dtExceldata;
            //int intVersionNo = 1; int intFileStatus = 1; int filenotfoundcount = 0; int OtherErrorCount = 0;
            //string strIpAddress = ""; string FileSize = "";
            try
            {
                SetColumnName();//set columns value here

                //writting log file started from here.
                // createlog("###############################  IMPORT LOG FOR Document type  " + cmbDocumentType.SelectedText + "    ############################################");
                createlog("...................................................................................................................................................");
                createlog(" Import started at : " + DateTime.Now);
                createlog("......................................................................................................................................");

                //get selected acadmic year
                //string acadmicYear = RetureExcuteScaler(string.Format(@"select companycode from tblcompanymaster where isselected=1"));

                //loop started on excel datatable
                foreach (DataRow dtExcelrow in dtExceldata.Rows)
                {
                    System.Threading.Thread.Sleep(50); //do some intense task here.  
                    backgroundWorker1.ReportProgress(w, DateTime.Now); //notify progress to main thread. We also pass time information in UserState to cover this property in the example.  

                    colValue = "";
                    ErrorOccured = false;

                    for (int j = 0; j < listBox3.Items.Count; j++)
                    {
                        DataView dv = new DataView(dtDocDetails, "ColumnName='" + Convert.ToString(listBox3.Items[j]) + "'", "", DataViewRowState.CurrentRows);
                        if (dv.Count > 0)
                        {

                            if (Convert.ToString(dv[0]["Data Type"]).ToLower() == "datetime")
                            {
                                if (dtExcelrow[listBox2.Items[j].ToString()].ToString().Trim() != "")
                                {
                                    string dateString;
                                    DateTime result;
                                    try
                                    {
                                        CultureInfo provider = CultureInfo.InvariantCulture;
                                        strDateFormat = "dd/MM/yyyy HH:mm:ss";
                                        dateString = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]); //Input Date
                                        //format = "dd/MM/yyyy";//// Only change this format as per input date formate 
                                        result = DateTime.ParseExact(dateString, strDateFormat, provider);
                                        string dateDDMMYYY = result.ToString("MM/dd/yyyy");
                                        dateDDMMYYY = dateDDMMYYY.Replace("-", "/");

                                        if (colValue == "")
                                            colValue = "'" + dateDDMMYYY + "'";
                                        else
                                            colValue = colValue + ",'" + dateDDMMYYY + "'";

                                    }
                                    catch (Exception ex)
                                    {
                                        if (colValue == "")
                                            colValue = "null";
                                        else
                                            colValue = colValue + ",null";
                                    }
                                }
                                else
                                {
                                    if (colValue == "")
                                        colValue = "null";
                                    else
                                        colValue = colValue + ",null";
                                }

                            }
                            else if (Convert.ToString(dv[0]["Data Type"]).ToLower() == "varchar")
                            {
                                var value = "";
                                if (listBox2.Items[j].ToString().ToLower().Trim() == "prdcode")
                                {
                                    value = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).PadLeft(4, '0');
                                }
                                else if (listBox2.Items[j].ToString().ToLower().Trim() == "branchcode") { value = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).PadLeft(3, '0'); }
                                else value = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).PadLeft(3, '0');
                                if (colValue == "")
                                {
                                    colValue = "'" + value.Replace("'", "").Trim() + "'";
                                }
                                else
                                {
                                    colValue = colValue + ",'" + value.Replace("'", "").Trim() + "'";
                                }
                            }
                            else
                            {
                                if (colValue == "")
                                    colValue = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).Replace("'", "").Trim();
                                else
                                    colValue = colValue + "," + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).Replace("'", "").Trim();

                            }
                            // break;

                        }
                        // }
                    }


                    try
                    {
                        string qury = "insert into " + TableName + " (Created_by,Created_date," + colname + " ) values(1,'" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'," + colValue + " ) select @@identity";
                        string custId = RetureExcuteScaler(qury);
                        if (custId == "")
                        {
                            createlog("Record Number : " + dtExcelrow["SR No"].ToString());
                            createlog("Reason :- Record not inserted.");
                        }

                    }
                    catch (Exception ex)
                    {
                        createlog("Record Number : " + dtExcelrow["SR No"].ToString());
                        createlog("Reason :- Error record.");
                    }
                    if (backgroundWorker1.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }
                    //increament record imported or no imported.
                    w += 1;
                }
            }
            catch (Exception ex)
            {
                // OtherErrorCount = OtherErrorCount + 1;
                createlog("Error occured : " + ex.Message);
            }

            TimeSpan duration = DateTime.Now - start;

            //we could return some useful information here, like calculation output, number of items affected, etc.. to the main thread.  
            e.Result = "Duration: " + duration.TotalMilliseconds.ToString() + " ms.";

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            btoCancel.Enabled = false;
            //btoStart.Enabled = true;
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled"); this.Cursor = Cursors.Default;
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString()); this.Cursor = Cursors.Default;
            }
            else
            {
                progressBar1.Value = progressBar1.Maximum;
                MessageBox.Show("The task has been completed. Results: " + e.Result.ToString());
                this.Cursor = Cursors.Default;
            }
        }
        #endregion

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
            //Application.Exit();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ImageField = comboBox1.Text.Trim();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            MessageBox.Show(comboBox2.Text.Trim());
            strDateFormat = comboBox2.Text.Trim();
        }


        #region ImageEncrypt
        public void ImageEncrypt(string fileIn, string fileOut, string Password)
        {

            // First we are going to open the file streams 

            using (FileStream fsIn = new FileStream(fileIn, FileMode.Open, FileAccess.Read))
            {

                using (FileStream fsOut = new FileStream(fileOut, FileMode.OpenOrCreate, FileAccess.Write))
                {

                    // Then we are going to derive a Key and an IV from the Password and create an algorithm 

                    //PasswordDeriveBytes pdb = new PasswordDeriveBytes(Password,

                    //    new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });

                    Rijndael alg = Rijndael.Create();

                    //alg.Key = pdb.GetBytes(32);

                    //alg.IV = pdb.GetBytes(16);
                    alg.Key = CreateKey(Password);
                    alg.IV = CreateIV(Password);

                    // Now create a crypto stream through which we are going to be pumping data. 

                    // Our fileOut is going to be receiving the encrypted bytes. 

                    using (CryptoStream cs = new CryptoStream(fsOut, alg.CreateEncryptor(), CryptoStreamMode.Write))
                    {

                        // Now will will initialize a buffer and will be processing the input file in chunks. 

                        // This is done to avoid reading the whole file (which can be huge) into memory. 

                        int bufferLen = 4096;

                        byte[] buffer = new byte[bufferLen];

                        int bytesRead;

                        do
                        {

                            // read a chunk of data from the input file 

                            bytesRead = fsIn.Read(buffer, 0, bufferLen);
                            // encrypt it 

                            cs.Write(buffer, 0, bytesRead);

                        } while (bytesRead != 0);
                    }
                }
            }
            // close everything 

            //cs.Close(); // this will also close the unrelying fsOut stream 

            //fsIn.Close();

        }
        #endregion ImageEncrypt

        // Decrypt a file into another file using a password 

        #region ImageDecrypt
        public void ImageDecrypt(string fileIn, string fileOut, string Password)
        {

            // First we are going to open the file streams 

            using (FileStream fsIn = new FileStream(fileIn, FileMode.Open, FileAccess.Read))
            {

                using (FileStream fsOut = new FileStream(fileOut, FileMode.OpenOrCreate, FileAccess.Write))
                {



                    // Then we are going to derive a Key and an IV from the Password and create an algorithm 

                    //PasswordDeriveBytes pdb = new PasswordDeriveBytes(Password,

                    //    new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });



                    Rijndael alg = Rijndael.Create();



                    //alg.Key = pdb.GetBytes(32);

                    //alg.IV = pdb.GetBytes(16);
                    alg.Key = CreateKey(Password);
                    alg.IV = CreateIV(Password);



                    // Now create a crypto stream through which we are going to be pumping data. 

                    // Our fileOut is going to be receiving the Decrypted bytes. 

                    using (CryptoStream cs = new CryptoStream(fsOut, alg.CreateDecryptor(), CryptoStreamMode.Write))
                    {

                        // Now will will initialize a buffer and will be processing the input file in chunks. 

                        // This is done to avoid reading the whole file (which can be huge) into memory. 

                        int bufferLen = 4096;

                        byte[] buffer = new byte[bufferLen];

                        int bytesRead;



                        do
                        {

                            // read a chunk of data from the input file 

                            bytesRead = fsIn.Read(buffer, 0, bufferLen);



                            // Decrypt it 

                            cs.Write(buffer, 0, bytesRead);



                        } while (bytesRead != 0);
                    }

                }
            }

            // close everything 

            // cs.Close(); // this will also close the unrelying fsOut stream 

            //        fsIn.Close();
        }
        #endregion ImageDecrypt

        #region CreateKey
        public byte[] CreateKey(string strPassword)
        {
            byte[] bytKey;
            byte[] bytSalt = System.Text.Encoding.ASCII.GetBytes("salt");
            PasswordDeriveBytes pdb = new PasswordDeriveBytes(strPassword, bytSalt);
            bytKey = pdb.GetBytes(32);
            return bytKey;
            // Return the key.
        }
        #endregion

        #region CreateIV
        public byte[] CreateIV(string strPassword)
        {
            byte[] bytIV;
            byte[] bytSalt = System.Text.Encoding.ASCII.GetBytes("salt");
            PasswordDeriveBytes pdb = new PasswordDeriveBytes(strPassword, bytSalt);
            bytIV = pdb.GetBytes(16);
            return bytIV;
            // Return the IV.
        }
        #endregion
    }
}

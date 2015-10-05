using System;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
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
namespace UploadApplication
{
    public partial class frmUpload : Form
    {
        SqlConnection Sqlconn = new SqlConnection(ConfigurationSettings.AppSettings["SqlCon"].ToString());
        bool ErrorOccured = false;
        int DocTypeID = 0;
        string TableName = "";
        public frmUpload()
        {
            InitializeComponent();
            //mandatory. Otherwise will throw an exception when calling ReportProgress method  
            backgroundWorker1.WorkerReportsProgress = true;

            //mandatory. Otherwise we would get an InvalidOperationException when trying to cancel the operation  
            //backgroundWorker1.WorkerSupportsCancellation = true;
            // This event will be raised on the worker thread when the worker starts
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            // This event will be raised when we call ReportProgress
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
        }
        void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string colname = ""; string colValue = ""; string strrec = "";
            string strImportedBy = "Capture";
            DateTime datImportedOn = System.DateTime.Now;
            DataTable dtDatabaseTable; DataTable dtExceldata;
            int intVersionNo = 1; int intFileStatus = 1; int filenotfoundcount = 0; int OtherErrorCount = 0;
            string strIpAddress = ""; string FileSize = "";
            try
            {
                //dtDatabaseTable = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where docId =" + Convert.ToString(cmbDocumentType.SelectedValue));
                dtDatabaseTable = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where docId =" + Convert.ToString(DocTypeID));

                dtExceldata = GetExcelDataTable("G:\\Projects\\Water Department Doc\\Mine Department", "TradersFileUpdate.xls");
                createlog("###############################  IMPORT LOG FOR Document type  " + cmbDocumentType.SelectedText + "    ############################################");
                createlog("...................................................................................................................................................");
                createlog(" Import started at : " + DateTime.Now);
                createlog("......................................................................................................................................");
                int i = 0;
                foreach (DataRow dtExcelrow in dtExceldata.Rows)
                {
                    int progressPercentage = (100 * i) / dtExceldata.Rows.Count;
                    backgroundWorker1.ReportProgress(progressPercentage);
                    // Simulate long task
                    System.Threading.Thread.Sleep(100);
                    i += i + 1; 
                    colValue = ""; colname = "";
                    ErrorOccured = false;
                    foreach (object o in listBox3.Items)
                    {
                        foreach (DataRow drDatabase in dtDatabaseTable.Rows)
                        {
                            if (drDatabase["label"].ToString() == o.ToString())
                            {
                                if (colname == "")
                                    colname = drDatabase["FieldName"].ToString();
                                else
                                    colname = colname + "," + drDatabase["FieldName"].ToString();
                            }
                        }

                    }
                    //foreach (object colName in listBox2.Items)
                    //{
                    for (int j = 0; j < listBox2.Items.Count; j++)
                    {
                        DataTable dtnew = new DataTable();
                        dtnew = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where label ='" + Convert.ToString(listBox3.Items[j]) + "'");
                        if (Convert.ToString(dtnew.Rows[0]["FieldType"]) == "DateTime")
                        {

                            string dateString, format;
                            DateTime result;
                            CultureInfo provider = CultureInfo.InvariantCulture;
                            dateString = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]); //Input Date
                            format = "dd/MM/yyyy";//// Only change this format as per input date formate 
                            result = DateTime.ParseExact(dateString, format, provider);
                            string dateDDMMYYY = result.ToString("MM/dd/yyyy");
                            dateDDMMYYY = dateDDMMYYY.Replace("-", "/");

                            if (colValue == "")
                                colValue = "'" + dateDDMMYYY + "'";
                            else
                                colValue = colValue + ",'" + dateDDMMYYY + "'";

                        }
                        else
                        {
                            if (colValue == "")
                                colValue = "'" + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]) + "'";
                            else
                                colValue = colValue + ",'" + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]) + "'";

                        }


                    }



                    //}
                    string strFileName = Convert.ToString(dtExcelrow["IMAGE 1"]);
                    string strFilePath = Path.GetDirectoryName(txtXLSFileName.Text);
                    string strFullFilePath = strFilePath + "\\" + strFileName;
                    FileInfo f = new FileInfo(strFullFilePath);
                    if (f.Exists)
                    {
                        FileInfo fInfo = new FileInfo(strFullFilePath);
                        if (fInfo.Length >= 1024)
                        {
                            Int64 size = (f.Length) / 1024;
                            FileSize = size + " KB";
                        }
                        else
                        {
                            FileSize = fInfo.Length.ToString() + " Bytes";
                        }
                        FileStream fs = new FileStream(strFullFilePath, FileMode.Open, FileAccess.Read);
                        BinaryReader br = new BinaryReader(fs);
                        byte[] imgFileImage = br.ReadBytes((int)fs.Length);
                        Int32 PageCount = 1;
                        if (Path.GetExtension(strFileName).ToLower() == ".tiff" || Path.GetExtension(strFileName).ToLower() == ".tif")
                        {
                            PageCount = FramCount(strFullFilePath); //dc.GetImagePageCount(FileUpload1);
                            if (PageCount == -1)
                                PageCount = 1;
                        }
                        InsertImageDetails(strFileName, strFilePath, imgFileImage, Convert.ToInt32(DocTypeID), strImportedBy, datImportedOn, FileSize, intVersionNo, intFileStatus, "", strIpAddress, TableName, PageCount, colname, colValue);
                        if (ErrorOccured)
                        {
                            OtherErrorCount = OtherErrorCount + 1;
                            strrec = "       Record Details : ";
                            foreach (DataColumn cc in dtExceldata.Columns)
                            {
                                strrec += cc.ColumnName + " = ";
                                strrec += dtExcelrow[cc.ColumnName].ToString() + " , ";
                            }
                            createlog(strrec);
                        }
                    }
                    else
                    {
                        filenotfoundcount = filenotfoundcount + 1;
                        createlog("......................................................................................................................................");
                        createlog("       File not found on path = " + strFullFilePath + "  , hence file and data is not imported ");
                        strrec = "        Record Details : ";
                        foreach (DataColumn cc in dtExceldata.Columns)
                        {
                            strrec += cc.ColumnName + " = ";
                            strrec += dtExcelrow[cc.ColumnName].ToString() + " , ";
                        }
                        createlog(strrec);
                        createlog("......................................................................................................................................");
                    }

                   
                }
               
            }
            catch (Exception ex)
            {
                OtherErrorCount = OtherErrorCount + 1;
                createlog("Error occured : " + ex.Message);

            }


          
        }
        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progress = e.ProgressPercentage;
            int maxProgress = progressBar1.Maximum;
            progressBar1.Value = Math.Min(progress, maxProgress);

            
        }
        private void frmUpload_Load(object sender, EventArgs e)
        {

            this.Size = new Size(483, 160);
            this.groupBox1.Size = new Size(450, 155);

            btnImportData.Visible = false;
            btnClear.Visible = false;

            btnBrowse.Focus();

            DataTable dt = GetDataTable("select docLabel,DocId from tblDocmaster");
            if (dt.Rows.Count > 0)
            {
                cmbDocumentType.DisplayMember = "docLabel";
                cmbDocumentType.ValueMember = "DocId";
                cmbDocumentType.DataSource = dt;

                DataTable dt1 = GetDataTable(string.Format("select fieldid,label from tbldocdetail  where docid={0}", Convert.ToString(dt.Rows[0][1])));
                LoadListBox(dt1, lstDatabaseColumn, "row");
            }
            else
            {
                MessageBox.Show("Create document type.");
                this.Close();
            }
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

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //FolderBrowserDialog fldBrowser = new FolderBrowserDialog();
            OpenFileDialog oldBrowser = new OpenFileDialog();
            oldBrowser.ShowDialog();

            txtXLSFileName.Text = oldBrowser.FileName;

            this.Cursor = Cursors.Default;
        }

        private void btnGetXLSColumn_Click(object sender, EventArgs e)
        {
            if (txtXLSFileName.Text.Length == 0)
            {
                MessageBox.Show("Select xls file to import data.");
                btnBrowse.Focus();
                return;
            }
            DataTable dt = GetExcelDataTable(Path.GetDirectoryName(txtXLSFileName.Text), Path.GetFileName(txtXLSFileName.Text));
            if (dt.Rows.Count > 0)
            {
                LoadListBox(dt, lstXLSColumn, "column");
                tableLayoutPanel3.Visible = true;
                btnImportData.Visible = true;
                btnClear.Visible = true;

                this.Size = new Size(483, 450);
                this.groupBox1.Size = new Size(450, 319);
            }
            else
            {
                this.Size = new Size(483, 160);
                this.groupBox1.Size = new Size(450, 155);

                tableLayoutPanel3.Visible = false;
                btnImportData.Visible = false;
                btnClear.Visible = false;
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
                lstDatabaseColumn.Items.Add(listBox3.SelectedItem);
                listBox3.Items.Remove(listBox3.SelectedItem);
            }
            else { MessageBox.Show("Select column name."); }
        }

        private void btnImportData_Click(object sender, EventArgs e)
        {
            //string colname = ""; string colValue = ""; string strrec = "";
            //string strImportedBy = "Capture";
            //DateTime datImportedOn = System.DateTime.Now;
            //DataTable dtDatabaseTable; DataTable dtExceldata;
            //int intVersionNo = 1; int intFileStatus = 1; int filenotfoundcount = 0; int OtherErrorCount = 0;
            //string strIpAddress = ""; string FileSize = "";
            //try
            //{
            //    dtDatabaseTable = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where docId =" + Convert.ToString(cmbDocumentType.SelectedValue));
            //    dtExceldata = GetExcelDataTable(Path.GetDirectoryName(txtXLSFileName.Text), Path.GetFileName(txtXLSFileName.Text));
            //    createlog("###############################  IMPORT LOG FOR Document type  " + cmbDocumentType.SelectedText + "    ############################################");
            //    createlog("...................................................................................................................................................");
            //    createlog(" Import started at : " + DateTime.Now);
            //    createlog("......................................................................................................................................");
            //    foreach (DataRow dtExcelrow in dtExceldata.Rows)
            //    {
            //        colValue = ""; colname = "";
            //        ErrorOccured = false;
            //        foreach (object o in listBox3.Items)
            //        {
            //            foreach (DataRow drDatabase in dtDatabaseTable.Rows)
            //            {
            //                if (drDatabase["label"].ToString() == o.ToString())
            //                {
            //                    if (colname == "")
            //                        colname = drDatabase["FieldName"].ToString();
            //                    else
            //                        colname = colname + "," + drDatabase["FieldName"].ToString();
            //                }
            //            }

            //        }
            //        //foreach (object colName in listBox2.Items)
            //        //{
            //        for (int j = 0; j < listBox2.Items.Count; j++)
            //        {
            //            DataTable dtnew = new DataTable();
            //            dtnew = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where label ='" + Convert.ToString(listBox3.Items[j]) + "'");
            //            if (Convert.ToString(dtnew.Rows[0]["FieldType"]) == "DateTime")
            //            {

            //                string dateString, format;
            //                DateTime result;
            //                CultureInfo provider = CultureInfo.InvariantCulture;
            //                dateString = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]); //MM/DD/YYYY
            //                format = "dd/MM/yyyy";//"d";
            //                result = DateTime.ParseExact(dateString, format, provider);
            //                string dateDDMMYYY = result.ToString("MM/dd/yyyy"); 
            //                dateDDMMYYY = dateDDMMYYY.Replace("-", "/");

            //                if (colValue == "")
            //                    colValue = "'" + dateDDMMYYY + "'";
            //                else
            //                    colValue = colValue + ",'" + dateDDMMYYY + "'";

            //            }
            //            else
            //            {
            //                if (colValue == "")
            //                    colValue = "'" + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]) + "'";
            //                else
            //                    colValue = colValue + ",'" + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]) + "'";

            //            }


            //        }



            //        //}
            //        string strFileName = Convert.ToString(dtExcelrow["IMAGE 1"]);
            //        string strFilePath = Path.GetDirectoryName(txtXLSFileName.Text);
            //        string strFullFilePath = strFilePath + "\\" + strFileName;
            //        FileInfo f = new FileInfo(strFullFilePath);
            //        if (f.Exists)
            //        {
            //            FileInfo fInfo = new FileInfo(strFullFilePath);
            //            if (fInfo.Length >= 1024)
            //            {
            //                Int64 size = (f.Length) / 1024;
            //                FileSize = size + " KB";
            //            }
            //            else
            //            {
            //                FileSize = fInfo.Length.ToString() + " Bytes";
            //            }
            //            FileStream fs = new FileStream(strFullFilePath, FileMode.Open, FileAccess.Read);
            //            BinaryReader br = new BinaryReader(fs);
            //            byte[] imgFileImage = br.ReadBytes((int)fs.Length);
            //            Int32 PageCount = 1;
            //            if (Path.GetExtension(strFileName).ToLower() == ".tiff" || Path.GetExtension(strFileName).ToLower() == ".tif")
            //            {
            //                PageCount = FramCount(strFullFilePath); //dc.GetImagePageCount(FileUpload1);
            //                if (PageCount == -1)
            //                    PageCount = 1;
            //            }
            //            InsertImageDetails(strFileName, strFilePath, imgFileImage, Convert.ToInt32(cmbDocumentType.SelectedValue), strImportedBy, datImportedOn, FileSize, intVersionNo, intFileStatus, "", strIpAddress, cmbDocumentType.SelectedText, PageCount, colname, colValue);
            //            if (ErrorOccured)
            //            {
            //                OtherErrorCount = OtherErrorCount + 1;
            //                strrec = "       Record Details : ";
            //                foreach (DataColumn cc in dtExceldata.Columns)
            //                {
            //                    strrec += cc.ColumnName + " = ";
            //                    strrec += dtExcelrow[cc.ColumnName].ToString() + " , ";
            //                }
            //                createlog(strrec);
            //            }
            //        }
            //        else
            //        {
            //            filenotfoundcount = filenotfoundcount + 1;
            //            createlog("......................................................................................................................................");
            //            createlog("       File not found on path = " + strFullFilePath + "  , hence file and data is not imported ");
            //            strrec = "        Record Details : ";
            //            foreach (DataColumn cc in dtExceldata.Columns)
            //            {
            //                strrec += cc.ColumnName + " = ";
            //                strrec += dtExcelrow[cc.ColumnName].ToString() + " , ";
            //            }
            //            createlog(strrec);
            //            createlog("......................................................................................................................................");
            //        }

            //    }
            //}
            //catch (Exception ex)
            //{
            //    OtherErrorCount = OtherErrorCount + 1;
            //    createlog("Error occured : " + ex.Message);

            //}

            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Done");
        }

        private void InsertImageDetails(string strFileName, string strFilePath, byte[] imgFileImage,
                                Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo,
                                Int32 intFileStatus, string strInitialPath, string strIpAddress, string TableName, Int32 PageCount, string columnName, string ColumnValue)
        {

            string strDate = System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");
            if (strDate.IndexOf("-") != -1)
                datImportedOn = DateTime.ParseExact(strDate, "MM-dd-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            else
                datImportedOn = DateTime.ParseExact(strDate, "MM/dd/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            ErrorOccured = false;
            DataTable dtGetTableNAme = GetDataTable("select docname from tbldocmaster where docid=" + intDocId);
            TableName = Convert.ToString(dtGetTableNAme.Rows[0]["docname"]);
            string DocInsert = InsertDocument(strFileName, strFilePath, imgFileImage, intDocId, strImportedBy, datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);
            if (DocInsert == "true")
            {
                DataTable maxidds = GetDataTable("select MAX(Id) from tblFile");
                string maxid = maxidds.Rows[0][0].ToString();
                if (ColumnValue != "")
                {
                    string qury = "";
                    if (strDate.IndexOf("-") != -1)
                        qury = "insert into " + TableName + " (File_Id,ImportedBy,ImportedOn," + columnName + " ) values(" + maxid + ",'Capture','" + datImportedOn.ToString() + "'," + ColumnValue + " )";
                    else
                        qury = "insert into " + TableName + " (File_Id,ImportedBy,ImportedOn," + columnName + " ) values(" + maxid + ",'Capture','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'," + ColumnValue + " )";
                    if (dataUpdate(qury) == true)
                    {
                        updatedata("Insert into TblFileActions(FileId,DocId,FileName,TransProcess,UserName,TransDate) values(" + maxid + "," + intDocId + ",'" + strFileName + "','Uploaded','Capture','" + datImportedOn + "')");
                    }
                    else
                    {
                        ErrorOccured = true;
                        createlog("Error occured while insert in " + TableName + " document table : " + qury);
                    }
                }
                else
                {

                }
            }
            else
            {

            }

        }
      
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
                // prmFileImage.SqlDbType = SqlDbType.Image;
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
                //string strDate = System.DateTime.Now.ToString("MM/dd/yyyy hh:MM:ss tt");
                //if (strDate.IndexOf("-") != -1)
                //    prmImportedOn.Value = DateTime.ParseExact(strDate, "MM-dd-yyyy hh:MM:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
                //else
                //    prmImportedOn.Value = DateTime.ParseExact(strDate, "MM/dd/yyyy hh:MM:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);

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
            DataTable dtExcelData = new DataTable();

            string _ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + System.IO.Path.Combine(Path, FileName) + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection objConn = new OleDbConnection(_ConnectionString);
            objConn.Open();

            DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt.Rows.Count > 0)
            {
                OleDbCommand ocmd = new OleDbCommand("SELECT * FROM [" + dt.Rows[0]["TABLE_NAME"].ToString() + "]", objConn);
                OleDbDataReader odr = ocmd.ExecuteReader();
                dtExcelData.Load(odr);
            }
            else
                MessageBox.Show("Excel not read, please try again.");

            return dtExcelData;
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
            string  retStr="";
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
                foreach (DataColumn col in dt.Columns)
                {
                    ListBox.Items.Add(col.ColumnName);
                }
            }
        }

        private void cmbDocumentType_SelectedIndexChanged(object sender, EventArgs e)
        {
            DocTypeID = Convert.ToInt32(cmbDocumentType.SelectedValue);
            TableName = RetureExcuteScaler("select DocNAme from tbldocmaster where docid= " + DocTypeID);
            DataTable dt = GetDataTable(string.Format("select fieldid,label from tbldocdetail  where docid={0}", Convert.ToString(cmbDocumentType.SelectedValue)));
                       LoadListBox(dt, lstDatabaseColumn, "row");
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
            tableLayoutPanel3.Visible = false;
            btnImportData.Visible = false;
            btnClear.Visible = false;

            //this.Size = new Point(483, 369);
            //this.groupBox1.Size = new Size(450, 319);
            this.Size = new Size(483, 160);
            this.groupBox1.Size = new Size(450, 155);

        }
    }
}

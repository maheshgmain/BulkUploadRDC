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
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Runtime.InteropServices;

namespace UploadApplication
{
    public partial class Form3 : Form
    {
        SqlConnection Sqlconn = new SqlConnection(ConfigurationSettings.AppSettings["SqlCon"].ToString());

        bool ErrorOccured = false;
        int DocTypeID = 0;
        string TableName = "", colname = "";
        DataTable dtExceldata = null;
        DataTable dtDocDetails = null;
        int w = 0;
        string ImageField = "", strDateFormat = "";
        public Form3()
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

            DataTable dt = GetDataTable("select docLabel,DocId from tblDocmaster");
            if (dt.Rows.Count > 0)
            {
                cmbDocumentType.DisplayMember = "docLabel";
                cmbDocumentType.ValueMember = "DocId";
                cmbDocumentType.DataSource = dt;

                dtDocDetails = GetDataTable(string.Format("select fieldid,label,FieldName,fieldtype from tbldocdetail  where docid={0}", Convert.ToString(dt.Rows[0][1])));
                LoadListBox(dtDocDetails, lstDatabaseColumn, "row");

                btnBrowse.Focus();
            }
            else
            {
                MessageBox.Show("Create document type.");
                this.Close();
            }


            //bind format of date in excel
            // CultureInfo ci = CultureInfo.InvariantCulture;
            CultureInfo ci = CultureInfo.GetCultureInfo("en-gb");
            CultureInfo ci1 = CultureInfo.GetCultureInfo("en-us");
            string[] fmts = ci.DateTimeFormat.GetAllDateTimePatterns();
            string[] fmts1 = ci1.DateTimeFormat.GetAllDateTimePatterns();
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
                createlog("Excel not read, please change excel.");
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
            // DocTypeID = 0;a
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
                    if (comboBox1.SelectedIndex != 0)
                    {
                        btoStart.Enabled = false;
                        btoCancel.Enabled = true;
                        backgroundWorker1.RunWorkerAsync();
                    }
                    else
                    {
                        MessageBox.Show("Please select image field Name");
                        comboBox1.Focus();
                    }
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

        private void cmbDocumentType_SelectedIndexChanged(object sender, EventArgs e)
        {
            //set doc type id
            DocTypeID = Convert.ToInt32(cmbDocumentType.SelectedValue);
            //set selected doctype table name
            TableName = RetureExcuteScaler("select DocName from tbldocmaster where docid= " + DocTypeID);
            //get columns name as per selected doctype
            // DataTable dt = GetDataTable(string.Format("select fieldid,label from tbldocdetail  where docid={0}", Convert.ToString(cmbDocumentType.SelectedValue)));
            //bind columns name to listbox
            //LoadListBox(dt, lstDatabaseColumn, "row");

            dtDocDetails = GetDataTable(string.Format("select fieldid,label,FieldName,fieldtype from tbldocdetail  where docid={0}", Convert.ToString(cmbDocumentType.SelectedValue)));
            LoadListBox(dtDocDetails, lstDatabaseColumn, "row");

        }

        private Int32 FramCount_new(string ImagePath)
        {
            Int32 frmCount;
            try
            {
                if (Path.GetExtension(ImagePath.ToLower()) != ".pdf")
                {
                    using (System.Drawing.Image img = System.Drawing.Image.FromFile(ImagePath))
                    {
                        frmCount = img.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);
                    }
                }
                else frmCount = pdfPageCount(ImagePath);
            }
            catch
            {
                return -1;
                createlog("Image frame count not found.");
            }
            return frmCount;
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
        private void InsertImageDetails_old(string strFileName, string strFilePath, byte[] imgFileImage,
                        Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo,
                        Int32 intFileStatus, string strInitialPath, string strIpAddress, string TableName,
            Int32 PageCount, string columnName, string ColumnValue, string strFullFilePath)
        {

            string strDate = System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
            if (strDate.IndexOf("-") != -1)
                datImportedOn = DateTime.ParseExact(strDate, "dd-MM-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            else
                datImportedOn = DateTime.ParseExact(strDate, "dd/MM/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);


            ErrorOccured = false;
            //insert data into tblfile

            string Folder = ConfigurationManager.AppSettings["Repository"].ToString() + TableName;
            string strtime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + DateTime.Now.Millisecond.ToString("000");
            if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
            {
                strFilePath = "~\\Repository\\Encrypt\\" + TableName + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc";
                //  ViewState["Folder"] = Folder;
                if (!Directory.Exists(Folder))
                {
                    Directory.CreateDirectory(Folder);
                }
                imgFileImage = new byte[1];
            }

            string DocInsert = InsertDocument(strFileName, strFilePath, imgFileImage, intDocId, strImportedBy, datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);

            if (DocInsert == "true")
            {
                if (ColumnValue != "")
                {
                    //get id of last record inserted
                    DataTable maxidds = GetDataTable("select MAX(Id) from tblFile");
                    string maxid = maxidds.Rows[0][0].ToString();

                    //insert into index table
                    string qury = "insert into " + TableName + " (File_Id,ImportedBy,ImportedOn," + columnName + " ) values(" + maxid + "," + strImportedBy + ",'" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'," + ColumnValue + " )";
                    if (dataUpdate(qury) == true)
                    {
                        //insert into tblfileactions table
                        updatedata("Insert into TblFileActions(FileId,DocId,FileName,TransProcess,UserName,TransDate) values(" + maxid + "," + intDocId + ",'" + strFileName + "','Uploaded'," + strImportedBy + ",'" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "')");

                        //save image
                        if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
                            ImageEncrypt(strFullFilePath, Folder + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc", "encrypt123");
                    }
                    else
                    {
                        ErrorOccured = true;
                        //if error occured then delete inserted record from tblfile
                        updatedata("delete from tblfile where id=" + maxid + ")");
                        createlog("Error occured while insert in " + TableName + " document table : " + qury);
                    }
                }
            }
        }

        /*Merge Pdf*/
        private void InsertImageDetails(string strFileName, string strFilePath, byte[] imgFileImage,
                Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo,
                Int32 intFileStatus, string strInitialPath, string strIpAddress, string TableName,
    Int32 PageCount, string columnName, string ColumnValue, string strFullFilePath)
        {

            // byte[] imgFileImage = null;
            // string strFileSize = "", strFileName = "", strFilePath = "";
            //Int32 PageCount = 0, intFileStatus = 1;

            string strDate = System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
            if (strDate.IndexOf("-") != -1)
                datImportedOn = DateTime.ParseExact(strDate, "dd-MM-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            else
                datImportedOn = DateTime.ParseExact(strDate, "dd/MM/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);


            ErrorOccured = false;
            //insert data into tblfile

            string Folder = ConfigurationManager.AppSettings["Repository"].ToString() + TableName;
            string strtime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + DateTime.Now.Millisecond.ToString("000");
            string strFilepathNew = "";
            // if (ConfigurationManager.AppSettings["Repository"].ToString() != "")
            if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
            {
                strFilepathNew = "~\\Repository\\Encrypt\\" + TableName + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc";
                //  ViewState["Folder"] = Folder;
                if (!Directory.Exists(Folder))
                {
                    Directory.CreateDirectory(Folder);
                }
                imgFileImage = new byte[1];
            }

            // string DocInsert = InsertDocument(strFileName, strFilePath, imgFileImage, intDocId, strImportedBy, datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);
            Int32 FileID = InsertDocumentReturnIdentity(strFileName, strFilePath, imgFileImage, intDocId, strImportedBy, datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);
            if (FileID > 0)
            {
                if (ColumnValue != "")
                {
                    //insert into index table
                    string qury = "insert into " + TableName + " (File_Id,ImportedBy,ImportedOn," + columnName + " ) values(" + FileID + ",'" + strImportedBy + "','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'," + ColumnValue + " ) select @@identity";
                    Int32 doctypeId = Convert.ToInt32(RetureExcuteScaler(qury));
                    if (doctypeId > 0)
                    {
                        //get merge image
                        strFilePath = MergeImages(Convert.ToString(FileID), TableName, intDocId);
                        if (strFilePath != "" && strFilePath != "Empty")
                        {
                            strFileName = Path.GetFileName(strFilePath); //Path.GetFileName(HttpContext.Current.Session["FileName"].ToString());
                            using (FileStream fs = new FileStream(strFilePath, FileMode.Open, FileAccess.Read))
                            {
                                BinaryReader br = new BinaryReader(fs);
                                imgFileImage = br.ReadBytes((int)fs.Length);
                            }
                            //get page count   
                            PageCount = FramCount_new(strFilePath);

                            if (imgFileImage.Length >= 1024)//(FileUpload1.FileContent.Length >= 1024)
                            {
                                Int64 size = (imgFileImage.Length) / 1024;
                                strFileSize = size + " KB";
                            }
                            else
                            {
                                strFileSize = imgFileImage.Length.ToString() + " Bytes";
                            }

                            //set status as per user check checkbox of pending or done
                            //if (ConfigurationManager.AppSettings["IsWorkFlow"].ToString().ToLower() == "yes")
                            //    intFileStatus = 0;
                            //else
                            intFileStatus = 5;

                            //save image
                            if (ConfigurationManager.AppSettings["Repository"].ToString() != "")
                                if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
                                    ImageEncrypt(strFilePath, Folder + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc", "encrypt123");
                            try
                            {
                                if (ConfigurationManager.AppSettings["Repository"].ToString() != "")
                                    imgFileImage = new byte[1];

                                UpdateImage(FileID, imgFileImage, Path.GetDirectoryName(strFilepathNew) + @"\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc", intFileStatus, PageCount, strFileSize, strFileName);
                            }
                            catch (Exception)
                            {
                                createlog("error in updating tblfile.--)" + FileID); ;
                            }
                            //insert into tblfileactions table
                            try
                            {


                                updatedata("Insert into TblFileActions(FileId,DocId,FileName,TransProcess,UserName,TransDate) values(" + FileID + "," + intDocId + ",'" + strFileName + "','Uploaded','" + strImportedBy + "','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "')");
                            }
                            catch (Exception)
                            {
                                createlog("error in inserting tblfileactiondetails.--)" + FileID); ;
                            }
                            //update section
                            UpdateSectionDetails(intDocId, FileID, TableName, doctypeId);


                        }
                        else
                        {
                            if (strFilePath == "")
                                createlog("Merger image not created." + qury);
                            // updatedata("delete from tblfile where id=" + FileID);
                            // updatedata("delete from " + TableName + " where file_id=" + FileID );
                        }
                    }
                    else
                    {
                        ErrorOccured = true;
                        //if error occured then delete inserted record from tblfile
                        updatedata("delete from tblfile where id=" + FileID);
                        createlog("Error occured while insert in " + TableName + " document table : " + qury);
                    }
                }
            }
            else
            {
                createlog("Image not insetted into tblfile.");
                ErrorOccured = true;
            }
        }


        //merge image if multiple selected
        private string MergeImages(string fileName, string TableName, Int32 docid)
        {
            string destination = "", contentType = "";
            string folderName = RetureExcuteScaler(string.Format(@"select ltrim(rtrim(column4)) from {0} where file_id={1}", TableName, fileName));

            try
            {
                if (folderName == "")
                {
                    createlog("Folder name empty.");
                    return "";
                }

                // ArrayList strarrayList =// (ArrayList)HttpContext.Current.Session["MultipleFiles"];

                DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(txtXLSFileName.Text) + @"/" + folderName);
                FileInfo[] fileInfo = dir.GetFiles();
                if (fileInfo.Length == 0)
                {
                    createlog("Folder is empty.--fileid" + fileName + " foldername:-" + folderName); return "Empty";
                }
                //fileInfo = dir.GetFiles("*.jpg");
                //if (fileInfo.Length > 0)
                //{
                //    convertJPG2PDF("payment receipt_new.pdf", dir.FullName);

                //}

                fileInfo = dir.GetFiles("*.pdf");
                string sqlquery = string.Format("select distinct tblsectionmaster.sec_id,doclabel docname,sec_name, 0 pagenofrom,0 pagenoto from tblsectionmaster,tbldocmaster where tblsectionmaster.doc_id=tbldocmaster.docid and tblsectionmaster.doc_id={0}", docid);
                DataTable dtsection = GetDataTable(sqlquery);
                if (dtsection == null)
                {
                    if (fileInfo.Length == 1)
                        return fileInfo[0].FullName;
                    else
                        destination = fileInfo[0].FullName;

                    for (Int32 i = 1; i < fileInfo.Length; i++)
                    {
                        if (Path.GetExtension(fileInfo[i].FullName).ToLower() == ".pdf")
                        {
                            // check whether or not the uploaded file is an image:
                            contentType = getFiletype.GetContentType(fileInfo[i].FullName);
                            if (contentType.Contains("pdf"))
                            {
                                destination = ADDPDF2PDF(destination, fileInfo[i].FullName, fileName);
                            }
                        }

                    }
                    return destination;
                }
                if (dtsection.Rows.Count == 0)
                {
                    if (fileInfo.Length == 1)
                        return fileInfo[0].FullName;
                    else
                        destination = fileInfo[0].FullName;

                    for (Int32 i = 1; i < fileInfo.Length; i++)
                    {
                        if (Path.GetExtension(fileInfo[i].FullName).ToLower() == ".pdf")
                        {
                            // check whether or not the uploaded file is an image:
                            contentType = getFiletype.GetContentType(fileInfo[i].FullName);
                            if (contentType.Contains("pdf"))
                            {
                                destination = ADDPDF2PDF(destination, fileInfo[i].FullName, fileName);
                            }
                        }
                        //destination = ADDPDF2PDF(destination, fileInfo[i].FullName, fileName);
                        //else
                        //{}+
                    }

                }
                else
                {
                    //if (fileInfo.Length == 1)
                    //    return fileInfo[0].FullName;

                    for (int i = 0; i < dtsection.Rows.Count; i++)
                    {
                        mergerSectionwiseImage(dtsection.Rows[i]["sec_name"].ToString().Trim().ToLower(), dir.FullName);
                        fileInfo = dir.GetFiles("*.pdf");
                        for (Int32 x = 0; x < fileInfo.Length; x++)
                        {
                            if (Convert.ToString(dtsection.Rows[i]["sec_name"]).ToLower().Trim()
                                == Path.GetFileNameWithoutExtension(fileInfo[x].FullName).ToString().ToLower().Trim())
                            {
                                if (Path.GetExtension(fileInfo[x].FullName).ToLower()
                                    == ".pdf")
                                {
                                    if (destination == "")
                                    {
                                        contentType = getFiletype.GetContentType(fileInfo[x].FullName);
                                        if (contentType.Contains("pdf"))
                                        {
                                            destination = fileInfo[x].FullName;
                                        }
                                        else if (contentType.Contains("jpeg"))
                                        {
                                            convertJPG(Path.GetDirectoryName(fileInfo[x].FullName) + Path.DirectorySeparatorChar + Convert.ToString(dtsection.Rows[i]["sec_name"]).ToLower().Trim() + "_" + i.ToString() + ".pdf", fileInfo[x].FullName);
                                            FileInfo f2 = new FileInfo(Path.GetDirectoryName(fileInfo[x].FullName) + Path.DirectorySeparatorChar + Convert.ToString(dtsection.Rows[i]["sec_name"]).ToLower().Trim() + "_" + i + ".pdf");
                                            if (f2.Exists)
                                            {
                                                f2.CopyTo(Path.GetDirectoryName(fileInfo[x].FullName) + Path.DirectorySeparatorChar + Convert.ToString(dtsection.Rows[i]["sec_name"]).ToLower().Trim() + ".pdf");
                                                f2.Delete();
                                                destination = Path.GetDirectoryName(f2.FullName) + Path.DirectorySeparatorChar + Convert.ToString(dtsection.Rows[i]["sec_name"]).ToLower().Trim() + ".pdf";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // check whether or not the uploaded file is an image:
                                        contentType = getFiletype.GetContentType(fileInfo[x].FullName);
                                        if (contentType.Contains("pdf"))
                                        {
                                            destination = ADDPDF2PDF(destination, fileInfo[x].FullName, fileName);
                                        }
                                    }
                                    break;
                                }
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                createlog("Error in merging methode.--" + folderName);
                return "";
            }
            return destination;
        }
        //add to pdf fi
        private string ADDPDF2PDF(string source, string broweFile, string fileName)
        {
            string destinaton = "";
            try
            {

                //main fikle
                PdfDocument outputDocument = new PdfDocument();

                //open the main document in which to add pages
                PdfDocument mainDocument = Open(source);
                // Open the document to import pages from it.
                PdfDocument inputDocument = Open(broweFile);
                // Iterate pages

                int mainCount = mainDocument.PageCount;
                int Inputcount = inputDocument.PageCount;

                int pageNum = 0;
                //first add source file pages in new document
                for (pageNum = 0; pageNum < mainCount; pageNum++)
                {
                    PdfPage page = mainDocument.Pages[pageNum]; ;
                    outputDocument.Pages.Add(page);
                }
                //then add browsefile pages into new document
                for (pageNum = 0; pageNum < Inputcount; pageNum++)
                {
                    PdfPage page = inputDocument.Pages[pageNum]; ;
                    outputDocument.Pages.Add(page);
                    //pageNumber++;
                }
                //set destinationPath
                //destinaton = Path.GetDirectoryName(source) + "//" + Path.GetFileNameWithoutExtension(source.Replace(" ", "")) + "_newOutput.pdf";
                destinaton = Path.GetDirectoryName(source) + "\\" + fileName + ".pdf";
                outputDocument.Save(destinaton);

                //close pdf document
                outputDocument.Close();
                mainDocument.Close();
                inputDocument.Close();
            }
            catch (Exception ex)
            {
                return "";
            }
            return destinaton;
        }

        //update merger image to database
        private bool UpdateImage(long FileID, byte[] bytes, string FilePath, Int32 intFileStatus,
    Int32 PageCount, string FileSize, string FileName)
        {
            bool temp = false;
            try
            {

                if (Sqlconn.State == ConnectionState.Closed)
                {
                    Sqlconn.Open();
                }
                SqlCommand com = new SqlCommand();
                com.Connection = Sqlconn;

                com.CommandType = CommandType.Text;
                com.CommandText = "update  tblFile set FileImage = @FileImage,FilePath=@FilePath,PageCount=@PageCount,FileSize=@FileSize,FileName=@FileName,filestatus=@filestatus where Id =@FileId";

                SqlParameter sqlFileImage = new SqlParameter("@FileImage", SqlDbType.VarBinary);
                sqlFileImage.Direction = ParameterDirection.Input;
                sqlFileImage.Value = bytes;
                com.Parameters.Add(sqlFileImage);


                SqlParameter sqlFilePath = new SqlParameter("@FilePath", SqlDbType.VarChar);
                sqlFilePath.Direction = ParameterDirection.Input;
                sqlFilePath.Value = FilePath;
                com.Parameters.Add(sqlFilePath);

                SqlParameter sqlFileID = new SqlParameter("@FileId", SqlDbType.BigInt);
                sqlFileID.Direction = ParameterDirection.Input;
                sqlFileID.Value = FileID;
                com.Parameters.Add(sqlFileID);


                SqlParameter sqlPageCount = new SqlParameter("@PageCount", SqlDbType.Int);
                sqlPageCount.Direction = ParameterDirection.Input;
                sqlPageCount.Value = PageCount;
                com.Parameters.Add(sqlPageCount);

                SqlParameter sqlFileSize = new SqlParameter("@FileSize", SqlDbType.VarChar);
                sqlFileSize.Direction = ParameterDirection.Input;
                sqlFileSize.Value = FileSize;
                com.Parameters.Add(sqlFileSize);


                SqlParameter sqlFileName = new SqlParameter("@FileName", SqlDbType.VarChar);
                sqlFileName.Direction = ParameterDirection.Input;
                sqlFileName.Value = FileName;
                com.Parameters.Add(sqlFileName);

                SqlParameter sqlIsuploaded = new SqlParameter("@filestatus", SqlDbType.Int);
                sqlIsuploaded.Direction = ParameterDirection.Input;
                sqlIsuploaded.Value = intFileStatus;
                com.Parameters.Add(sqlIsuploaded);

                com.ExecuteNonQuery();
                temp = true;

            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                Sqlconn.Close();
            }
            return temp;

        }

        //private Int32 pdfPageCount(string broweFile)
        //{// Open the document to import pages from it.
        //    PdfDocument inputDocument = PdfReader.Open(broweFile, PdfDocumentOpenMode.Import);
        //    // Iterate pages
        //    int pageCount = inputDocument.PageCount;
        //    inputDocument.Close();
        //    return pageCount;

        //}

        public static Int32 pdfPageCount(string broweFile)
        {
            int pageCount = 0;
            // Open the document to import pages from it.
            try
            {
                PdfDocument inputDocument = Open(broweFile);//PdfReader.Open(broweFile, PdfDocumentOpenMode.Import);
                // Iterate pages
                pageCount = inputDocument.PageCount;
                inputDocument.Close();

            }
            catch (Exception ex)
            {
                return -1;

                //           sourceStream.Position = 0;
                //MemoryStream outputStream = new MemoryStream();
                //iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(sourceStream);
                //iTextSharp.text.pdf.PdfStamper pdfStamper = new iTextSharp.text.pdf.PdfStamper(reader, outputStream);
                //pdfStamper.FormFlattening = true;
                //pdfStamper.Writer.SetPdfVersion(iTextSharp.text.pdf.PdfWriter.PDF_VERSION_1_4);
                //pdfStamper.Writer.CloseStream = false;
                //pdfStamper.Close();

                //outDoc = PdfReader.Open(outputStream, PdfDocumentOpenMode.Import);

            }
            return pageCount;

        }

        /// <summary>
        /// uses itextsharp 4.1.6 to convert any pdf to 1.4 compatible pdf, called instead of PdfReader.open
        /// </summary>
        static public PdfDocument Open(string PdfPath)
        {
            using (FileStream fileStream = new FileStream(PdfPath, FileMode.Open, FileAccess.Read))
            {
                int len = (int)fileStream.Length;
                Byte[] fileArray = new Byte[len];
                fileStream.Read(fileArray, 0, len);
                fileStream.Close();

                return Open(fileArray);
            }
        }

        /// <summary>
        /// uses itextsharp 4.1.6 to convert any pdf to 1.4 compatible pdf, called instead of PdfReader.open
        /// </summary>
        static public PdfDocument Open(byte[] fileArray)
        {
            return Open(new MemoryStream(fileArray));
        }

        /// <summary>
        /// uses itextsharp 4.1.6 to convert any pdf to 1.4 compatible pdf, called instead of PdfReader.open
        /// </summary>
        static public PdfDocument Open(MemoryStream sourceStream)
        {
            PdfDocument outDoc = null;
            sourceStream.Position = 0;

            try
            {
                outDoc = PdfReader.Open(sourceStream, PdfDocumentOpenMode.Import);
            }
            catch (PdfSharp.Pdf.IO.PdfReaderException)
            {
                //workaround if pdfsharp doesn't support this pdf
                sourceStream.Position = 0;
                MemoryStream outputStream = new MemoryStream();
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(sourceStream);
                iTextSharp.text.pdf.PdfStamper pdfStamper = new iTextSharp.text.pdf.PdfStamper(reader, outputStream);
                pdfStamper.FormFlattening = true;
                pdfStamper.Writer.SetPdfVersion(iTextSharp.text.pdf.PdfWriter.PDF_VERSION_1_4);
                pdfStamper.Writer.CloseStream = false;
                pdfStamper.Close();

                outDoc = PdfReader.Open(outputStream, PdfDocumentOpenMode.Import);
            }

            return outDoc;
        }


        private bool UpdateSectionDetails(Int32 docid, Int64 fileid,
            string TabelName, Int32 DoctypeID)
        {
            Int32 Issuccess = 0;
            List<Search> ilist = new List<Search>();
            //return if no file r selected


            try
            {

                string sqlquery = string.Format("select distinct tblsectionmaster.sec_id,doclabel docname,sec_name, 0 pagenofrom,0 pagenoto from tblsectionmaster,tbldocmaster where tblsectionmaster.doc_id=tbldocmaster.docid and tblsectionmaster.doc_id={0}", docid);
                DataTable dtsection = GetDataTable(sqlquery);
                if (dtsection.Rows.Count > 0)
                {
                    Int32 pageFrom = 0, secid = 0;
                    string secName = "";
                    foreach (DataRow row in dtsection.Rows)
                    {
                        Search business = new Search();
                        business.PageFromNO = pageFrom;

                        secid = Convert.ToInt32(row["sec_id"]);
                        secName = Convert.ToString(row["sec_name"]);

                        Int32 PageToNO = getsectionPage(secName, ref pageFrom, fileid);

                        if (PageToNO == 0 || PageToNO == -1)
                        {
                            PageToNO = 0; business.PageFromNO = 0;
                        }
                        else
                        {
                            business.PageFromNO = business.PageFromNO == 0 ? 1 : business.PageFromNO;
                            PageToNO = PageToNO + (business.PageFromNO - 1);
                        }
                        //thisn ilist required to ger pervious from page number
                        ilist.Add(business);

                        if (chkexists(string.Format(@"select * from tblSectionDetail where Sec_Id={0} and Doc_Id={1} and DocType_Id={2}", secid, docid, DoctypeID)) == true)
                        {
                            Issuccess = ExecuteNonQuery(
                            string.Format(@"update tblSectionDetail set PageNoFrom={0},pagenoto={3} where Sec_Id={1} and DocType_Id={2}", business.PageFromNO, secid, DoctypeID, PageToNO));
                        }
                        else
                        {
                            sqlquery = string.Format("insert into tblsectiondetail (DocType_Id,Doc_Id,Sec_Id,PageNoFrom,pagenoto) values({0},{1},{2},{3},{4})", DoctypeID, docid, secid, business.PageFromNO, PageToNO);
                            Issuccess = ExecuteNonQuery(sqlquery);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                createlog("error while inserting section details-fileid" + fileid);
                return false;
            }
            if (Issuccess > 0) return true; else return false;
        }
        public Int32 ExecuteNonQuery(string Command)
        {
            Int32 value = 0;
            try
            {
                if (Sqlconn.State == ConnectionState.Closed)
                    Sqlconn.Open();
                SqlCommand cmd = new SqlCommand(Command, Sqlconn);
                value = Convert.ToInt32(cmd.ExecuteNonQuery());
            }
            catch (Exception ex)
            {
                return -1;
            }
            finally
            {
                if (Sqlconn.State == ConnectionState.Open)
                    Sqlconn.Close();
            }
            return value;

        }
        public Boolean chkexists(string qury)
        {
            Boolean boolval = false;
            try
            {

                if (Sqlconn.State == ConnectionState.Closed)
                {
                    Sqlconn.Open();
                }

                SqlDataAdapter sqlDa = new SqlDataAdapter(qury, Sqlconn);
                DataSet ds = new DataSet();
                sqlDa.Fill(ds);
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        boolval = true;
                    }
                    else
                    {
                        boolval = false;

                    }
                }
                else
                    boolval = false;


            }
            catch (Exception ex)
            {
                ex.ToString();
            }
            finally
            {

                Sqlconn.Close();


            }
            return boolval;

        }

        //get page count of pass section name
        private Int32 getsectionPage(string sectionName, ref Int32 pageFrom, Int64 fileID)
        {
            Int32 pagecount = 0;
            try
            {

                string folderName = RetureExcuteScaler(string.Format(@"select ltrim(rtrim(column4)) from {0} where file_id={1}", TableName, fileID));
                if (folderName == "")
                    return -2;

                DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(txtXLSFileName.Text) + @"/" + folderName);
                FileInfo[] fileInfo = dir.GetFiles("*.pdf");
                string filePath = "";
                for (Int32 i = 0; i < fileInfo.Length; i++)
                {
                    if (Path.GetFileNameWithoutExtension(fileInfo[i].FullName).ToLower().Trim() == sectionName.ToLower().Trim())
                    {
                        filePath = fileInfo[i].FullName;
                        if (pageFrom == 0) pageFrom = 1;
                        break;
                    }
                }
                //ArrayList strarrayList = (ArrayList)HttpContext.Current.Session["MultipleFiles"];

                //foreach (string file in strarrayList)
                //    if (Path.GetFileNameWithoutExtension(file).ToLower().Trim() == sectionName.ToLower().Trim())
                //    {
                //        filePath = file;
                //        break;
                //    }

                //if (filePath == "")
                //{
                //    createlog("Image name & section name not match." + folderName); return -1;
                //}
                if (filePath != "")
                {
                    if (Path.GetExtension(filePath).ToLower() != ".pdf")
                        using (FileStream fs = File.OpenRead(filePath))
                        {
                            Stream retStream = new MemoryStream();
                            using (System.Drawing.Image img = System.Drawing.Image.FromStream(fs, true, true))
                            {
                                pagecount = img.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);
                            }
                        }
                    else
                        pagecount = pdfPageCount(filePath);
                }
                pageFrom = pageFrom + pagecount;

            }
            catch (Exception ex)
            {
                return -1;
            }
            return pagecount;
        }
        /**/


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

        //insert into tblfile and return identity column
        public Int32 InsertDocumentReturnIdentity(string strFileName, string strFilePath, byte[] imgFileImage, Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo, int intFileStatus, string strInitialPath, string strIpAddress, Int32 PageCount)
        {
            Int32 temp = 0;
            try
            {

                if (Sqlconn.State == ConnectionState.Closed)
                {
                    Sqlconn.Open();
                }
                string strQuery = "insert into tblFile (FileName,FilePath,fileImage,DocId,ImportedBy,ImportedOn,FileSize,VersionNo,FileStatus,Initialpath,IpAddress,PageCount) values(@FileName,@FilePath,@fileImage,@DocId,@ImportedBy,@ImportedOn,@FileSize,@VersionNo,@FileStatus,@InitialPath,@IpAddress,@PageCount) select @@identity";
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
                //prmImportedOn.Value = datImportedOn;
                prmImportedOn.Value = datImportedOn;// System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");
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


                //int i = sqlCmd.ExecuteNonQuery();
                temp = Convert.ToInt32(sqlCmd.ExecuteScalar());
                //temp = "true";

            }
            catch (System.Exception ex)
            {
                temp = -1;
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
            DataTable dtDatabaseTable = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where docId =" + Convert.ToString(DocTypeID));
            if (dtDatabaseTable.Rows.Count > 0)
            {
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
            string strImportedBy = Program.strUserName;
            DateTime datImportedOn = System.DateTime.Now;
            int intVersionNo = 1; int intFileStatus = 1; int filenotfoundcount = 0; int OtherErrorCount = 0;
            string strIpAddress = ""; string FileSize = "";
            try
            {
                SetColumnName();//set columns value here

                //writting log file started from here.
                createlog("###############################  IMPORT LOG FOR Document type  " + cmbDocumentType.SelectedText + "    ############################################");
                createlog("...................................................................................................................................................");
                createlog(" Import started at : " + DateTime.Now);
                createlog("......................................................................................................................................");


                //loop started on excel datatable
                foreach (DataRow dtExcelrow in dtExceldata.Rows)
                {
                    System.Threading.Thread.Sleep(50); //do some intense task here.  
                    backgroundWorker1.ReportProgress(w, DateTime.Now); //notify progress to main thread. We also pass time information in UserState to cover this property in the example.  

                    colValue = "";
                    ErrorOccured = false;

                    for (int j = 0; j < listBox3.Items.Count; j++)
                    {
                        DataView dv = new DataView(dtDocDetails, "label='" + Convert.ToString(listBox3.Items[j]) + "'", "", DataViewRowState.CurrentRows);
                        if (dv.Count > 0)
                        {
                            if (Convert.ToString(dv[0]["FieldType"]) == "DateTime")
                            {

                                string dateString;
                                DateTime result;
                                CultureInfo provider = CultureInfo.InvariantCulture;
                                dateString = Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).Trim(); //Input Date
                                //format = "dd/MM/yyyy";//// Only change this format as per input date formate 
                                result = DateTime.ParseExact(dateString, strDateFormat, provider);
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
                                    colValue = "'" + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).Trim().Replace(",", "").Replace("'", "") + "'";
                                else
                                    colValue = colValue + ",'" + Convert.ToString(dtExcelrow[listBox2.Items[j].ToString()]).Trim().Replace(",", "").Replace("'", "") + "'";
                            }
                        }
                    }
                    //string strFileName = Convert.ToString(dtExcelrow["IMAGE 1"]);
                    string strFileName = Convert.ToString(dtExcelrow[ImageField]); ;//Selected Image field from dropdownlist.

                    string strFilePath = Path.GetDirectoryName(txtXLSFileName.Text);
                    string strFullFilePath = "";// strFilePath + "\\" + strFileName;
                    //FileInfo f = new FileInfo(strFullFilePath);
                    //if (f.Exists)
                    //{
                    //FileInfo fInfo = new FileInfo(strFullFilePath);
                    //if (fInfo.Length >= 1024)
                    //{
                    //    Int64 size = (f.Length) / 1024;
                    //    FileSize = size + " KB";
                    //}
                    //else
                    //{
                    //    FileSize = fInfo.Length.ToString() + " Bytes";
                    //}
                    //FileStream fs = new FileStream(strFullFilePath, FileMode.Open, FileAccess.Read);
                    //BinaryReader br = new BinaryReader(fs);
                    byte[] imgFileImage = null;//br.ReadBytes((int)fs.Length);
                    Int32 PageCount = 1;
                    if (Path.GetExtension(strFileName).ToLower() == ".tiff" || Path.GetExtension(strFileName).ToLower() == ".tif")
                    {
                        PageCount = FramCount(strFullFilePath); //dc.GetImagePageCount(FileUpload1);
                        if (PageCount == -1)
                            PageCount = 1;
                    }

                    if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "in")
                        InsertImageDetails(strFileName, strFilePath, imgFileImage, Convert.ToInt32(DocTypeID), strImportedBy, datImportedOn, FileSize, intVersionNo, intFileStatus, "", strIpAddress, TableName, PageCount, colname, colValue, "");
                    else
                    {
                        //imgFileImage = null;

                        //string Folder = ConfigurationManager.AppSettings["Repository"].ToString() + TableName;
                        //string strtime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + DateTime.Now.Millisecond.ToString("000");
                        //strFilePath = "~\\Repository\\Encrypt\\" + TableName + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc";
                        ////  ViewState["Folder"] = Folder;
                        //if (!Directory.Exists(Folder))
                        //{
                        //    Directory.CreateDirectory(Folder);
                        //}
                        InsertImageDetails("", strFilePath, null, Convert.ToInt32(DocTypeID), strImportedBy, datImportedOn, FileSize, intVersionNo, intFileStatus, "", strIpAddress, TableName, PageCount, colname, colValue, strFullFilePath);
                    }
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
                    //}
                    //else
                    //{
                    //    filenotfoundcount = filenotfoundcount + 1;
                    //    createlog("......................................................................................................................................");
                    //    createlog("       File not found on path = " + strFullFilePath + "  , hence file and data is not imported ");
                    //    strrec = "        Record Details : ";
                    //    foreach (DataColumn cc in dtExceldata.Columns)
                    //    {
                    //        strrec += cc.ColumnName + " = ";
                    //        strrec += dtExcelrow[cc.ColumnName].ToString() + " , ";
                    //    }
                    //    createlog(strrec);
                    //    createlog("......................................................................................................................................");
                    //}

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
                OtherErrorCount = OtherErrorCount + 1;
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
            // Application.Exit();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ImageField = comboBox1.Text.Trim();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            strDateFormat = comboBox2.Text.Trim();
        }


        #region ImageEncrypt
        public void ImageEncrypt(string fileIn, string fileOut, string Password)
        {

            // First we are going to open the file streams 
            try
            {


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
            catch (Exception)
            {
                createlog("error in encryption-" + fileIn);

            }
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


        [DllImport(@"urlmon.dll", CharSet = CharSet.Auto)]
        private extern static System.UInt32 FindMimeFromData(
            System.UInt32 pBC,
            [MarshalAs(UnmanagedType.LPStr)] System.String pwzUrl,
            [MarshalAs(UnmanagedType.LPArray)] byte[] pBuffer,
            System.UInt32 cbSize,
            [MarshalAs(UnmanagedType.LPStr)] System.String pwzMimeProposed,
            System.UInt32 dwMimeFlags,
            out System.UInt32 ppwzMimeOut,
            System.UInt32 dwReserverd
        );
        public string getMimeFromFile(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException(filename + " not found");

            byte[] buffer = new byte[256];
            using (FileStream fs = new FileStream(filename, FileMode.Open))
            {
                if (fs.Length >= 256)
                    fs.Read(buffer, 0, 256);
                else
                    fs.Read(buffer, 0, (int)fs.Length);
            }
            try
            {
                System.UInt32 mimetype;
                FindMimeFromData(0, null, buffer, 256, null, 0, out mimetype, 0);
                System.IntPtr mimeTypePtr = new IntPtr(mimetype);
                string mime = Marshal.PtrToStringUni(mimeTypePtr);
                Marshal.FreeCoTaskMem(mimeTypePtr);
                return mime;
            }
            catch (Exception e)
            {
                return "unknown/unknown";
            }
        }

        private string convertJPG2PDF(string PdfFileName, string Folderpath)
        {
            // string strPdfFile = Server.MapPath("~/Temps/") + "mahesh.pdf";
            //   string strFileName = Server.MapPath("~/Temps/") + "test.jpg";
            //System.Drawing.Image jpgImage = new System.Drawing.Bitmap(strFileName);


            iTextSharp.text.Document document = new iTextSharp.text.Document();


            using (var stream = new FileStream(Folderpath + Path.DirectorySeparatorChar + PdfFileName, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                iTextSharp.text.pdf.PdfWriter.GetInstance(document, stream);


                string[] strFileArray = Directory.GetFiles(Folderpath, "*.jpg");

                //foreach (string strFileName in strFileArray)
                //{
                for (int i = 0; i < strFileArray.Length; i++)
                {

                    if (i == 0)
                    {
                        System.Drawing.Image jpgImage = new System.Drawing.Bitmap(strFileArray[i]);
                        document = new iTextSharp.text.Document(
               new iTextSharp.text.Rectangle(0, 0, jpgImage.Width, jpgImage.Height));
                        document.Open();
                    }
                    using (var imageStream = new FileStream(strFileArray[i], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {

                        var image = iTextSharp.text.Image.GetInstance(imageStream);

                        document.Add(image);
                        document.NewPage();
                    }
                }
                document.Close();
                return Folderpath + Path.DirectorySeparatorChar + PdfFileName;
            }
        }

        private void mergerSectionwiseImage(string sectionName, string folderpath)
        {
            string destination = "", contentType = "";
            Int32 i = 0;
            try
            {
                DirectoryInfo dir = new DirectoryInfo(folderpath);
                FileInfo[] fileinfo = dir.GetFiles("*" + sectionName.ToLower().Trim().Substring(0, 5) + "*.*");
                if (fileinfo.Length > 0)
                {
                    if (sectionName.ToLower() == "invoices")
                    {
                        i = 0;
                        foreach (FileInfo filename1 in fileinfo)
                        {
                            // check whether or not the uploaded file is an image:
                            contentType = getFiletype.GetContentType(filename1.FullName);
                            if (contentType.Contains("pdf"))
                            {
                                if (sectionName.ToLower() == Path.GetFileNameWithoutExtension(filename1.FullName).ToLower().Trim())
                                { destination = filename1.FullName; }
                            }
                            else if (contentType.Contains("jpeg"))
                            {
                                convertJPG(folderpath + Path.DirectorySeparatorChar + sectionName + "_" + i.ToString() + ".pdf", filename1.FullName);
                                FileInfo f2 = new FileInfo(Path.GetDirectoryName(filename1.FullName) + Path.DirectorySeparatorChar + sectionName + "_" + i + ".pdf");
                                if (f2.Exists)
                                {
                                    f2.CopyTo(Path.GetDirectoryName(filename1.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf");
                                    f2.Delete();
                                    destination = Path.GetDirectoryName(f2.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf";
                                }
                                //destination = Path.DirectorySeparatorChar + sectionName + ".pdf";
                            }
                            i++;
                        }

                        foreach (FileInfo filename in fileinfo)
                        {
                            // check whether or not the uploaded file is an image:
                            contentType = getFiletype.GetContentType(filename.FullName);
                            if (contentType.Contains("pdf"))
                            {
                                // do something with the image ...
                                if (sectionName.ToLower() != Path.GetFileNameWithoutExtension(filename.FullName).ToLower().Trim())
                                {
                                    if (destination == "")
                                    {
                                        // destination = filename.FullName;
                                        filename.CopyTo(Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf");
                                        filename.Delete();
                                        destination = Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf";
                                    }
                                    else
                                        destination = ADDPDF2PDF(destination, filename.FullName, Path.GetFileNameWithoutExtension(destination));
                                }
                            }
                            else if (contentType.Contains("jpeg"))
                            {
                                convertJPG(folderpath + Path.DirectorySeparatorChar + sectionName + "_" + i + ".pdf", filename.FullName);

                                if (destination == "")
                                {
                                    FileInfo f1 = new FileInfo(Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + "_" + i + ".pdf");
                                    if (f1.Exists)
                                    {
                                        f1.CopyTo(Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf");
                                        f1.Delete();
                                        destination = Path.GetDirectoryName(f1.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf";
                                    }
                                }
                                else
                                {
                                    destination = ADDPDF2PDF(destination, folderpath + Path.DirectorySeparatorChar + sectionName + "_" + i + ".pdf", Path.GetFileNameWithoutExtension(destination));
                                    FileInfo f3 = new FileInfo(folderpath + Path.DirectorySeparatorChar + sectionName + "_" + i + ".pdf");
                                    f3.Delete();
                                }
                            }
                            i++;
                        }

                    }

                    if (sectionName.ToLower() == "payment_receipts")
                    {
                        FileInfo[] file = dir.GetFiles("*.*");
                        if (file.Length > 0)
                        {
                            for (Int32 z = 0; z < file.Length; z++)
                            {
                                // check whether or not the uploaded file is an image:
                                contentType = getFiletype.GetContentType(file[z].FullName);
                                if (contentType.Contains("jpeg") || contentType.Contains("pjpeg"))
                                {
                                    convertJPG(folderpath + Path.DirectorySeparatorChar + "payment_" + z + ".pdf", file[z].FullName);
                                }
                            }
                        }
                        fileinfo = dir.GetFiles("*" + sectionName.ToLower().Trim().Substring(0, 5) + "*.pdf");
                        foreach (FileInfo filename1 in fileinfo)
                            if (sectionName.ToLower() == Path.GetFileNameWithoutExtension(filename1.FullName).ToLower().Trim())
                            { destination = filename1.FullName; }

                        //if(newFile!="")
                        //    destination = ADDPDF2PDF(destination, newFile, Path.GetFileNameWithoutExtension(destination));

                        foreach (FileInfo filename in fileinfo)
                        {
                            contentType = getFiletype.GetContentType(filename.FullName);
                            if (contentType.Contains("pdf"))
                            {
                                if (sectionName.ToLower() != Path.GetFileNameWithoutExtension(filename.FullName).ToLower().Trim())
                                {
                                    if (destination == "")
                                    {
                                        // destination = filename.FullName;
                                        filename.CopyTo(Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf");
                                        filename.Delete();
                                        destination = Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf";
                                    }
                                    else { destination = ADDPDF2PDF(destination, filename.FullName, Path.GetFileNameWithoutExtension(destination)); filename.Delete(); }
                                }
                            }
                        }
                    }
                    if (sectionName.ToLower() == "welcome_kit")
                    {

                        foreach (FileInfo filename1 in fileinfo)
                        {
                            if (sectionName.ToLower() == Path.GetFileNameWithoutExtension(filename1.FullName).ToLower().Trim())
                            { destination = filename1.FullName; }
                        }

                        foreach (FileInfo filename in fileinfo)
                        {
                            // check whether or not the uploaded file is an image:
                            contentType = getFiletype.GetContentType(filename.FullName);
                            if (contentType.Contains("pdf"))
                            {
                                // do something with the image ...
                                if (sectionName.ToLower() != Path.GetFileNameWithoutExtension(filename.FullName).ToLower().Trim())
                                {
                                    if (destination == "")
                                    {
                                        // destination = filename.FullName;
                                        filename.CopyTo(Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf");
                                        filename.Delete();
                                        destination = Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf";
                                    }
                                    else
                                    {
                                        destination = ADDPDF2PDF(destination, filename.FullName, Path.GetFileNameWithoutExtension(destination));
                                        filename.Delete();
                                    }
                                }
                            }

                        }

                    }
                }
                else
                {
                    if (sectionName.ToLower() == "payment_receipts")
                    {
                        FileInfo[] file = dir.GetFiles("*.*");
                        if (file.Length > 0)
                        {
                            for (Int32 z = 0; z < file.Length; z++)
                            {
                                // check whether or not the uploaded file is an image:
                                contentType = getFiletype.GetContentType(file[z].FullName);
                                if (contentType.Contains("jpeg"))
                                {
                                    convertJPG(folderpath + Path.DirectorySeparatorChar + "payment_" + z + ".pdf", file[z].FullName);
                                }
                            }
                        }
                        fileinfo = dir.GetFiles("*" + sectionName.ToLower().Trim().Substring(0, 5) + "*.pdf");
                        foreach (FileInfo filename1 in fileinfo)
                            if (sectionName.ToLower() == Path.GetFileNameWithoutExtension(filename1.FullName).ToLower().Trim())
                            { destination = filename1.FullName; }

                        //if(newFile!="")
                        //    destination = ADDPDF2PDF(destination, newFile, Path.GetFileNameWithoutExtension(destination));

                        foreach (FileInfo filename in fileinfo)
                        {
                            contentType = getFiletype.GetContentType(filename.FullName);
                            if (contentType.Contains("pdf"))
                            {
                                if (sectionName.ToLower() != Path.GetFileNameWithoutExtension(filename.FullName).ToLower().Trim())
                                {
                                    if (destination == "")
                                    {
                                        // destination = filename.FullName;
                                        filename.CopyTo(Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf");
                                        filename.Delete();
                                        destination = Path.GetDirectoryName(filename.FullName) + Path.DirectorySeparatorChar + sectionName + ".pdf";
                                    }
                                    else
                                    {
                                        destination = ADDPDF2PDF(destination, filename.FullName, Path.GetFileNameWithoutExtension(destination));
                                        filename.Delete();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void convertJPG(string strPdfFile, string strFileName)
        {
            // string strPdfFile = Server.MapPath("~/Temps/") + "mahesh.pdf";
            //   string strFileName = Server.MapPath("~/Temps/") + "test.jpg";
            System.Drawing.Image jpgImage = new System.Drawing.Bitmap(strFileName);


            iTextSharp.text.Document document = new iTextSharp.text.Document(
                new iTextSharp.text.Rectangle(0, 0, jpgImage.Width, jpgImage.Height));

            using (var stream = new FileStream(strPdfFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                iTextSharp.text.pdf.PdfWriter.GetInstance(document, stream);
                document.Open();

                // string[] strFileArray = Directory.GetFiles(textBox1.Text, "*.jpg");

                //foreach (string strFileName in strFileArray)
                //{
                using (var imageStream = new FileStream(strFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {

                    var image = iTextSharp.text.Image.GetInstance(imageStream);

                    document.Add(image);
                    document.NewPage();
                }
                //}
                document.Close();
            }
        }
    }
}

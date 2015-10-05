using System;
using System.Configuration;
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
using System.Globalization;
using System.Security.Cryptography;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Runtime.InteropServices;
using PdfToImage;
namespace UploadApplication
{
    public partial class Form4 : Form
    {
        SqlConnection Sqlconn = new SqlConnection(ConfigurationSettings.AppSettings["SqlCon"].ToString());
        int w = 0, DocTypeID = 0; string TableName = "", colname = "";
        DataTable dtDocDetails = new DataTable();
        public Form4()
        {
            InitializeComponent();

            //mandatory. Otherwise will throw an exception when calling ReportProgress method  
            backgroundWorker1.WorkerReportsProgress = true;

            //mandatory. Otherwise we would get an InvalidOperationException when trying to cancel the operation  
            backgroundWorker1.WorkerSupportsCancellation = true;

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btnClear_Click(object sender, EventArgs e)
        {
            txtXLSFileName.Text = "";
            btoStart.Enabled = false;
            btoCancel.Enabled = false;

            w = 0;
            txtXLSFileName.Text = "";
            lblimportedcount.Text = "0";
            lbltotalcount.Text = "0";

            progressBar1.Value = 0;
            chkbranchcode.Checked = false;
            btnBrowse.Focus();
        }

        private void btoStart_Click(object sender, EventArgs e)
        {

            btoStart.Enabled = false;
            btoCancel.Enabled = true;
            backgroundWorker1.RunWorkerAsync();

        }

        private void btoCancel_Click(object sender, EventArgs e)
        {
            //notify background worker we want to cancel the operation.  
            //this code doesn't actually cancel or kill the thread that is executing the job.  
            btoStart.Enabled = true;
            btoCancel.Enabled = false;
            backgroundWorker1.CancelAsync();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {

            this.Cursor = Cursors.WaitCursor;
            FolderBrowserDialog oldBrowser = new FolderBrowserDialog();
            oldBrowser.ShowDialog();
            txtXLSFileName.Text = oldBrowser.SelectedPath;
            if (txtXLSFileName.Text != "")
                btoStart.Enabled = true;
            else btoStart.Enabled = false;

            DirectoryInfo dir = new DirectoryInfo(txtXLSFileName.Text.Trim());
            if (dir.GetFiles("*.pdf").Length > 0)
            {
                lbltotalcount.Text = Convert.ToString(dir.GetFiles("*.pdf").Length);

                progressBar1.Maximum = dir.GetFiles("*.pdf").Length + 1;
            }
            else
            {
                MessageBox.Show("Selected folder does not contains pdf files.");
                txtXLSFileName.Focus(); return;
            }

            this.Cursor = Cursors.Default;

        }

        private void cmbDocumentType_SelectedIndexChanged(object sender, EventArgs e)
        {
            //set doc type id
            DocTypeID = Convert.ToInt32(cmbDocumentType.SelectedValue);
            //set selected doctype table name
            TableName = RetureExcuteScaler("select DocName from tbldocmaster where docid= " + DocTypeID);


        }


        private void SetColumnName()
        {
            colname = "";
            dtDocDetails = GetDataTable("Select FieldName,label,fieldtype from tbldocdetail where docId =" + Convert.ToString(DocTypeID));
            if (dtDocDetails.Rows.Count > 0)
            {
                //foreach (object o in listBox3.Items)
                //{
                foreach (DataRow drDatabase in dtDocDetails.Rows)
                {
                    //if (drDatabase["label"].ToString() == o.ToString())
                    //{
                    if (colname == "")
                        colname = "[" + drDatabase["FieldName"].ToString() + "]";
                    else
                        colname = colname + "," + "[" + drDatabase["FieldName"].ToString() + "]";
                    //}
                }

                // }
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


            string strImportedBy = Program.strUserName;
            DateTime datImportedOn = System.DateTime.Now;
            int intVersionNo = 1; int intFileStatus = 1;
            string prodCode = "", acNumber = "", branchCode = "", strIpAddress = "", FileSize = "", strFilePath = "", strFullFilePath = "", colValue = "";

            try
            {


                //writting log file started from here.
                // createlog("###############################  IMPORT LOG FOR Document type  " + cmbDocumentType.SelectedText + "    ############################################");
                createlog("...................................................................................................................................................");
                createlog(" Import started at : " + DateTime.Now);
                createlog("......................................................................................................................................");

                SetColumnName();//set columns value here

                DirectoryInfo dir = new DirectoryInfo(txtXLSFileName.Text.Trim());

                //for notfound metadata
                string filepath = txtXLSFileName.Text + "//NotuploadFiles//";
                if (Directory.Exists(filepath) == false)
                    Directory.CreateDirectory(filepath);
                if (dir.GetFiles("*.pdf").Length > 0)
                {
                    FileInfo[] fileInfo = dir.GetFiles("*.pdf");
                    //loop started on excel datatable

                    int UploadedfileCounter = 0;
                    foreach (FileInfo file in fileInfo)
                    {
                        System.Threading.Thread.Sleep(50); //do some intense task here.  
                        backgroundWorker1.ReportProgress(w, DateTime.Now); //notify progress to main thread. We also pass time information in UserState to cover this property in the example.  
                        string branchName ="";
                        //first split filename
                       
                            branchName = dir.Name;
                         
                            branchCode = RetureExcuteScaler(string.Format(@"select distinct ltrim(rtrim([Branch Code])) from tblcustomerDatails where rtrim(ltrim(lower([Branch Name])))='{0}'", branchName.Trim().ToLower()));
                            string[] filenameArray = file.Name.Split('_');
                        //Commented for new file name given after june lot 
                            //if (filenameArray.Length == 2)
                            //{
                            //    if (filenameArray[0].Substring(0, 1) == "0")
                            //    {
                            //        prodCode = filenameArray[0];
                            //        acNumber = filenameArray[1].Split('.')[0];
                            //    }
                            //    else
                            //    {
                            //        acNumber = filenameArray[0];
                            //        prodCode = filenameArray[1].Split('.')[0];
                            //    }
                            //}
                            //else if (file.Name.Split('-').Length==2)
                            //{
                            //    filenameArray = file.Name.Split('-');
                            //    if (filenameArray.Length == 2)
                            //    {
                            //        if (filenameArray[0].Substring(0, 1) == "0")
                            //        {
                            //            prodCode = filenameArray[0];
                            //            acNumber = filenameArray[1].Split('.')[0];
                            //        }
                            //        else
                            //        {
                            //            acNumber = filenameArray[0];
                            //            prodCode = filenameArray[1].Split('.')[0];
                            //        }
                            //    }
                            //}
                            //else
                            //{
                            //    acNumber = filenameArray[0].Split('.')[0];
                            //    prodCode = "0101";
                            //}

                            if (filenameArray.Length == 3)
                            {
                                branchCode = filenameArray[0];
                                prodCode = filenameArray[1];
                                acNumber = filenameArray[2].Split('.')[0];
                            }
                               
                        if (branchCode == "")
                        {
                            MessageBox.Show("Branch name not exists."); btoStart.Enabled = true;
                            btoCancel.Enabled = false;
                            backgroundWorker1.CancelAsync();
                            return;
                        }
                        createlog("branchName:" + branchName + "  Branch Code:"+ branchCode +"  Ac No:"+ acNumber +"  Product Code:"+ prodCode);
                        colValue = "";
                        string strQuery=string.Format(@"select top 1 [Branch Code],[Branch Name],[Customer Number],[Product Code],[Account Number],[Customer Name],Convert(varchar(12),[Account opening date],101) [Account opening date] from tblcustomerDatails where [Branch Code]='{0}' and [Product Code]='{1}' and [Account Number]='{2}'", branchCode.Trim(), prodCode.Trim(), acNumber.Trim());
                        createlog("strQuery:" + strQuery);
                        DataTable dt = GetDataTable(strQuery);

                        if (dt == null)
                        {
                            createlog("dt == null");
                            file.MoveTo(filepath + file.Name);
                            continue;
                        }

                        if (dt.Rows.Count == 0)
                        {
                            createlog("dt.Rows.Count == 0");
                            file.MoveTo(filepath + file.Name);
                            continue;
                        }
                        string strChkUpload = string.Format(@"select  count(*) as count from table56 where Column1='{0}' and Column3='{1}' and column4='{2}'", branchName, prodCode.Trim(), acNumber.Trim());
                        createlog("strChkUpload:" + strChkUpload);
                        DataTable dtstrChkUpload = GetDataTable(strChkUpload);
                        if (Convert.ToInt32(dtstrChkUpload.Rows[0]["count"]) > 0)
                        {
                            UploadedfileCounter += 1;
                            file.Delete();
                            createlog("UploadedfileCounter:" + UploadedfileCounter);
                            continue;
                        }

                        for (Int32 i = 0; i < dtDocDetails.Rows.Count; i++)
                        {
                            //for (int j = 0; j < dtDocDetails.Rows.Count; j++)
                            //{
                            //DataView dv = new DataView(dtDocDetails, "label='" + dt.Columns[i].ColumnName + "'", "", DataViewRowState.CurrentRows);
                            //if (dv.Count > 0)
                            //{
                            if (Convert.ToString(dtDocDetails.Rows[i]["FieldType"]) == "DateTime")
                            {

                                string dateString;
                                //DateTime result;
                                //CultureInfo provider = CultureInfo.InvariantCulture;
                                dateString = Convert.ToString(dt.Rows[0][dtDocDetails.Rows[i]["label"].ToString()]).Trim(); //Input Date
                                //////format = "dd/MM/yyyy";//// Only change this format as per input date formate 
                                //result = DateTime.ParseExact(dateString, strDateFormat, provider);
                                //string dateDDMMYYY = result.ToString("MM/dd/yyyy");
                                //dateDDMMYYY = dateDDMMYYY.Replace("-", "/");

                                if (colValue == "")
                                    colValue = "'" + dateString + "'";
                                else
                                    colValue = colValue + ",'" + dateString + "'";

                            }
                            else if (Convert.ToString(dtDocDetails.Rows[i]["FieldType"]) == "Text")
                            {
                                if (colValue == "")
                                    colValue = "'" + Convert.ToString(dt.Rows[0][dtDocDetails.Rows[i]["label"].ToString()]).Trim().Replace(",", "").Replace("'", "") + "'";
                                else
                                    colValue = colValue + ",'" + Convert.ToString(dt.Rows[0][dtDocDetails.Rows[i]["label"].ToString()]).Trim().Replace(",", "").Replace("'", "") + "'";
                            }
                            else
                            {
                                if (colValue == "")
                                    colValue = Convert.ToString(dt.Rows[0][dtDocDetails.Rows[i]["label"].ToString()]).Trim().Replace(",", "").Replace("'", "");
                                else
                                    colValue = colValue + "," + Convert.ToString(dt.Rows[0][dtDocDetails.Rows[i]["label"].ToString()]).Trim().Replace(",", "").Replace("'", "");

                            }
                            // }
                            //}
                        }
                        if (file.Length >= 1024)
                        {
                            Int64 size = (file.Length) / 1024;
                            FileSize = size + " KB";
                        }
                        else
                        {
                            FileSize = file.Length.ToString() + " Bytes";
                        }

                        Int32 PageCount = 1;


                        PageCount = FramCount_new(file.FullName); //dc.GetImagePageCount(FileUpload1);
                        if (PageCount == -1)
                            PageCount = 1;


                        if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "in")
                        {
                            FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);
                            BinaryReader br = new BinaryReader(fs);
                            byte[] imgFileImage = br.ReadBytes((int)fs.Length);
                            InsertImageDetails(file.Name, file.FullName, imgFileImage, DocTypeID, strImportedBy, datImportedOn, FileSize, intVersionNo, intFileStatus, "", strIpAddress, TableName, PageCount, "", file, colname, colValue);
                        }
                        else
                        {
                            InsertImageDetails(file.Name, strFilePath, null, DocTypeID, strImportedBy, datImportedOn, FileSize, intVersionNo, intFileStatus, "", strIpAddress, TableName, PageCount, strFullFilePath, file, colname, colValue);
                        }
                        //if cnacel button clicked
                        if (backgroundWorker1.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }
                        //increament record imported or no imported.//
                        w += 1;
                    }
                }
                else { MessageBox.Show("Selected folder does not contains pdf files."); txtXLSFileName.Focus(); return; }
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
        /*insert into database*/

        private void InsertImageDetails(string strFileName, string strFilePath, byte[] imgFileImage,
                Int32 intDocId, string strImportedBy, DateTime datImportedOn, string strFileSize, Int32 intVersionNo,
                Int32 intFileStatus, string strInitialPath, string strIpAddress, string TableName,
    Int32 PageCount, string strFullFilePath, FileInfo file, string columnName, string ColumnValue)
        {

            string strDate = System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
            if (strDate.IndexOf("-") != -1)
                datImportedOn = DateTime.ParseExact(strDate, "dd-MM-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            else
                datImportedOn = DateTime.ParseExact(strDate, "dd/MM/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);


            //insert data into tblfile

            //string Folder = ConfigurationManager.AppSettings["Repository"].ToString() + TableName;

            string Folder = ConfigurationManager.AppSettings["Repository"].ToString();
            // string strtime = DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + DateTime.Now.Millisecond.ToString("000");
            if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
            {
                // strFilePath = "~\\Repository\\Encrypt\\" + TableName + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc";
                strFilePath = "~\\Repository\\Encrypt\\" + TableName + "\\" + strFileName;
                if (!Directory.Exists(Folder))
                {
                    Directory.CreateDirectory(Folder);
                }
               
                imgFileImage = new byte[1];
            }

           // string DocInsert = InsertDocument(strFileName, strFilePath, imgFileImage, intDocId, Program.UserId, datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);
            string DocInsert = InsertDocument(strFileName, strFilePath, imgFileImage, intDocId, "admin", datImportedOn, strFileSize, intVersionNo, intFileStatus, strInitialPath, strIpAddress, PageCount);
            if (DocInsert == "true")
            {
               // createlog("DocInsert == true");
                //get id of last record inserted
                // DataTable maxidds = 
                string maxid = RetureExcuteScaler("select MAX(Id) from tblFile");
             //   createlog("maxid " + maxid);
                //insert into index table
                string qury = "insert into " + TableName + " (File_Id,ImportedBy,ImportedOn," + columnName + " ) values(" + maxid + ",'0','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt") + "'," + ColumnValue + " ) select @@identity";
                //createlog("insert into index table Query " + qury);
                Int32 doctypeId = Convert.ToInt32(RetureExcuteScaler(qury));
                if (doctypeId > 0)
                {
                    //createlog("doctypeId " + doctypeId);
                    //insert into tblfileactions table
                    //string updateQuery="Insert into TblFileActions(FileId,DocId,FileName,TransProcess,UserName,TransDate) values(" + maxid + "," + intDocId + ",'" + strFileName + "','Document Uploaded','" + strImportedBy + "','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");
                   // createlog("updateQuery " + updateQuery);
                    //string updateQuery = "Insert into TblFileActions(FileId,DocId,FileName,TransProcess,UserName,TransDate) values(" + maxid + "," + intDocId + ",'" + strFileName + "','Document Uploaded','admin','" + System.DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");
                    //createlog("updateQuery " + updateQuery);
                    //updatedata(updateQuery);
                    if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
                    {
                        //copy file to repository
                        if (File.Exists(Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf") == false)
                        {
                          //  createlog("File not Exists " + Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf");
                            //file.CopyTo(Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf");
                            file.MoveTo(Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf");
                            
                        }
                        else
                        {
                            //createlog("File  Exists " + Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf");
                            File.Delete(Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf");
                            file.MoveTo(Folder + "//" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf");
                        }
                        if (Path.GetExtension(strFileName).Trim().ToLower() == ".pdf")
                        {
                            if (!Directory.Exists(Folder + "\\PageTiffFiles"))
                            {
                                Directory.CreateDirectory(Folder + "\\PageTiffFiles");
                            }

                            //string OutPutTiffFileName = Path.GetFileNameWithoutExtension(strFileName) + "_" + Convert.ToString(HttpContext.Current.Session["FolderId"]) + ".pdf";
                            string OutPutTiffFileName = Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf";
                            string strSaveLocation = Folder + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + maxid + ".pdf";
                          //  createlog("OutPutTiffFileName :" + OutPutTiffFileName + "   strSaveLocation :" + strSaveLocation);

                            if (convertPDF_TIFF(strSaveLocation, OutPutTiffFileName) == true)
                            {
                                //createlog("convertPDF_TIFF = true");
                                //insert section Details
                                UpdateSectionDetails(intDocId, Convert.ToInt32(maxid), TableName, doctypeId, PageCount);
                            }

                        }
                    }
                    ////save image
                    //if (ConfigurationManager.AppSettings["SaveImage"].ToString().ToLower() == "out")
                    //    ImageEncrypt(strFullFilePath, Folder + "\\" + Path.GetFileNameWithoutExtension(strFileName) + "_" + strtime + ".enc", "encrypt123");
                }
                else
                {
                    //ErrorOccured = true;
                    //if error occured then delete inserted record from tblfile
                    updatedata("delete from tblfile where id=" + maxid + ")");
                    createlog("Error occured while insert in " + TableName + " document table : " + qury);
                }

            }
        }
        //update section
        private bool UpdateSectionDetails(Int32 docid, Int64 fileid,
           string TabelName, Int32 DoctypeID, Int32 pageCount)
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
                    Int32 pageFrom = 0, PageToNO = 0, secid = 0;
                    string secName = "";
                    foreach (DataRow row in dtsection.Rows)
                    {
                        Search business = new Search();
                        business.PageFromNO = pageFrom;

                        secid = Convert.ToInt32(row["sec_id"]);
                        secName = Convert.ToString(row["sec_name"]);

                        switch (secName.ToLower().Trim())
                        {
                            case "account opening form":
                                pageFrom = 1; PageToNO = 4;
                                break;
                            case "photo identity proof":
                                pageFrom = 5; PageToNO = 5;
                                break;
                            case "address proof":
                                pageFrom = 6; PageToNO = 6;
                                break;
                            case "other documents":
                                pageFrom = 7; PageToNO = pageCount;
                                break;
                        }
                        if (chkexists(string.Format(@"select * from tblSectionDetail where Sec_Id={0} and Doc_Id={1} and DocType_Id={2}", secid, docid, DoctypeID)) == true)
                        {
                            Issuccess = ExecuteNonQuery(
                            string.Format(@"update tblSectionDetail set PageNoFrom={0},pagenoto={3} where Sec_Id={1} and DocType_Id={2}", pageFrom, secid, DoctypeID, PageToNO));
                        }
                        else
                        {
                            sqlquery = string.Format("insert into tblsectiondetail (DocType_Id,Doc_Id,Sec_Id,PageNoFrom,pagenoto) values({0},{1},{2},{3},{4})", DoctypeID, docid, secid, pageFrom, PageToNO);
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
        //insert into tblfile and return identity column
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
                // ErrorOccured = true;
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

                createlog("Image frame count not found.");
                return -1;
            }
            return frmCount;
        }
        /*end*/
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

        private void Form4_Load(object sender, EventArgs e)
        {
            DateTime datImportedOn = new DateTime();
            string strDate = System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");
            if (strDate.IndexOf("-") != -1)
                datImportedOn = DateTime.ParseExact(strDate, "dd-MM-yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
            else
                datImportedOn = DateTime.ParseExact(strDate, "dd/MM/yyyy hh:mm:ss tt", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);

            DataTable dt = GetDataTable("select docLabel,DocId from tblDocmaster");
            if (dt.Rows.Count > 0)
            {
                cmbDocumentType.DisplayMember = "docLabel";
                cmbDocumentType.ValueMember = "DocId";
                cmbDocumentType.DataSource = dt;

                //dtDocDetails = GetDataTable(string.Format("select fieldid,label,FieldName,fieldtype from tbldocdetail  where docid={0}", Convert.ToString(dt.Rows[0][1])));
                // LoadListBox(dtDocDetails, lstDatabaseColumn, "row");
                DocTypeID = Convert.ToInt32(dt.Rows[0]["DocId"]);
                //set selected doctype table name
                TableName = RetureExcuteScaler("select DocName from tbldocmaster where docid= " + DocTypeID);

                btnBrowse.Focus();
            }
            else
            {
                MessageBox.Show("Create document type.");
                this.Close();
            }
        }



        /**/
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
                //  ErrorOccured = true;
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
                createlog(ex.Message);
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

        public bool convertPDF_TIFF(string filename, string OutPutFileName)
        {
            bool Converted = false;
            try
            {
                //This is the object that perform the real conversion!
                PDFConvert converter = new PDFConvert();
                createlog(filename+' '+OutPutFileName);


                //Setup the converter
                // if (numericThreads.Value > 0)
                converter.RenderingThreads = -1;// (int)numericThreads.Value;

                //if (((int)numericTextSampling.Value > 0) && ((int)numericTextSampling.Value != 3))
                converter.TextAlphaBit = -1;// (int)numericTextSampling.Value;

                // if (((int)numericGraphSampling.Value > 0) && ((int)numericGraphSampling.Value != 3))
                converter.TextAlphaBit = -1;//(int)numericGraphSampling.Value;


                converter.OutputToMultipleFile = true;//checkSingleFile.Checked;

                converter.FirstPageToConvert = -1;// (int)numericFirstPage.Value;
                converter.LastPageToConvert = -1;// (int)numericLastPage.Value;
                converter.FitPage = true;//checkFitTopage.Checked;
                converter.JPEGQuality = 25;// (int)numQuality.Value;
                converter.ResolutionX = 100;
                converter.ResolutionY = 100;
                // converter.OutputFormat = "tifflzw";//comboFormat.Text;
                converter.OutputFormat = "jpeg";//comboFormat.Text;
                System.IO.FileInfo input = new FileInfo(filename);

                //string OutputFilePath = ConfigurationManager.AppSettings["Repository"].ToString() + TableName + "\\PageTiffFiles\\";//HttpContext.Current.Server.MapPath("~\\Repository\\PageTiffFiles");
                string OutputFilePath = ConfigurationManager.AppSettings["Repository"].ToString() +  "PageTiffFiles\\";//HttpContext.Current.Server.MapPath("~\\Repository\\PageTiffFiles");
                if (!Directory.Exists(OutputFilePath))
                {
                    Directory.CreateDirectory(OutputFilePath);
                }
                string output = string.Format("{0}\\{1}{2}", OutputFilePath, OutPutFileName, ".tif");
                //If the output file exist alrady be sure to add a random name at the end until is unique!
                //while (System.IO.File.Exists(output))
                //{
                //    output = output.Replace(".tif", string.Format("{1}{0}", ".tif", DateTime.Now.Ticks));
                //}

                Converted = converter.Convert(input.FullName, output);
            }
            catch (Exception ex)
            {
                createlog(ex.Message);
            }


            return Converted;
        }
    }

}

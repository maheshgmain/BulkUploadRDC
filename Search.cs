using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for Search
/// </summary>
public class Search
{
    public Search()
    {
        //
        // TODO: Add constructor logic here
        //
    }

    private string strDocTypeName;
    private int intDocTypeId;
    private int intDocId;

    private string strFieldName;
    private string strValueList;
    private string strLabel;
    private string strFieldType;
    private string strTableName;
    private int intFileId;
    private string strFileName;
    public string[] ColumnLabel { get; set; }
    public string[] ColumnValue { get; set; }
    private string _Extension;


    public int DocId
    {
        get { return intDocId; }
        set { intDocId = value; }
    }
    public int FileId
    {
        get { return intFileId; }
        set { intFileId = value; }
    }
    public string TableName
    {
        get { return strTableName; }
        set { strTableName = value; }
    }
    public string FileName
    {
        get { return strFileName; }
        set { strFileName = value; }
    }
    //public string TableName
    //{
    //    get { return strTableName; }
    //    set { strTableName = value; }
    //}
    public string DocTypeName
    {
        get { return strDocTypeName; }
        set { strDocTypeName = value; }
    }
    public int DocTypeId
    {
        get { return intDocTypeId; }
        set { intDocTypeId = value; }
    }
    public string FieldName
    {
        get { return strFieldName; }
        set { strFieldName = value; }
    }
    public string ValueList
    {
        get { return strValueList; }
        set { strValueList = value; }
    }
    public string Label
    {
        get { return strLabel; }
        set { strLabel = value; }
    }
    public string FieldType
    {
        get { return strFieldType; }
        set { strFieldType = value; }
    }

    public string Extension
    {
        get { return _Extension; }
        set { _Extension = value; }
    }

    private Int64 _GroupId;
    public Int64 GroupId { get { return _GroupId; } set { _GroupId = value; } }

    private string _GroupName;

    public string GroupName
    {
        get { return _GroupName; }
        set { _GroupName = value; }
    }

    private string _GroupDescription;

    public string GroupDescription
    {
        get { return _GroupDescription; }
        set { _GroupDescription = value; }
    }
    private string _ColGroupName;

    public string ColGroupName
    {
        get { return _ColGroupName; }
        set { _ColGroupName = value; }
    }

    private string _ColGrouDescription;

    public string ColGrouDescription
    {
        get { return _ColGrouDescription; }
        set { _ColGrouDescription = value; }
    }
    private Int32 _Colnum;

    public Int32 Colnum
    {
        get { return _Colnum; }
        set { _Colnum = value; }
    }
    private Int32 _Rownum;

    public Int32 Rownum
    {
        get { return _Rownum; }
        set { _Rownum = value; }
    }

    private Int32 _UserCount;

    public Int32 UserCount
    {
        get { return _UserCount; }
        set { _UserCount = value; }
    }
    private string _HighestQualification;

    public string HighestQualification
    {
        get { return _HighestQualification; }
        set { _HighestQualification = value; }
    }
    private string _Email;

    public string Email
    {
        get { return _Email; }
        set { _Email = value; }
    }
    private string _Address;

    public string Address
    {
        get { return _Address; }
        set { _Address = value; }
    }

    private string _UserName;

    public string UserName
    {
        get { return _UserName; }
        set { _UserName = value; }
    }

    private string _Password;

    public string Password
    {
        get { return _Password; }
        set { _Password = value; }
    }
    private string _Status;
    public string Status
    {

        get
        {
            return _Status;
        }
        set
        {
            _Status = value;
        }

    }
    private long _UserId;
    public long UserId
    {
        get { return _UserId; }
        set { _UserId = value; }
    }

    private string _ContactNo;
    public string ContactNo
    {

        get
        {
            return _ContactNo;
        }
        set
        {
            _ContactNo = value;
        }

    }
    private string _colFirstName;

    public string ColFirstName
    {
        get { return _colFirstName; }
        set { _colFirstName = value; }
    }
    private string _colLastName;

    public string ColLastName
    {
        get { return _colLastName; }
        set { _colLastName = value; }
    }
    private string _colHighestQualification;

    public string ColHighestQualification
    {
        get { return _colHighestQualification; }
        set { _colHighestQualification = value; }
    }
    private string _colAddress;

    public string ColAddress
    {
        get { return _colAddress; }
        set { _colAddress = value; }
    }
    private string _colContactNo;

    public string ColContactNo
    {
        get { return _colContactNo; }
        set { _colContactNo = value; }
    }
    private string _colEmail;

    public string ColEmail
    {
        get { return _colEmail; }
        set { _colEmail = value; }
    }
    private string _colUserName;

    public string ColUserName
    {
        get { return _colUserName; }
        set { _colUserName = value; }
    }
    private string _Name;

    public string Name
    {
        get { return _Name; }
        set { _Name = value; }
    }

    private string _FatherName;

    public string FatherName
    {
        get { return _FatherName; }
        set { _FatherName = value; }
    }
    private string _SurName;

    public string SurName
    {
        get { return _SurName; }
        set { _SurName = value; }
    }
    private string _Date;
    public string Date
    {
        get { return _Date; }
        set { _Date = value; }
    }
    private Int32 _NO;
    public Int32 NO
    {
        get { return _NO; }
        set { _NO = value; }
    }
    private Int32 _PageFromNO;
    public Int32 PageFromNO
    {
        get { return _PageFromNO; }
        set { _PageFromNO = value; }
    }
    private Int32 _PageToNO;
    public Int32 PageToNO
    {
        get { return _PageToNO; }
        set { _PageToNO = value; }
    }
    private Int32 _SecID;
    public Int32 SecID
    {
        get { return _SecID; }
        set { _SecID = value; }
    }

    private Int32 _SecIndexId;
    public Int32 SecIndexId
    {
        get { return _SecIndexId; }
        set { _SecIndexId = value; }
    }
    private string _SecIndexValue;
    public string SecIndexValue
    {
        get { return _SecIndexValue; }
        set { _SecIndexValue = value; }
    }

    private string _IsworkFlow;
    public string IsworkFlow
    {
        get { return _IsworkFlow; }
        set { _IsworkFlow = value; }
    }
    private string _WorkflowType;
    public string WorkflowType
    {
        get { return _WorkflowType; }
        set { _WorkflowType = value; }
    }
    /*password policy*/
    private Int32 _passwordLength;
    public Int32 passwordLength
    {
        get { return _passwordLength; }
        set { _passwordLength = value; }
    }
    private bool _lowerupperChar;
    public bool lowerupperChar
    {
        get { return _lowerupperChar; }
        set { _lowerupperChar = value; }
    }
    private bool _numberChar;
    public bool numberChar
    {
        get { return _numberChar; }
        set { _numberChar = value; }
    }
    private bool _onespecialChar;
    public bool onespecialChar
    {
        get { return _onespecialChar; }
        set { _onespecialChar = value; }
    }
    private bool _twospecialChar;
    public bool twospecialChar
    {
        get { return _twospecialChar; }
        set { _twospecialChar = value; }
    }
    private Int32 _passwordExpire;
    public Int32 passwordExpire
    {
        get { return _passwordExpire; }
        set { _passwordExpire = value; }
    }
    /*end password policy*/


}

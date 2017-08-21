using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using DataAccessLayer;
using System.Data;

/// <summary>
/// Summary description for DayActivity
/// </summary>
public class DayActivity
{
    string manual = System.Configuration.ConfigurationManager.AppSettings["manualFolder"];
    string connectionstring = "Data Source=mignt014;Initial Catalog=DoxPro_Env_Flex;Integrated Security=True";
	public DayActivity( DateTime date )
	{
        Date = date;
        PsSucced = 100;
        PsMovedToManual = CountFilesByDate( manual );
        PsError = 1;
        InvoceSSucced = 47;
        InvoicesFailed = 3;        

        psTotal = PsSucced + PsMovedToManual + PsError ;
        invoiceTotal = InvoceSSucced + InvoicesFailed;

        using (IDataSupplier db = DataManager.GetDataSupplier(DataManager.defaultType, connectionstring))
        {
            string sql =
                @"select COUNT(id) 
                from user_action_tab 
                where param_3=@docType
                and action_date between @date and DateAdd(day,1,@date)" ;

            db.OpenQuery();
            db.AddParameter("date", Date);
            db.AddParameter("docType", "18");
            DataSet ds = db.GetData(sql);
            PsSucced = (int)ds.Tables[0].Rows[0][0];
            db.OpenQuery();
            db.AddParameter("date", Date);
            db.AddParameter("docType", "16");
            ds = db.GetData(sql ) ;
            InvoceSSucced = (int)ds.Tables[0].Rows[0][0];
        }

	}

    int psTotal, invoiceTotal  ;

    public double Percent(int n, int t)
    {
        return 100 * n / t;
    }

    public DateTime Date { get; set; }
    public int PsSucced { get; set; }
    public int PsMovedToManual { get; set; }
    public int PsError { get; set; }
    public int InvoceSSucced { get; set; }
    public int InvoicesFailed { get; set; }

    public double PsSuccedPercent { get { return 100 * PsSucced / psTotal ; }}
    public double PsMovedToManualPercent { get { return 100 * PsMovedToManual / psTotal ; }}
    public double PsErrorPercent { get { return 100 * PsError / psTotal; } }

    public double InvoceSSuccedPercent { get { return Percent(InvoceSSucced, invoiceTotal); } }
    public double InvoicesFailedPercent { get { return Percent(InvoicesFailed, invoiceTotal); } }

    public int CountFilesByDate(string path)
    {
        DirectoryInfo di = new DirectoryInfo(path);
        FileSystemInfo[] files = di.GetFileSystemInfos();
        return files.Where( f => f.CreationTime >= Date && f.CreationTime < Date.AddDays(1) && f.Extension == ".pdf" ).Count() ;
    }

    static int CountFilesInPath(string name)
    {
        return Directory.GetFiles( System.Configuration.ConfigurationManager.AppSettings[name], "*.*").Count();
    }

    public static int PsWaitingInQueue
    {
        get { return CountFilesInPath( "psNewPath" ) ; }
    }

    public static int PsTotalManual
    {
        get { return CountFilesInPath("psManualPath"); }
    }

    public static int PsTotalError
    {
        get { return CountFilesInPath("psErrorPath"); }
    }

    public static int InvoicesTotalWaitingInQueue
    {
        get { return CountFilesInPath( "invoiceNewPath" ) ; }
    }

    public static int InvoicesTotalFailed
    {
        get { return CountFilesInPath("invoiceErrorPath"); }
    }

}
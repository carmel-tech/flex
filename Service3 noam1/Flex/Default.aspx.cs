using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    string queuePath = System.Configuration.ConfigurationManager.AppSettings["queuePath"];

    string CurrentFile
    {
        get
        {
            return ViewState["fileName"].ToString();
        }
        set
        {
            ViewState["fileName"] = value;
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            RefreshList();
            if (filesInQueue.Items.Count != 0)
            {
                filesInQueue.SelectedIndex = 0;
                filesInQueue_SelectedIndexChanged(null, null);
            }
//            Page.ClientScript.RegisterClientScriptInclude("Registration", ResolveUrl("~/default.js"));
        }
    }

    void RefreshList()
    {
        bool migdal = Factory.SelectedIndex == 0;
        //string [] files = Directory.GetFiles(queuePath, "*.pdf");
        string[] files = Directory.GetFiles(queuePath, "*.pdf").OrderByDescending(filename => new FileInfo(filename).CreationTime).ToArray();
        //var files = Directory.EnumerateFiles(queuePath, "*.pdf").OrderByDescending(filename  => new FileInfo(filename).CreationTime);
        List<string> subset = new List<string>();
        for (int i = 0; i < files.Length; i++)
        { 
            int n =0;
            string name = files[i].Substring(queuePath.Length + 1);
            if ((name.Length >= 4 && int.TryParse(name.Substring(0, 4), out n))
                ||
                (name.Length >= 5 && name.StartsWith("BC"))
              )
            {
                if (migdal)
                {
                    subset.Add(name);
                }
            }
            else
            {
                if (!migdal)
                {
                    subset.Add(name);
                }
            }
        }
        filesInQueue.DataSource = subset;
        filesInQueue.DataBind();
        Total.Text = subset.Count.ToString();
    }
    protected void Remove_Click(object sender, EventArgs e)
    {
        CurrentFile = Path.Combine(queuePath, filesInQueue.SelectedValue);
        File.Delete(CurrentFile);
        int last = filesInQueue.SelectedIndex;
        RefreshList();

        if (filesInQueue.Items.Count != 0 )
        {
            filesInQueue.SelectedIndex = last;
            filesInQueue_SelectedIndexChanged(null, null);
        }
        customer.Text = String.Empty;
        psNumber.Text = String.Empty;
    }

    protected void submitBarCode_Click(object sender, EventArgs e)
    {
        if (customer.Text.Length != 4 || psNumber.Text.Length != 6)
        {
            if (filesInQueue.Items.Count != 0)
            {
                filesInQueue_SelectedIndexChanged(null, null);
            }
            return;
        }
        string newPackingSleepsPath = System.Configuration.ConfigurationManager.AppSettings["newPath"];
        string barcode = string.Format("BC{0}{1}.pdf", customer.Text, psNumber.Text) ;
        string newName = Path.Combine(newPackingSleepsPath, barcode);
        if (File.Exists(newName))
        {
            File.Delete(newName);
        }
        try
        {
            File.Move(CurrentFile, newName);
        }
        catch (Exception)
        {
        }

        int last = filesInQueue.SelectedIndex;
        RefreshList();

        if (filesInQueue.Items.Count != 0)
        {
            if (last >= filesInQueue.Items.Count)
            {
                last = filesInQueue.Items.Count - 1;
            }
            filesInQueue.SelectedIndex = last;
            filesInQueue_SelectedIndexChanged(null, null);
        }
        customer.Text = String.Empty;
        psNumber.Text = String.Empty;
    }

    protected void manualQueueButton_Click(object sender, EventArgs e)
    {
        string manualFolder =  System.Configuration.ConfigurationManager.AppSettings["manualFolder"];
        string newFile = Path.Combine(manualFolder, Path.GetFileName(CurrentFile));
        if (File.Exists(newFile))
        {
            File.Delete(newFile);
        }
        try
        {
            File.Move(CurrentFile, newFile);
        }
        catch (Exception)
        {
        }
    }
    protected void Factory_OnSelectedIndexChanged(object sender, EventArgs e)
    {
        RefreshList();
        if (filesInQueue.Items.Count != 0)
        {
            filesInQueue.SelectedIndex = 0;
            filesInQueue_SelectedIndexChanged(null, null);
        }
        customer.Text = String.Empty;
        psNumber.Text = String.Empty;
    }
    protected void filesInQueue_SelectedIndexChanged(object sender, EventArgs e)
    {
        try
        {
            CurrentFile = Path.Combine(queuePath, filesInQueue.SelectedValue);
            System.Web.UI.HtmlControls.HtmlGenericControl pdf = new System.Web.UI.HtmlControls.HtmlGenericControl();
            if (! Request.Browser.Type.ToUpper().Contains("IE"))
            {
                pdf.TagName = "embed";
                pdf.ID = "pdf";
                pdf.Attributes.Add("src", String.Format("pdf.aspx?filename={0}", CurrentFile));
            }
            else
            {
                pdf.TagName = "object";
                pdf.ID = "pdf";
                pdf.Attributes.Add("classid", "clsid:{CA8A9780-280D-11CF-A24D-444553540000}");
                pdf.Attributes.Add("data", String.Format("pdf.aspx?filename={0}", CurrentFile));
                pdf.Attributes.Add("name", "pdf");
                //pdf.Attributes.Add("onload", "javascript: pdf.SetZoom(8.0);");
            }
            pdf.Attributes.Add("type", "application/pdf");
            pdf.Style.Add("width", "900px");
            pdf.Style.Add("height", "700px");
            pdfframe.Controls.Add(pdf);
            System.Web.UI.HtmlControls.HtmlGenericControl script = new System.Web.UI.HtmlControls.HtmlGenericControl();
            script.TagName = "script";
            script.Attributes.Add("src", "default.js");
            script.Attributes.Add("type", "text/javascript");
            pdfframe.Controls.Add(script);
        }
        catch
        { 
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class Pdf : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            string fileName = Request.QueryString["filename"];
            if (fileName != String.Empty)
            {
                byte[] content = File.ReadAllBytes(fileName);
                Response.ContentType = "application/pdf";
                Response.BinaryWrite(content);
            }
        }
        catch { }
    }
}
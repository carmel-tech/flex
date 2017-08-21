using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DashBoard : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        List<DayActivity> days = new List<DayActivity>();
        DateTime d = new DateTime(2012, 2, 1);
        for (int i = 0; i < 10; i++, d=d.AddDays(1))
        {
            days.Add(new DayActivity(d));
        }
        data.DataSource = days;
        data.DataBind();
    }
}
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport;


namespace ConsumptionAnalysisReport.UI_ConsumptionAnalysisReport
{
    public partial class MShuiNiFenMoConsumption : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        [WebMethod]
        public static string GetMShuiNiFenMoConsumption(string mySelectTime)
        {
            DataTable table = MShuiNiFenMoConsumptionService.GetMShuiNiFenMoConsumption(mySelectTime);
            string json = EasyUIJsonParser.DataGridJsonParser.DataTableToJson(table);
            return json;
        }
    }
}
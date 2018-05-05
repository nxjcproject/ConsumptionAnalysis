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
    public partial class HuiZhuanYaoConsumption : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        [WebMethod]
        public static string GetHZYConsumptionInfo(string m_SelectFirstTime, string m_SelectSecondTime, string m_selectThirdTime)
        {
            DataTable table = HuiZhuanYaoConsumptionService.GetHZYConsumptionTable(m_SelectFirstTime, m_SelectSecondTime, m_selectThirdTime);
            string json = EasyUIJsonParser.DataGridJsonParser.DataTableToJson(table);
            return json;
        }
    }
}
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
    public partial class ShuiNiFenMoFengGuPingConsumption : WebStyleBaseForEnergy.webStyleBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
#if DEBUG
            ////////////////////调试用,自定义的数据授权
            List<string> m_DataValidIdItems = new List<string>() { "zc_nxjc_byc", "zc_nxjc_qtx", "zc_nxjc_ychc" };
            AddDataValidIdGroup("ProductionOrganization", m_DataValidIdItems);
#elif RELEASE
#endif
            //以下是接收js脚本中post过来的参数
            string m_FunctionName = Request.Form["myFunctionName"] == null ? "" : Request.Form["myFunctionName"].ToString();             //方法名称,调用后台不同的方法
            string m_Parameter1 = Request.Form["myParameter1"] == null ? "" : Request.Form["myParameter1"].ToString();                   //方法的参数名称1
            string m_Parameter2 = Request.Form["myParameter2"] == null ? "" : Request.Form["myParameter2"].ToString();                   //方法的参数名称2
            if (m_FunctionName == "ExcelStream")
            {
                //ExportFile("xls", "导出报表1.xls");
                string m_ExportTable = m_Parameter1.Replace("&lt;", "<");
                m_ExportTable = m_ExportTable.Replace("&gt;", ">");
                ShuiNiFenMoFengGuPingConsumptionService.ExportExcelFile("xls", m_Parameter2 + "水泥制备峰谷平用电分析.xls", m_ExportTable);
            }
        }

        [WebMethod]
        public static string GetShuiNiFenMoFengGuPingConsumption(string m_searchDate)
        {
            List<string> m_OganizationIds = WebStyleBaseForEnergy.webStyleBase.GetDataValidIdGroup("ProductionOrganization");
            DataTable table = ShuiNiFenMoFengGuPingConsumptionService.GetShuiNiFenMoFengGuPingConsumption(m_searchDate, m_OganizationIds.ToArray());
            table.Columns.Remove("OrganizationId");
            table.Columns.Remove("runLevelChangedCount");
            string json = EasyUIJsonParser.DataGridJsonParser.DataTableToJson(table);
            return json;
        }
    }
}
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
    public partial class HuiZhuanYaoProductionCount : WebStyleBaseForEnergy.webStyleBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
#if DEBUG
            ////////////////////调试用,自定义的数据授权
            List<string> m_DataValidIdItems = new List<string>() { "zc_nxjc_byc", "zc_nxjc_qtx", "zc_nxjc_ychc", "zc_nxjc_szsc", "zc_nxjc_whsmc", "zc_nxjc_znc", "zc_nxjc_klqc", "zc_nxjc_tsc" };
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
                HuiZhuanYaoProductionCountService.ExportExcelFile("xls", m_Parameter2 + "回转窑系统产量统计.xls", m_ExportTable);
            }
        }

        [WebMethod]
        public static string GetConsumptionAnalysisInfo(string m_SelectTime)
        {
            List<string> m_OganizationIds = WebStyleBaseForEnergy.webStyleBase.GetDataValidIdGroup("ProductionOrganization");
            DataTable table = HuiZhuanYaoProductionCountService.GetConsumptionAnalysisTable(m_SelectTime, m_OganizationIds.ToArray());
            table.Columns["CompanyName"].SetOrdinal(0);
            table.Columns["ProductionLine"].SetOrdinal(1);
            table.Columns["A1"].SetOrdinal(2);
            table.Columns["A2"].SetOrdinal(3);
            table.Columns["A3"].SetOrdinal(4);
            table.Columns["A4"].SetOrdinal(5);
            table.Columns["A5"].SetOrdinal(6);
            table.Columns["A6"].SetOrdinal(7);
            table.Columns["A7"].SetOrdinal(8);
            table.Columns["A8"].SetOrdinal(9);
            table.Columns["A9"].SetOrdinal(10);//前期随后台table列处理不当，为了打印时不乱序，所以对列重新排序
            string json = EasyUIJsonParser.DataGridJsonParser.DataTableToJson(table);
            return json;
        }
    }
}
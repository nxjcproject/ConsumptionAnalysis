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
    public partial class MeiFenZhiBeiConsumption : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //以下是接收js脚本中post过来的参数
            string m_FunctionName = Request.Form["myFunctionName"] == null ? "" : Request.Form["myFunctionName"].ToString();             //方法名称,调用后台不同的方法
            string m_Parameter1 = Request.Form["myParameter1"] == null ? "" : Request.Form["myParameter1"].ToString();                   //方法的参数名称1
            string m_Parameter2 = Request.Form["myParameter2"] == null ? "" : Request.Form["myParameter2"].ToString();                   //方法的参数名称2
            if (m_FunctionName == "ExcelStream")
            {
                //ExportFile("xls", "导出报表1.xls");
                string m_ExportTable = m_Parameter1.Replace("&lt;", "<");
                m_ExportTable = m_ExportTable.Replace("&gt;", ">");
                MeiFenZhiBeiConsumptionService.ExportExcelFile("xls", m_Parameter2 + "煤粉制备系统电耗对比分析.xls", m_ExportTable);
            }
        }

        [WebMethod]
        public static string GetConsumptionAnalysisInfo(string mySelectTime)
        {
            DataTable table = MeiFenZhiBeiConsumptionService.GetConsumptionAnalysisTable(mySelectTime);
            string json = EasyUIJsonParser.DataGridJsonParser.DataTableToJson(table);
            return json;
        }
    }
}
﻿using System;
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
        }

        [WebMethod]
        public static string GetConsumptionAnalysisInfo(string m_SelectTime)
        {
            List<string> m_OganizationIds = WebStyleBaseForEnergy.webStyleBase.GetDataValidIdGroup("ProductionOrganization");
            DataTable table = HuiZhuanYaoProductionCountService.GetConsumptionAnalysisTable(m_SelectTime, m_OganizationIds.ToArray());
            string json = EasyUIJsonParser.DataGridJsonParser.DataTableToJson(table);
            return json;
        }
    }
}
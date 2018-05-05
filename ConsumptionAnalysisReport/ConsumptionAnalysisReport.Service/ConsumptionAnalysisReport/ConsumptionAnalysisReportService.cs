using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class ConsumptionAnalysisReportService
    {
        public static DataTable GetConsumptionAnalysisTable()
        {
            //建表
            DataTable table = new DataTable();
            //建列
            DataColumn dc1 = new DataColumn("month9", Type.GetType("System.Double"));
            DataColumn dc2 = new DataColumn("month10", Type.GetType("System.Double"));
            DataColumn dc3 = new DataColumn("month11", Type.GetType("System.Double"));
            DataColumn dc4 = new DataColumn("1month11", Type.GetType("System.Double"));
            DataColumn dc5 = new DataColumn("CompanyName", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("FenCompanyName", Type.GetType("System.String"));
            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            table.Columns.Add(dc5);
            table.Columns.Add(dc6);
            //创建每一行
            DataRow dr1 = table.NewRow();
            dr1["month9"] = 2.17;
            dr1["month10"] = 0.76;
            dr1["month11"] = DBNull.Value;
            dr1["1month11"] = DBNull.Value;
            dr1["CompanyName"] = "银川水泥";
            dr1["FenCompanyName"] = "兰山";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["month9"] = 2.4;
            dr2["month10"] = DBNull.Value;
            dr2["month11"] = DBNull.Value;
            dr2["1month11"] = DBNull.Value;
            dr2["CompanyName"] = "银川水泥";
            dr2["FenCompanyName"] = "一分厂";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["month9"] = 1.67;
            dr3["month10"] = 1.98;
            dr3["month11"] = 4.36;
            dr3["1month11"] = 2.18;
            dr3["CompanyName"] = "青铜峡水泥";
            dr3["FenCompanyName"] = "二分厂";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["month9"] = 1.75;
            dr4["month10"] = 1.91;
            dr4["month11"] = 2.65;
            dr4["1month11"] = 2.06;
            dr4["CompanyName"] = "青铜峡水泥";
            dr4["FenCompanyName"] = "太阳山";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["month9"] = 2.1;
            dr5["month10"] = 2.04;
            dr5["month11"] = 2.77;
            dr5["1month11"] = 1.41;
            dr5["CompanyName"] = "中宁水泥";
            dr5["FenCompanyName"] = "中宁水泥";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["month9"] = 2.37;
            dr6["month10"] = 2.52;
            dr6["month11"] = 2.98;
            dr6["1month11"] = 2.39;
            dr6["CompanyName"] = "六盘山水泥";
            dr6["FenCompanyName"] = "六盘山水泥";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["month9"] = 2.42;
            dr7["month10"] = 2.42;
            dr7["month11"] = 2.75;
            dr7["1month11"] = 1.98;
            dr7["CompanyName"] = "天水水泥";
            dr7["FenCompanyName"] = "天水水泥";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["month9"] = 2.11;
            dr8["month10"] = 2.33;
            dr8["month11"] = 2.01;
            dr8["1month11"] = 1.36;
            dr8["CompanyName"] = "白银水泥";
            dr8["FenCompanyName"] = "白银水泥";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["month9"] = 1.5;
            dr9["month10"] = 1.56;
            dr9["month11"] = 3.84;
            dr9["1month11"] = 2.70;
            dr9["CompanyName"] = "乌海赛马";
            dr9["FenCompanyName"] = "乌海赛马";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["month9"] = 2.6;
            dr10["month10"] = 2.61;
            dr10["month11"] = 4.96;
            dr10["1month11"] = 3.88;
            dr10["CompanyName"] = "喀喇沁水泥";
            dr10["FenCompanyName"] = "喀喇沁水泥";
            table.Rows.Add(dr10);

            return table;
        }
    }
}

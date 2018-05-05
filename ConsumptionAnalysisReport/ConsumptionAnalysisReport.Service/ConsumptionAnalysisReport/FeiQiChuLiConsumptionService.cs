using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class FeiQiChuLiConsumptionService
    {
        public static DataTable GetFeiQiChuLiConsumption(string mySelectTime)
        {
            string m_StartDate = mySelectTime + "-01";
            string m_EndDate = DateTime.Parse(m_StartDate).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
            List<string> m_OrganizationIds = new List<string>();
            DataTable m_VariableFormulaTable = GetConsumptionAnalysisTableFormula(m_OrganizationIds);

            //GetFixedConsumptionAnalysisTable(m_StartDate, m_EndDate);
            DataTable m_RealElectricityConsumptionTable = ComonFunction.ElectricityConsumptionTemplate1.GetElectricityConsumption(m_VariableFormulaTable, "ProductionLine", m_StartDate, m_EndDate);
            return m_RealElectricityConsumptionTable;
        }
        private static DataTable GetConsumptionAnalysisTableFormula(List<string> myOrganizationIds)
        {
            //建表
            DataTable table = new DataTable();
            //建列
            DataColumn dc1 = new DataColumn("CompanyName", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("ProductionLine", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("1", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("2", Type.GetType("System.String"));

            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);

            //创建每一行

            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "zc_nxjc_ychc_yfcf_clinker01";
            dr1["1"] = "kilnSystem";
            dr1["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr1);
         
            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker02";
            dr2["1"] = "kilnSystem";
            dr2["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker03";
            dr3["1"] = "kilnSystem";
            dr3["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr3);


            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker04";
            dr4["1"] = "kilnSystem";
            dr4["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr4);


            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_efc_clinker02";
            dr5["1"] = "kilnSystem";
            dr5["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr5);


            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "zc_nxjc_qtx_efc_clinker03";
            dr6["1"] = "kilnSystem";
            dr6["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr6);


            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "zc_nxjc_qtx_tys_clinker04";
            dr7["1"] = "kilnSystem";
            dr7["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "zc_nxjc_qtx_tys_clinker05";
            dr8["1"] = "kilnSystem";
            dr8["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr8);


            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "zc_nxjc_znc_znf_clinker01";
            dr9["1"] = "kilnSystem";
            dr9["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr9);


            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "zc_nxjc_znc_znf_clinker02";
            dr10["1"] = "kilnSystem";
            dr10["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr10);


            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "zc_nxjc_lpsc_lpsf_clinker01";
            dr11["1"] = "kilnSystem";
            dr11["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr11);


            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker01";
            dr12["1"] = "kilnSystem";
            dr12["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr12);


            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker02";
            dr13["1"] = "kilnSystem";
            dr13["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr13);


            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "zc_nxjc_whsmc_whsmf_clinker01";
            dr14["1"] = "kilnSystem";
            dr14["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr14);


            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "zc_nxjc_byc_byf_clinker01";
            dr15["1"] = "kilnSystem";
            dr15["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr15);


            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "zc_nxjc_klqc_klqf_clinker01";
            dr16["1"] = "kilnSystem";
            dr16["2"] = "kilnTailExhaustFan";
            table.Rows.Add(dr16);

            return table;
        }
        public static DataTable GetFixFeiQiChuLiConsumption()
        {
            //建表
            DataTable table = new DataTable();
            //建列
            DataColumn dc1 = new DataColumn("1", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("2", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("3", Type.GetType("System.Double"));
            DataColumn dc4 = new DataColumn("4", Type.GetType("System.Double"));

            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);

            //创建每一行

            DataRow dr1 = table.NewRow();
            dr1["1"] = "银川水泥";
            dr1["2"] = "1#";
            dr1["3"] = DBNull.Value;
            dr1["4"] = DBNull.Value;
            table.Rows.Add(dr1);


            DataRow dr2 = table.NewRow();
            dr2["1"] = "银川水泥";
            dr2["2"] = "2#";
            dr2["3"] = DBNull.Value;
            dr2["4"] = DBNull.Value;
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["1"] = "银川水泥";
            dr3["2"] = "3#";
            dr3["3"] = DBNull.Value;
            dr3["4"] = DBNull.Value;
            table.Rows.Add(dr3);


            DataRow dr4 = table.NewRow();
            dr4["1"] = "银川水泥";
            dr4["2"] = "4#";
            dr4["3"] = DBNull.Value;
            dr4["4"] = DBNull.Value;
            table.Rows.Add(dr4);


            DataRow dr5 = table.NewRow();
            dr5["1"] = "青铜峡水泥";
            dr5["2"] = "2#";
            dr5["3"] = 7.23;
            dr5["4"] = 4.32;
            table.Rows.Add(dr5);


            DataRow dr6 = table.NewRow();
            dr6["1"] = "青铜峡水泥";
            dr6["2"] = "3#";
            dr6["3"] = 4.02;
            dr6["4"] = 2.44;
            table.Rows.Add(dr6);


            DataRow dr7 = table.NewRow();
            dr7["1"] = "青铜峡水泥";
            dr7["2"] = "4#";
            dr7["3"] = 8.02;
            dr7["4"] = 7.71;
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["1"] = "青铜峡水泥";
            dr8["2"] = "5#";
            dr8["3"] = 5.57;
            dr8["4"] = 5.29;
            table.Rows.Add(dr8);


            DataRow dr9 = table.NewRow();
            dr9["1"] = "中宁水泥";
            dr9["2"] = "1#";
            dr9["3"] = 4.96;
            dr9["4"] = 3.47;
            table.Rows.Add(dr9);


            DataRow dr10 = table.NewRow();
            dr10["1"] = "中宁水泥";
            dr10["2"] = "2#";
            dr10["3"] = 4.56;
            dr10["4"] = 2.92;
            table.Rows.Add(dr10);


            DataRow dr11 = table.NewRow();
            dr11["1"] = "六盘山水泥";
            dr11["2"] = "1#";
            dr11["3"] = 8.00;
            dr11["4"] = 8.24;
            table.Rows.Add(dr11);


            DataRow dr12 = table.NewRow();
            dr12["1"] = "天水水泥";
            dr12["2"] = "1#";
            dr12["3"] = 6.57;
            dr12["4"] = 6.49;
            table.Rows.Add(dr12);


            DataRow dr13 = table.NewRow();
            dr13["1"] = "天水水泥";
            dr13["2"] = "2#";
            dr13["3"] = 6.36;
            dr13["4"] = 6.24;
            table.Rows.Add(dr13);


            DataRow dr14 = table.NewRow();
            dr14["1"] = "乌海赛马";
            dr14["2"] = "1#";
            dr14["3"] = 5.23;
            dr14["4"] = 5.05;
            table.Rows.Add(dr14);


            DataRow dr15 = table.NewRow();
            dr15["1"] = "白银水泥";
            dr15["2"] = "1#";
            dr15["3"] = 5.21;
            dr15["4"] = 4.87;
            table.Rows.Add(dr15);


            DataRow dr16 = table.NewRow();
            dr16["1"] = "喀喇沁水泥";
            dr16["2"] = "1#";
            dr16["3"] = 3.99;
            dr16["4"] = 3.74;
            table.Rows.Add(dr16);

            return table;
        }
    }
}

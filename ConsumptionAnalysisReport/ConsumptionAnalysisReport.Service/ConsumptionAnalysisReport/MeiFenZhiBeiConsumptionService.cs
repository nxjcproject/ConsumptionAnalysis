using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using ConsumptionAnalysisReport.Infrastructure.Configuration;
using SqlServerDataAdapter;
namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class MeiFenZhiBeiConsumptionService
    {
        public static DataTable GetConsumptionAnalysisTable(string myDate)
        {
            string m_StartDate = myDate + "-01";
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
            DataColumn dc5 = new DataColumn("3", Type.GetType("System.String"));
            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            table.Columns.Add(dc5);

            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "zc_nxjc_ychc_yfcf_clinker01";
            dr1["1"] = "coalPreparation";
            dr1["2"] = "coalMillMainMotor";
            dr1["3"] = "coalMillFan";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker02";
            dr2["1"] = "coalPreparation";
            dr2["2"] = "coalMillMainMotor";
            dr2["3"] = "coalMillFan";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker03";
            dr3["1"] = "coalPreparation";
            dr3["2"] = "coalMillMainMotor";
            dr3["3"] = "coalMillFan";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker04";
            dr4["1"] = "coalPreparation";
            dr4["2"] = "coalMillMainMotor";
            dr4["3"] = "coalMillFan";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_efc_clinker02";
            dr5["1"] = "coalPreparation";
            dr5["2"] = "coalMillMainMotor";
            dr5["3"] = "coalMillFan";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "zc_nxjc_qtx_efc_clinker03";
            dr6["1"] = "coalPreparation";
            dr6["2"] = "coalMillMainMotor";
            dr6["3"] = "coalMillFan";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "zc_nxjc_qtx_tys_clinker04";
            dr7["1"] = "coalPreparation";
            dr7["2"] = "coalMillMainMotor";
            dr7["3"] = "coalMillFan";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "zc_nxjc_qtx_tys_clinker05";
            dr8["1"] = "coalPreparation";
            dr8["2"] = "coalMillMainMotor";
            dr8["3"] = "coalMillFan";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "zc_nxjc_znc_znf_clinker01";
            dr9["1"] = "coalPreparation";
            dr9["2"] = "coalMillMainMotor";
            dr9["3"] = "coalMillFan";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "zc_nxjc_znc_znf_clinker02";
            dr10["1"] = "coalPreparation";
            dr10["2"] = "coalMillMainMotor";
            dr10["3"] = "coalMillFan";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "zc_nxjc_lpsc_lpsf_clinker01";
            dr11["1"] = "coalPreparation";
            dr11["2"] = "coalMillMainMotor";
            dr11["3"] = "coalMillFan";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker01";
            dr12["1"] = "coalPreparation";
            dr12["2"] = "coalMillMainMotor";
            dr12["3"] = "coalMillFan";
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker02";
            dr13["1"] = "coalPreparation";
            dr13["2"] = "coalMillMainMotor";
            dr13["3"] = "coalMillFan";
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "zc_nxjc_whsmc_whsmf_clinker01";
            dr14["1"] = "coalPreparation";
            dr14["2"] = "coalMillMainMotor";
            dr14["3"] = "coalMillFan";
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "zc_nxjc_byc_byf_clinker01";
            dr15["1"] = "coalPreparation";
            dr15["2"] = "coalMillMainMotor";
            dr15["3"] = "coalMillFan";
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "zc_nxjc_klqc_klqf_clinker01";
            dr16["1"] = "coalPreparation";
            dr16["2"] = "discMillRoll";
            dr16["3"] = "coalMillFan";
            table.Rows.Add(dr16);
            return table;
        }

        
        private static DataTable GetFixedConsumptionAnalysisTable(string myStartDate, string myEndDate)
        {
            //建表
            DataTable table = new DataTable();
            //建列
            DataColumn dc1 = new DataColumn("CompanyName", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("ProductionLine", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("1", Type.GetType("System.Double"));
            DataColumn dc4 = new DataColumn("2", Type.GetType("System.Double"));
            DataColumn dc5 = new DataColumn("3", Type.GetType("System.Double"));
            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            table.Columns.Add(dc5);

            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "1#";
            dr1["1"] = DBNull.Value;
            dr1["2"] = DBNull.Value;
            dr1["3"] = DBNull.Value;
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["1"] = DBNull.Value;
            dr2["2"] = DBNull.Value;
            dr2["3"] = DBNull.Value;
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["1"] = DBNull.Value;
            dr3["2"] = DBNull.Value;
            dr3["3"] = DBNull.Value;
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1"] = DBNull.Value;
            dr4["2"] = DBNull.Value;
            dr4["3"] = DBNull.Value;
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["1"] = 73.58;
            dr5["2"] = 13.55;
            dr5["3"] = 28.29;
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["1"] = 64.43;
            dr6["2"] = 15.93;
            dr6["3"] = 21.97;
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["1"] = 45.47;
            dr7["2"] = 18.96;
            dr7["3"] = 22.49;
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["1"] = 45.01;
            dr8["2"] = 17.91;
            dr8["3"] = 22.83;
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1"] = 32.62;
            dr9["2"] = 13.97;
            dr9["3"] = 18.76;
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1"] = 31.64;
            dr10["2"] = 28.96;
            dr10["3"] = 9.31;
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1"] = 41.26;
            dr11["2"] = 23.19;
            dr11["3"] = 13.04;
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["1"] = 32.99;
            dr12["2"] = 13.24;
            dr12["3"] = 16.11;
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["1"] = 29.28;
            dr13["2"] = 12.06;
            dr13["3"] = 13.93;
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "1#";
            dr14["1"] = 37.28;
            dr14["2"] = 15.02;
            dr14["3"] = 18.25;
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["1"] = 31.06;
            dr15["2"] = 12.54;
            dr15["3"] = 14.33;
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["1"] = 32.60;
            dr16["2"] = 16.25;
            dr16["3"] = 14.64;
            table.Rows.Add(dr16);

            return table;
        }
    }
}

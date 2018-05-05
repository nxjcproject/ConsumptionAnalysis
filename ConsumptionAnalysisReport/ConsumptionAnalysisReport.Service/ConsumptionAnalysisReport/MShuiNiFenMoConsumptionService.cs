using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class MShuiNiFenMoConsumptionService
    {
        public static DataTable GetMShuiNiFenMoConsumption(string mySelectTime)
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
            DataColumn dc5 = new DataColumn("3", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("4", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("5", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("6", Type.GetType("System.String"));
            DataColumn dc9 = new DataColumn("7", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("8", Type.GetType("System.String"));

            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            table.Columns.Add(dc5);
            table.Columns.Add(dc6);
            table.Columns.Add(dc7);
            table.Columns.Add(dc8);
            table.Columns.Add(dc9);
            table.Columns.Add(dc10);

            //创建每一行
            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "zc_nxjc_ychc_lsf_cementmill08";
            dr1["1"] = "cementGrind";
            dr1["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr1["3"] = DBNull.Value;
            dr1["4"] = DBNull.Value;
            dr1["5"] = "cementMillMainMotor";
            dr1["6"] = DBNull.Value;
            dr1["7"] = "cementMillHoist";
            dr1["8"] = "mainExhaustFan";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_cementmill09";
            dr2["1"] = "cementGrind";
            dr2["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr2["3"] = DBNull.Value;
            dr2["4"] = DBNull.Value;
            dr2["5"] = "cementMillMainMotor";
            dr2["6"] = DBNull.Value;
            dr2["7"] = "cementMillHoist";
            dr2["8"] = "mainExhaustFan";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "青铜峡水泥";
            dr3["ProductionLine"] = "zc_nxjc_qtx_efc_cementmill01";
            dr3["1"] = "cementGrind";
            dr3["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr3["3"] = "doubleDriveHoist1 + doubleDriveHoist2";
            dr3["4"] = DBNull.Value;
            dr3["5"] = "cementMillMainMotor";
            dr3["6"] = "powderSelectingStorehouse";
            dr3["7"] = "cementOutputHoist";
            dr3["8"] = "mainExhaustFan";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "青铜峡水泥";
            dr4["ProductionLine"] = "zc_nxjc_qtx_tys_cementmill04";
            dr4["1"] = "cementGrind";
            dr4["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr4["3"] = "hoist55107";
            dr4["4"] = "rollerPressCirculatingFan";  
            dr4["5"] = "cementMillMainMotor";
            dr4["6"] = "powderSelectingStorehouse";
            dr4["7"] = "hoist55120";
            dr4["8"] = "mainExhaustFan";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_tys_cementmill05";
            dr5["1"] = "cementGrind";
            dr5["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr5["3"] = "hoist551011 + hoist551012";
            dr5["4"] = DBNull.Value;
            dr5["5"] = "cementMillMainMotor";
            dr5["6"] = "powderSelectingStorehouse";
            dr5["7"] = DBNull.Value;
            dr5["8"] = "mainExhaustFan";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "天水水泥";
            dr6["ProductionLine"] = "zc_nxjc_tsc_tsf_cementmill01";
            dr6["1"] = "cementGrind";
            dr6["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr6["3"] = "cementMillHoist1";
            dr6["4"] = "circulatingFan";
            dr6["5"] = "cementMillMainMotor";
            dr6["6"] = "powderSelectingStorehouse";
            dr6["7"] = "cementMillHoist2";
            dr6["8"] = "mainExhaustFan";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "天水水泥";
            dr7["ProductionLine"] = "zc_nxjc_tsc_tsf_cementmill02";
            dr7["1"] = "cementGrind";
            dr7["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr7["3"] = "cementMillHoist1";
            dr7["4"] = "circulatingFan";
            dr7["5"] = "cementMillMainMotor";
            dr7["6"] = "powderSelectingStorehouse";
            dr7["7"] = "cementMillHoist2";
            dr7["8"] = "mainExhaustFan";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "乌海赛马";
            dr8["ProductionLine"] = "zc_nxjc_whsmc_whsmf_cementmill01";
            dr8["1"] = "cementGrind";
            dr8["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr8["3"] = "hoist551071";
            dr8["4"] = "circulatingFan";
            dr8["5"] = "cementMillMainMotor";
            dr8["6"] = "powderSelectingMachine";
            dr8["7"] = "hoist551201";
            dr8["8"] = "cementExhaustFan";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "白银水泥";
            dr9["ProductionLine"] = "zc_nxjc_byc_byf_cementmill01";
            dr9["1"] = "cementGrind";
            dr9["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr9["3"] = "circleHoist1 + circleHoist2";
            dr9["4"] = DBNull.Value;
            dr9["5"] = "cementMillMainMotor";
            dr9["6"] = "powderSelectingStorehouse";
            dr9["7"] = DBNull.Value;
            dr9["8"] = "mainExhaustFan";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "白银水泥";
            dr10["ProductionLine"] = "zc_nxjc_byc_byf_cementmill02";
            dr10["1"] = "cementGrind";
            dr10["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr10["3"] = "circleHoist1 + circleHoist2";
            dr10["4"] = DBNull.Value;
            dr10["5"] = "cementMillMainMotor";
            dr10["6"] = "powderSelectingStorehouse";
            dr10["7"] = DBNull.Value;
            dr10["8"] = "mainExhaustFan";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "喀喇沁水泥";
            dr11["ProductionLine"] = "zc_nxjc_klqc_klqf_cementmill01";
            dr11["1"] = "cementGrind";
            dr11["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr11["3"] = "hoist84A01 + hoist84A07";
            dr11["4"] = DBNull.Value;
            dr11["5"] = "cementMillMainMotor";
            dr11["6"] = "powderSelectingStorehouse";
            dr11["7"] = DBNull.Value;
            dr11["8"] = "mainExhaustFan";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "喀喇沁水泥";
            dr12["ProductionLine"] = "zc_nxjc_klqc_klqf_cementmill02";
            dr12["1"] = "cementGrind";
            dr12["2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr12["3"] = "hoist84B07 + hoist84B01";
            dr12["4"] = DBNull.Value;
            dr12["5"] = "cementMillMainMotor";
            dr12["6"] = "powderSelectingStorehouse";
            dr12["7"] = DBNull.Value;
            dr12["8"] = "mainExhaustFan";
            table.Rows.Add(dr12);

            return table;
        }

        public static DataTable GetFixMShuiNiFenMoConsumption()
        {
            //建表
            DataTable table = new DataTable();
            //建列
            DataColumn dc1 = new DataColumn("1", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("2", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("3", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("4", Type.GetType("System.String"));
            DataColumn dc5 = new DataColumn("5", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("6", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("7", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("8", Type.GetType("System.String"));
            DataColumn dc9 = new DataColumn("9", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("10", Type.GetType("System.String"));

            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            table.Columns.Add(dc5);
            table.Columns.Add(dc6);
            table.Columns.Add(dc7);
            table.Columns.Add(dc8);
            table.Columns.Add(dc9);
            table.Columns.Add(dc10);


            //创建每一行

            DataRow dr1 = table.NewRow();
            dr1["1"] = "银川水泥";
            dr1["2"] = "8#";
            dr1["3"] = DBNull.Value;
            dr1["4"] = DBNull.Value;
            dr1["5"] = "-";
            dr1["6"] = "-";
            dr1["7"] = DBNull.Value;
            dr1["8"] = "-";
            dr1["9"] = DBNull.Value;
            dr1["10"] = DBNull.Value;
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["1"] = "银川水泥";
            dr2["2"] = "9#";
            dr2["3"] = DBNull.Value;
            dr2["4"] = DBNull.Value;
            dr2["5"] = "-";
            dr2["6"] = "-";
            dr2["7"] = DBNull.Value;
            dr2["8"] = "-";
            dr2["9"] = DBNull.Value;
            dr2["10"] = DBNull.Value;
            table.Rows.Add(dr2);




            DataRow dr3 = table.NewRow();
            dr3["1"] = "青铜峡水泥";
            dr3["2"] = "1#";
            dr3["3"] = "57.83";
            dr3["4"] = "57.83";
            dr3["5"] = "0.79";
            dr3["6"] = "-";
            dr3["7"] = "19.41";
            dr3["8"] = "1.07";
            dr3["9"] = "0.44";
            dr3["10"] = "6.75";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["1"] = "青铜峡水泥";
            dr4["2"] = "4#";
            dr4["3"] = "36.00";
            dr4["4"] = "9.22";
            dr4["5"] = "0.38";
            dr4["6"] = "1.44";
            dr4["7"] = "20.26";
            dr4["8"] = "0.32";
            dr4["9"] = "0.51";
            dr4["10"] = "1.41";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["1"] = "青铜峡水泥";
            dr5["2"] = "5#";
            dr5["3"] = "37.37";
            dr5["4"] = "7.94";
            dr5["5"] = "1.16";
            dr5["6"] = "-";
            dr5["7"] = "21.97";
            dr5["8"] = "0.99";
            dr5["9"] = "-";
            dr5["10"] = "2.84";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["1"] = "天水水泥";
            dr6["2"] = "1#";
            dr6["3"] = "41.37";
            dr6["4"] = "8.75";
            dr6["5"] = "1.18";
            dr6["6"] = "1.88";
            dr6["7"] = "22.24";
            dr6["8"] = "0.65";
            dr6["9"] = "0.42";
            dr6["10"] = "3.77";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["1"] = "天水水泥";
            dr7["2"] = "2#";
            dr7["3"] = "27.87";
            dr7["4"] = "3.70";
            dr7["5"] = "0.39";
            dr7["6"] = "1.17";
            dr7["7"] = "16.74";
            dr7["8"] = "0.45";
            dr7["9"] = "0.29";
            dr7["10"] = "2.96";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["1"] = "乌海赛马";
            dr8["2"] = "1#";
            dr8["3"] = "33.23";
            dr8["4"] = "8.38";
            dr8["5"] = "0.96";
            dr8["6"] = "1.73";
            dr8["7"] = "22.76";
            dr8["8"] = "0.52";
            dr8["9"] = "0.46";
            dr8["10"] = "2.16";
            table.Rows.Add(dr8);


            DataRow dr9 = table.NewRow();
            dr9["1"] = "白银水泥";
            dr9["2"] = "1#";
            dr9["3"] = "35.40";
            dr9["4"] = "8.60";
            dr9["5"] = "0.93";
            dr9["6"] = "-";
            dr9["7"] = "18.62";
            dr9["8"] = "0.41";
            dr9["9"] = "-";
            dr9["10"] = "5.25";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["1"] = "白银水泥";
            dr10["2"] = "2#";
            dr10["3"] = "34.63";
            dr10["4"] = "7.01";
            dr10["5"] = "0.80";
            dr10["6"] = "-";
            dr10["7"] = "18.43";
            dr10["8"] = "0.48";
            dr10["9"] = "-";
            dr10["10"] = "5.39";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["1"] = "喀喇沁水泥";
            dr11["2"] = "1#";
            dr11["3"] = "35.25";
            dr11["4"] = "11.10";
            dr11["5"] = "0.57";
            dr11["6"] = "-";
            dr11["7"] = "14.85";
            dr11["8"] = "0.80";
            dr11["9"] = "-";
            dr11["10"] = "4.82";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["1"] = "喀喇沁水泥";
            dr12["2"] = "2#";
            dr12["3"] = "停运";
            dr12["4"] = "停运";
            dr12["5"] = "停运";
            dr12["6"] = "-";
            dr12["7"] = "停运";
            dr12["8"] = "停运";
            dr12["9"] = "-";
            dr12["10"] = "停运";
            table.Rows.Add(dr12);

            return table;
        }
    }
}

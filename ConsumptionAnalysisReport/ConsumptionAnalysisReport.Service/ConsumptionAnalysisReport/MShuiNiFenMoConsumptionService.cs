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
            DataColumn dc3 = new DataColumn("A1", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("A2", Type.GetType("System.String"));
            DataColumn dc5 = new DataColumn("A3", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("A4", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("A5", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("A6", Type.GetType("System.String"));
            DataColumn dc9 = new DataColumn("A7", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("A8", Type.GetType("System.String"));

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
            dr1["A1"] = "cementGrind";
            dr1["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr1["A3"] = "entryRollerHoist";
            dr1["A4"] = "mainExhaustFan";
            dr1["A5"] = "cementMillMainMotor";
            dr1["A6"] = "powderSelectingMachine";
            dr1["A7"] = "cementMillHoist";
            dr1["A8"] = "mainExhaustFan";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_cementmill09";
            dr2["A1"] = "cementGrind";
            dr2["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr2["A3"] = "entryRollerHoist";
            dr2["A4"] = "mainExhaustFan";
            dr2["A5"] = "cementMillMainMotor";
            dr2["A6"] = "powderSelectingMachine";
            dr2["A7"] = "cementMillHoist";
            dr2["A8"] = "mainExhaustFan";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "青铜峡水泥";
            dr3["ProductionLine"] = "zc_nxjc_qtx_efc_cementmill01";
            dr3["A1"] = "cementGrind";
            dr3["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr3["A3"] = "doubleDriveHoist1 + doubleDriveHoist2";
            dr3["A4"] = DBNull.Value;
            dr3["A5"] = "cementMillMainMotor";
            dr3["A6"] = "powderSelectingStorehouse";
            dr3["A7"] = "cementOutputHoist";
            dr3["A8"] = "mainExhaustFan";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "青铜峡水泥";
            dr4["ProductionLine"] = "zc_nxjc_qtx_tys_cementmill04";
            dr4["A1"] = "cementGrind";
            dr4["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr4["A3"] = "hoist55107";
            dr4["A4"] = "rollerPressCirculatingFan";  
            dr4["A5"] = "cementMillMainMotor";
            dr4["A6"] = "powderSelectingStorehouse";
            dr4["A7"] = "hoist55120";
            dr4["A8"] = "mainExhaustFan";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_tys_cementmill05";
            dr5["A1"] = "cementGrind";
            dr5["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr5["A3"] = "hoist551011 + hoist551012";
            dr5["A4"] = DBNull.Value;
            dr5["A5"] = "cementMillMainMotor";
            dr5["A6"] = "powderSelectingStorehouse";
            dr5["A7"] = "Hoist5511";
            dr5["A8"] = "mainExhaustFan";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "天水水泥";
            dr6["ProductionLine"] = "zc_nxjc_tsc_tsf_cementmill01";
            dr6["A1"] = "cementGrind";
            dr6["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr6["A3"] = "cementMillHoist1";
            dr6["A4"] = "circulatingFan";
            dr6["A5"] = "cementMillMainMotor";
            dr6["A6"] = "powderSelectingStorehouse";
            dr6["A7"] = "cementMillHoist2";
            dr6["A8"] = "mainExhaustFan";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "天水水泥";
            dr7["ProductionLine"] = "zc_nxjc_tsc_tsf_cementmill02";
            dr7["A1"] = "cementGrind";
            dr7["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr7["A3"] = "cementMillHoist1";
            dr7["A4"] = "circulatingFan";
            dr7["A5"] = "cementMillMainMotor";
            dr7["A6"] = "powderSelectingStorehouse";
            dr7["A7"] = "cementMillHoist2";
            dr7["A8"] = "mainExhaustFan";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "乌海赛马";
            dr8["ProductionLine"] = "zc_nxjc_whsmc_whsmf_cementmill01";
            dr8["A1"] = "cementGrind";
            dr8["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr8["A3"] = "hoist551071";
            dr8["A4"] = "circulatingFan";
            dr8["A5"] = "cementMillMainMotor";
            dr8["A6"] = "powderSelectingMachine";
            dr8["A7"] = "hoist551201";
            dr8["A8"] = "cementExhaustFan";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "白银水泥";
            dr9["ProductionLine"] = "zc_nxjc_byc_byf_cementmill01";
            dr9["A1"] = "cementGrind";
            dr9["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr9["A3"] = "circleHoist1 + circleHoist2";
            dr9["A4"] = DBNull.Value;
            dr9["A5"] = "cementMillMainMotor";
            dr9["A6"] = "powderSelectingStorehouse";
            dr9["A7"] = "cementOutputHoist";
            dr9["A8"] = "mainExhaustFan";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "白银水泥";
            dr10["ProductionLine"] = "zc_nxjc_byc_byf_cementmill02";
            dr10["A1"] = "cementGrind";
            dr10["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr10["A3"] = "circleHoist1 + circleHoist2";
            dr10["A4"] = DBNull.Value;
            dr10["A5"] = "cementMillMainMotor";
            dr10["A6"] = "powderSelectingStorehouse";
            dr10["A7"] = "cementOutputHoist";
            dr10["A8"] = "mainExhaustFan";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "喀喇沁水泥";
            dr11["ProductionLine"] = "zc_nxjc_klqc_klqf_cementmill01";
            dr11["A1"] = "cementGrind";
            dr11["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr11["A3"] = "hoist84A01 + hoist84A07";
            dr11["A4"] = DBNull.Value;
            dr11["A5"] = "cementMillMainMotor";
            dr11["A6"] = "powderSelectingStorehouse";
            dr11["A7"] = DBNull.Value;
            dr11["A8"] = "mainExhaustFan";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "喀喇沁水泥";
            dr12["ProductionLine"] = "zc_nxjc_klqc_klqf_cementmill02";
            dr12["A1"] = "cementGrind";
            dr12["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr12["A3"] = "hoist84B07 + hoist84B01";
            dr12["A4"] = DBNull.Value;
            dr12["A5"] = "cementMillMainMotor";
            dr12["A6"] = "powderSelectingStorehouse";
            dr12["A7"] = DBNull.Value;
            dr12["A8"] = "mainExhaustFan";
            table.Rows.Add(dr12);

            return table;
        }

        public static void ExportExcelFile(string myFileType, string myFileName, string myData)
        {
            if (myFileType == "xls")
            {
                UpDownLoadFiles.DownloadFile.ExportExcelFile(myFileName, myData);
            }
        }      
    }
}

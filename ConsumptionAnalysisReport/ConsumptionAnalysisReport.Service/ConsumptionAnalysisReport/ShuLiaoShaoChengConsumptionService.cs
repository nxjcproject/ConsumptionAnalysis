using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class ShuLiaoShaoChengConsumptionService
    {
        public static DataTable GetConsumptionAnalysisTable(string mySelectTime)
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
            DataColumn dc11 = new DataColumn("9", Type.GetType("System.String"));
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
            table.Columns.Add(dc11);

            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "zc_nxjc_ychc_yfcf_clinker01";
            dr1["1"] = "clinkerBurning";            //熟料烧成工序
            dr1["2"] = "rectifierTransformer";           //窑主机
            dr1["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";            //均化库风机 
            dr1["4"] = "clinkerHoist";            //入窑提升机
            dr1["5"] = "oneTimeFan01";         //一次风机
            dr1["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";            //送煤风机
            dr1["7"] = "highTemperatureFan";            //高温风机
            dr1["8"] = "kilnHeadExhaustFan";            //头排风机
            dr1["9"] = "clinkerF9AC + clinkerF11AC + clinkerF12AC + clinkerF5AC + clinkerF8AC + clinkerF2AC"
                       + " + clinkerF6AC + clinkerF7AC + clinkerF10AC + clinkerF1AC + clinkerF3AC + clinkerF13AC";          //篦冷机风机
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker02";
            dr2["1"] = "clinkerBurning";
            dr2["2"] = DBNull.Value;
            dr2["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr2["4"] = "clinkerHoist";
            dr2["5"] = "oneTimeFan01 + oneTimeFan02";
            dr2["6"] = DBNull.Value;
            dr2["7"] = "highTemperatureFan";
            dr2["8"] = "kilnHeadExhaustFan";
            dr2["9"] = "clinkerF1A1 + clinkerF1A2 + clinkerF5AC + clinkerF3AC + clinkerF2AC";        //将来加表
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker03";
            dr3["1"] = "clinkerBurning";
            dr3["2"] = "kilnMainMotor";
            dr3["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr3["4"] = "clinkerHoist";
            dr3["5"] = "oneTimeFan01 + oneTimeFan02";
            dr3["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr3["7"] = "highTemperatureFan";
            dr3["8"] = "kilnHeadExhaustFan";
            dr3["9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC"
                       + " + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC + clinkerF13AC";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker04";
            dr4["1"] = "clinkerBurning";
            dr4["2"] = "kilnMainMotor";
            dr4["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr4["4"] = "clinkerHoist1 + clinkerHoist2";
            dr4["5"] = "oneTimeFan01 + oneTimeFan02";
            dr4["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr4["7"] = "highTemperatureFan";
            dr4["8"] = "kilnHeadExhaustFan";
            dr4["9"] = "clinkerF1AC + clinkerF3AC + clinkerF4AC + clinkerF2AC + clinkerF5AC + clinkerF6AC + clinkerF7AC"
                       + " + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_efc_clinker02";
            dr5["1"] = "clinkerBurning";
            dr5["2"] = "kilnMainMotor";
            dr5["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr5["4"] = "clinkerHoist";
            dr5["5"] = "oneTimeFan01 + oneTimeFan02";
            dr5["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr5["7"] = "highTemperatureFan";
            dr5["8"] = "kilnHeadExhaustFan";
            dr5["9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "zc_nxjc_qtx_efc_clinker03";
            dr6["1"] = "clinkerBurning";   //熟料烧成工序
            dr6["2"] = "kilnMainMotor";     //窑主机
            dr6["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";   //均化库风机 
            dr6["4"] = "clinkerHoist";     //入窑提升机
            dr6["5"] = DBNull.Value;       //一次风机
            dr6["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";     //送煤风机
            dr6["7"] = "highTemperatureFan";        //高温风机
            dr6["8"] = "kilnHeadExhaustFan";       //头排风机
            dr6["9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";       //篦冷机风机
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "zc_nxjc_qtx_tys_clinker04";
            dr7["1"] = "clinkerBurning";
            dr7["2"] = "kilnMainMotor";
            dr7["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr7["4"] = "clinkerHoist";
            dr7["5"] = "hootsBlower33020 + hootsBlower33021";
            dr7["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr7["7"] = "highTemperatureFan";
            dr7["8"] = "kilnHeadExhaustFan";
            dr7["9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "zc_nxjc_qtx_tys_clinker05";
            dr8["1"] = "clinkerBurning";
            dr8["2"] = "kilnMainMotor";
            dr8["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr8["4"] = "clinkerHoist";
            dr8["5"] = DBNull.Value;
            dr8["6"] = DBNull.Value;
            dr8["7"] = "highTemperatureFan";
            dr8["8"] = "kilnHeadExhaustFan";
            dr8["9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "zc_nxjc_znc_znf_clinker01";
            dr9["1"] = "clinkerBurning";
            dr9["2"] = "kilnMainMotor";
            dr9["3"] = DBNull.Value;
            dr9["4"] = "clinkerHoist";
            dr9["5"] = "netWindFan";
            dr9["6"] = DBNull.Value;
            dr9["7"] = "highTemperatureFan";
            dr9["8"] = "leftFan";
            dr9["9"] = "fanN1 + fanN2 + fanN3 + fanN4 + fanN5 + fanN6 + fanN7";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "zc_nxjc_znc_znf_clinker02";
            dr10["1"] = "clinkerBurning";        //熟料烧成工序
            dr10["2"] = "kilnMainMotor"; //窑主机
            dr10["3"] = DBNull.Value; //均化库风机 
            dr10["4"] = "clinkerHoist"; //入窑提升机
            dr10["5"] = "netWindFan";    //一次风机
            dr10["6"] = "rootsBlower7316 + kilnTailRotorScale";  //送煤风机
            dr10["7"] = "highTemperatureFan";  //高温风机
            dr10["8"] = "leftFan";     //头排风机
            dr10["9"] = "fanV1 + fanV2 + fanV3 + fanV4 + fanV5 + fanV6 + fanV7 + fanV8";    //篦冷机风机
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "zc_nxjc_lpsc_lpsf_clinker01";
            dr11["1"] = "clinkerBurning";
            dr11["2"] = "kilnMainMotor";
            dr11["3"] = DBNull.Value;
            dr11["4"] = "clinkerHoist";
            dr11["5"] = DBNull.Value;
            dr11["6"] = DBNull.Value;
            dr11["7"] = "highTemperatureFan";
            dr11["8"] = "kilnHeadExhaustFan";
            dr11["9"] = "clinkerF2M + clinkerF3M + clinkerF4M + clinkerF5M + clinkerF6M + clinkerF7M + clinker5730 + clinker5732 + coalMillFan";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker01";
            dr12["1"] = "clinkerBurning";
            dr12["2"] = "kilnMainMotor";
            dr12["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr12["4"] = "kilnHoist";
            dr12["5"] = "oneTimeFan01 + oneTimeFan02 + oneTimeFan03";
            dr12["6"] = "coalMilRootsBlower13 + coalMilRootsBlower14 + coalMilRootsBlower15";
            dr12["7"] = "highTemperatureFan";
            dr12["8"] = "kilnHeadExhaustFan";
            dr12["9"] = "grateCoolerFan11 + grateCoolerFan10 + grateCoolerFan07 + grateCoolerFan13 + grateCoolerFan12 + grateCoolerFan03 + grateCoolerFan04 + grateCoolerFan05 + grateCoolerFan06 + grateCoolerFan08 + grateCoolerFan09";
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker02";
            dr13["1"] = "clinkerBurning";
            dr13["2"] = "kilnMainMotor";
            dr13["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr13["4"] = "kilnHoist";
            dr13["5"] = "oneTimeFan01 + oneTimeFan02 + oneTimeFan03";
            dr13["6"] = "coalMilRootsBlower14 + coalMilRootsBlower13 + coalMilRootsBlower15";
            dr13["7"] = "highTemperatureFan";
            dr13["8"] = "kilnHeadExhaustFan";
            dr13["9"] = "grateCoolerFan07 + grateCoolerFan10 + grateCoolerFan13 + grateCoolerFan11 + grateCoolerFan12 + grateCoolerFan08 + grateCoolerFan09 + grateCoolerFan06 + grateCoolerFan05 + grateCoolerFan03 + grateCoolerFan04";
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "zc_nxjc_whsmc_whsmf_clinker01";
            dr14["1"] = "clinkerBurning";
            dr14["2"] = "kilnMainMotor";
            dr14["3"] = "RootsBlower22015 + RootsBlower22016 + RootsBlower22118";
            dr14["4"] = "kilnHoist";
            dr14["5"] = "oneTimeFan33020 + oneTimeFan33021";
            dr14["6"] = "RootsBlower43113 + RootsBlower43114 + RootsBlower43115";
            dr14["7"] = "highTemperatureFan";
            dr14["8"] = "kilnHeadExhaustFan";
            dr14["9"] = "coolingFanV3 + coolingFanV1 + coolingFanV2 + coolingFanV6 + coolingFanV7 + coolingFanV4 + coolingFanV5 + coolingFanV8 + coolingFanV10 + coolingFanV11 + coolingFanV9";
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "zc_nxjc_byc_byf_clinker01";
            dr15["1"] = "clinkerBurning";
            dr15["2"] = "kilnMainMotor";
            dr15["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3 + rawMaterialLibraryRootsBlower4";
            dr15["4"] = "clinkerHoist1 + clinkerHoist2";
            dr15["5"] = "oneTimeFan01 + oneTimeFan02";
            dr15["6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr15["7"] = "highTemperatureFan";
            dr15["8"] = "kilnHeadExhaustFan";
            dr15["9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC + clinkerF12AC + clinkerFVOA + clinkerFVOB + clinkerFVOC";
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "zc_nxjc_klqc_klqf_clinker01";
            dr16["1"] = "clinkerBurning";
            dr16["2"] = "kilnMainMotor";
            dr16["3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3 + rawMaterialLibraryRootsBlower4";
            dr16["4"] = "kilnHoist + kilnHoist1";
            dr16["5"] = "rootsBlower5704 + rootsBlower5705";
            dr16["6"] = "rootsBlower7513 + rootsBlower7514 + rootsBlower7515";
            dr16["7"] = "highTemperatureFan";
            dr16["8"] = "kilnHeadExhaustFan";
            dr16["9"] = "coolingFan5708 + coolingFan5714 + coolingFan5715 + coolingFan5712 + coolingFan5711 + coolingFan5713 + coolingFanF2 + coolingFanF3";
            table.Rows.Add(dr16);

            return table;
        }
        public static DataTable GetFixShengLiaoFenMoDianHao()
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
            DataColumn dc11 = new DataColumn("9", Type.GetType("System.String"));
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
            table.Columns.Add(dc11);

            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "1#";
            dr1["1"] = DBNull.Value;
            dr1["2"] = DBNull.Value;
            dr1["3"] = DBNull.Value;
            dr1["4"] = DBNull.Value;
            dr1["5"] = DBNull.Value;
            dr1["6"] = DBNull.Value;
            dr1["7"] = DBNull.Value;
            dr1["8"] = DBNull.Value;
            dr1["9"] = DBNull.Value;
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["1"] = DBNull.Value;
            dr2["2"] = DBNull.Value;
            dr2["3"] = DBNull.Value;
            dr2["4"] = DBNull.Value;
            dr2["5"] = DBNull.Value;
            dr2["6"] = DBNull.Value;
            dr2["7"] = DBNull.Value;
            dr2["8"] = DBNull.Value;
            dr2["9"] = DBNull.Value;
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["1"] = DBNull.Value;
            dr3["2"] = DBNull.Value;
            dr3["3"] = DBNull.Value;
            dr3["4"] = DBNull.Value;
            dr3["5"] = DBNull.Value;
            dr3["6"] = DBNull.Value;
            dr3["7"] = DBNull.Value;
            dr3["8"] = DBNull.Value;
            dr3["9"] = DBNull.Value;
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1"] = DBNull.Value;
            dr4["2"] = DBNull.Value;
            dr4["3"] = DBNull.Value;
            dr4["4"] = DBNull.Value;
            dr4["5"] = DBNull.Value;
            dr4["6"] = DBNull.Value;
            dr4["7"] = DBNull.Value;
            dr4["8"] = DBNull.Value;
            dr4["9"] = DBNull.Value;
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["1"] = "41.30";
            dr5["2"] = "2.04";
            dr5["3"] = "0.45";
            dr5["4"] = "0.30";
            dr5["5"] = "0.12";
            dr5["6"] = "0.73";
            dr5["7"] = "11.18";
            dr5["8"] = "3.28";
            dr5["9"] = "5.16";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["1"] = "41.09";
            dr6["2"] = "2.47";
            dr6["3"] = "0.47";
            dr6["4"] = "0.49";
            dr6["5"] = "-";
            dr6["6"] = "0.43";
            dr6["7"] = "7.86";
            dr6["8"] = "3.96";
            dr6["9"] = "6.36";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["1"] = "24.11";
            dr7["2"] = "2.37";
            dr7["3"] = "0.25";
            dr7["4"] = "0.53";
            dr7["5"] = "0.70";
            dr7["6"] = "0.87";
            dr7["7"] = "9.79";
            dr7["8"] = "1.19";
            dr7["9"] = "6.29";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["1"] = "24.96";
            dr8["2"] = "2.34";
            dr8["3"] = "0.42";
            dr8["4"] = "0.50";
            dr8["5"] = "-";
            dr8["6"] = "-";
            dr8["7"] = "7.76";
            dr8["8"] = "1.82";
            dr8["9"] = "8.76";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1"] = "24.00";
            dr9["2"] = "1.54";
            dr9["3"] = "-";
            dr9["4"] = "0.43";
            dr9["5"] = "0.66";
            dr9["6"] = "-";
            dr9["7"] = "12.82";
            dr9["8"] = "0.65";
            dr9["9"] = "5.37";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1"] = "24.32";
            dr10["2"] = "1.8";
            dr10["3"] = "-";
            dr10["4"] = "0.51";
            dr10["5"] = "0.23";
            dr10["6"] = "0.55";
            dr10["7"] = "12.76";
            dr10["8"] = "2.49";
            dr10["9"] = "6.78";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1"] = "32.37";
            dr11["2"] = "2.14";
            dr11["3"] = "-";
            dr11["4"] = "0.18";
            dr11["5"] = "-";
            dr11["6"] = "-";
            dr11["7"] = "13.03";
            dr11["8"] = "5.58";
            dr11["9"] = "7.28";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["1"] = "23.97";
            dr12["2"] = "2.08";
            dr12["3"] = "0.52";
            dr12["4"] = "0.49";
            dr12["5"] = "0.73";
            dr12["6"] = "0.85";
            dr12["7"] = "13.67";
            dr12["8"] = "2.36";
            dr12["9"] = "5.84";
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["1"] = "22.98";
            dr13["2"] = "1.97";
            dr13["3"] = "0.57";
            dr13["4"] = "0.47";
            dr13["5"] = "0.54";
            dr13["6"] = "1.11";
            dr13["7"] = "11.26";
            dr13["8"] = "1.84";
            dr13["9"] = "5.55";
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "1#";
            dr14["1"] = "21.72";
            dr14["2"] = "2.33";
            dr14["3"] = "0.34";
            dr14["4"] = "0.52";
            dr14["5"] = "0.72";
            dr14["6"] = "0.45";
            dr14["7"] = "10.36";
            dr14["8"] = "3.10";
            dr14["9"] = "11.78";
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["1"] = "29.38";
            dr15["2"] = "2.24";
            dr15["3"] = "0.30";
            dr15["4"] = "1.02";
            dr15["5"] = "0.38";
            dr15["6"] = "0.87";
            dr15["7"] = "13.01";
            dr15["8"] = "3.15";
            dr15["9"] = "7.07";
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["1"] = "23.55";
            dr16["2"] = "2.22";
            dr16["3"] = "0.31";
            dr16["4"] = "0.56";
            dr16["5"] = "0.75";
            dr16["6"] = "0.76";
            dr16["7"] = "7.33";
            dr16["8"] = "2.40";
            dr16["9"] = "7.34";
            table.Rows.Add(dr16);

            return table;
        }
    }
}

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
            DataColumn dc3 = new DataColumn("A1", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("A2", Type.GetType("System.String"));
            DataColumn dc5 = new DataColumn("A3", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("A4", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("A5", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("A6", Type.GetType("System.String"));
            DataColumn dc9 = new DataColumn("A7", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("A8", Type.GetType("System.String"));
            DataColumn dc11 = new DataColumn("A9", Type.GetType("System.String"));
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
            dr1["A1"] = "clinkerBurning";            //熟料烧成工序
            dr1["A2"] = "rectifierTransformer";           //窑主机
            dr1["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";            //均化库风机 
            dr1["A4"] = "clinkerHoist";            //入窑提升机
            dr1["A5"] = "oneTimeFan01 + oneTimeFan";         //一次风机   
            dr1["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";            //送煤风机
            dr1["A7"] = "highTemperatureFan";            //高温风机
            dr1["A8"] = "kilnHeadExhaustFan";            //头排风机
            dr1["A9"] = "clinkerF5AC + clinkerF8AC + clinkerF2AC + clinkerF6AC + clinkerF7AC + clinkerF10AC";          //篦冷机风机
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker02";
            dr2["A1"] = "clinkerBurning";
            dr2["A2"] = "kilnMainMotor";
            dr2["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr2["A4"] = "clinkerHoist";
            dr2["A5"] = "oneTimeFan01 + oneTimeFan02";
            dr2["A6"] = "rootsBlower1 + rootsBlower2 + rootsBlower3";
            dr2["A7"] = "highTemperatureFan";
            dr2["A8"] = "kilnHeadExhaustFan";
            dr2["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC";        //将来加表
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker03";
            dr3["A1"] = "clinkerBurning";
            dr3["A2"] = "kilnMainMotor";
            dr3["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr3["A4"] = "clinkerHoist";
            dr3["A5"] = "oneTimeFan01 + oneTimeFan02";   
            dr3["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr3["A7"] = "highTemperatureFan";
            dr3["A8"] = "kilnHeadExhaustFan";
            dr3["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC"
                       + " + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC + clinkerF13AC";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker04";
            dr4["A1"] = "clinkerBurning";
            dr4["A2"] = "kilnMainMotor";
            dr4["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr4["A4"] = "clinkerHoist1 + clinkerHoist2";
            dr4["A5"] = "oneTimeFan01 + oneTimeFan02";
            dr4["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr4["A7"] = "highTemperatureFan";
            dr4["A8"] = "kilnHeadExhaustFan";
            dr4["A9"] = "clinkerF1AC + clinkerF3AC + clinkerF4AC + clinkerF2AC + clinkerF5AC + clinkerF6AC + clinkerF7AC"
                       + " + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_efc_clinker02";
            dr5["A1"] = "clinkerBurning";
            dr5["A2"] = "kilnMainMotor";
            dr5["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr5["A4"] = "clinkerHoist";
            dr5["A5"] = "oneTimeFan01 + oneTimeFan02";
            dr5["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr5["A7"] = "highTemperatureFan";
            dr5["A8"] = "kilnHeadExhaustFan";
            dr5["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "zc_nxjc_qtx_efc_clinker03";
            dr6["A1"] = "clinkerBurning";   //熟料烧成工序
            dr6["A2"] = "kilnMainMotor";     //窑主机
            dr6["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";   //均化库风机 
            dr6["A4"] = "clinkerHoist";     //入窑提升机
            dr6["A5"] = "coalPowderInjectionFan1 + coalPowderInjectionFan2";       //一次风机
            dr6["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";     //送煤风机
            dr6["A7"] = "highTemperatureFan";        //高温风机
            dr6["A8"] = "kilnHeadExhaustFan";       //头排风机
            dr6["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";       //篦冷机风机
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "zc_nxjc_qtx_tys_clinker04";
            dr7["A1"] = "clinkerBurning";
            dr7["A2"] = "kilnMainMotor";
            dr7["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr7["A4"] = "clinkerHoist";
            dr7["A5"] = "hootsBlower33020 + hootsBlower33021";
            dr7["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr7["A7"] = "highTemperatureFan";
            dr7["A8"] = "kilnHeadExhaustFan";
            dr7["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "zc_nxjc_qtx_tys_clinker05";
            dr8["A1"] = "clinkerBurning";
            dr8["A2"] = "kilnMainMotor";
            dr8["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr8["A4"] = "clinkerHoist";
            dr8["A5"] = "rootsBlower33016 + oneTimeFan33024 + oneTimeFan33015";
            dr8["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr8["A7"] = "highTemperatureFan";
            dr8["A8"] = "kilnHeadExhaustFan";
            dr8["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "zc_nxjc_znc_znf_clinker01";
            dr9["A1"] = "clinkerBurning";
            dr9["A2"] = "kilnMainMotor";
            dr9["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr9["A4"] = "clinkerHoist";
            dr9["A5"] = "netWindFan";
            dr9["A6"] = "RootsBlowerTail + RootsBlowerBackup + RootsBlowerHead";
            dr9["A7"] = "highTemperatureFan";
            dr9["A8"] = "leftFan";
            dr9["A9"] = "fanN1 + fanN2 + fanN3 + fanN4 + fanN5 + fanN6 + fanN7";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "zc_nxjc_znc_znf_clinker02";
            dr10["A1"] = "clinkerBurning";        //熟料烧成工序
            dr10["A2"] = "kilnMainMotor"; //窑主机
            dr10["A3"] = "rawMaterialLibraryRootsBlower4213 + rawMaterialLibraryRootsBlower4214 + rawMaterialLibraryRootsBlower4215 + rawMaterialLibraryRootsBlower4216"; //均化库风机 
            dr10["A4"] = "clinkerHoist"; //入窑提升机
            dr10["A5"] = "netWindFan";    //一次风机
            dr10["A6"] = "rootsBlower7316 + kilnTailRotorScale";  //送煤风机
            dr10["A7"] = "highTemperatureFan";  //高温风机
            dr10["A8"] = "leftFan";     //头排风机
            dr10["A9"] = "fanV1 + fanV2 + fanV3 + fanV4 + fanV5 + fanV6 + fanV7 + fanV8";    //篦冷机风机
            table.Rows.Add(dr10);

            //DataRow dr11 = table.NewRow();
            //dr11["CompanyName"] = "六盘山水泥";
            //dr11["ProductionLine"] = "zc_nxjc_lpsc_lpsf_clinker01";
            //dr11["A1"] = "clinkerBurning";
            //dr11["A2"] = "kilnMainMotor";
            //dr11["A3"] = DBNull.Value;
            //dr11["A4"] = "clinkerHoist";
            //dr11["A5"] = DBNull.Value;
            //dr11["A6"] = DBNull.Value;
            //dr11["A7"] = "highTemperatureFan";
            //dr11["A8"] = "kilnHeadExhaustFan";
            //dr11["A9"] = "clinkerF2M + clinkerF3M + clinkerF4M + clinkerF5M + clinkerF6M + clinkerF7M + clinker5730 + clinker5732 + coalMillFan";
            //table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker01";
            dr12["A1"] = "clinkerBurning";
            dr12["A2"] = "kilnMainMotor";
            dr12["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr12["A4"] = "kilnHoist";
            dr12["A5"] = "oneTimeFan01 + oneTimeFan02 + oneTimeFan03";
            dr12["A6"] = "coalMilRootsBlower13 + coalMilRootsBlower14 + coalMilRootsBlower15";
            dr12["A7"] = "highTemperatureFan";
            dr12["A8"] = "kilnHeadExhaustFan";
            dr12["A9"] = "grateCoolerFan11 + grateCoolerFan10 + grateCoolerFan07 + grateCoolerFan13 + grateCoolerFan12 + grateCoolerFan06";
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker02";
            dr13["A1"] = "clinkerBurning";
            dr13["A2"] = "kilnMainMotor";
            dr13["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3";
            dr13["A4"] = "kilnHoist";
            dr13["A5"] = "oneTimeFan01 + oneTimeFan02 + oneTimeFan03";
            dr13["A6"] = "coalMilRootsBlower14 + coalMilRootsBlower13 + coalMilRootsBlower15";
            dr13["A7"] = "highTemperatureFan";
            dr13["A8"] = "kilnHeadExhaustFan";
            dr13["A9"] = "grateCoolerFan07 + grateCoolerFan10 + grateCoolerFan13 + grateCoolerFan11 + grateCoolerFan12 + grateCoolerFan08 + grateCoolerFan09 + grateCoolerFan06 + grateCoolerFan05 + grateCoolerFan03 + grateCoolerFan04";
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "zc_nxjc_whsmc_whsmf_clinker01";
            dr14["A1"] = "clinkerBurning";
            dr14["A2"] = "kilnMainMotor";
            dr14["A3"] = "RootsBlower22015 + RootsBlower22016 + RootsBlower22118";
            dr14["A4"] = "hoist22111";
            dr14["A5"] = "oneTimeFan33020 + oneTimeFan33021";
            dr14["A6"] = "RootsBlower43113 + RootsBlower43114 + RootsBlower43115";
            dr14["A7"] = "highTemperatureFan";
            dr14["A8"] = "kilnHeadExhaustFan";
            dr14["A9"] = "coolingFanV3 + coolingFanV1 + coolingFanV2 + coolingFanV6 + coolingFanV7 + coolingFanV4 + coolingFanV5 + coolingFanV8 + coolingFanV10 + coolingFanV11 + coolingFanV9";
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "zc_nxjc_byc_byf_clinker01";
            dr15["A1"] = "clinkerBurning";
            dr15["A2"] = "kilnMainMotor";
            dr15["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3 + rawMaterialLibraryRootsBlower4";
            dr15["A4"] = "clinkerHoist1 + clinkerHoist2";
            dr15["A5"] = "oneTimeFanG27 + oneTimeFanG28";
            dr15["A6"] = "coalMilRootsBlower1 + coalMilRootsBlower2 + coalMilRootsBlower3";
            dr15["A7"] = "highTemperatureFan";
            dr15["A8"] = "kilnHeadExhaustFan";
            dr15["A9"] = "clinkerF1AC + clinkerF2AC + clinkerF3AC + clinkerF4AC + clinkerF5AC + clinkerF6AC + clinkerF7AC + clinkerF8AC + clinkerF9AC + clinkerF10AC + clinkerF11AC + clinkerF12AC + clinkerFVOA + clinkerFVOB + clinkerFVOC";
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "zc_nxjc_klqc_klqf_clinker01";
            dr16["A1"] = "clinkerBurning";
            dr16["A2"] = "kilnMainMotor";
            dr16["A3"] = "rawMaterialLibraryRootsBlower1 + rawMaterialLibraryRootsBlower2 + rawMaterialLibraryRootsBlower3 + rawMaterialLibraryRootsBlower4";
            dr16["A4"] = "kilnHoist + kilnHoist1";
            dr16["A5"] = "rootsBlower5704 + rootsBlower5705";
            dr16["A6"] = "rootsBlower7513 + rootsBlower7514 + rootsBlower7515";
            dr16["A7"] = "highTemperatureFan";
            dr16["A8"] = "kilnHeadExhaustFan";
            dr16["A9"] = "coolingFan5708 + coolingFan5714 + coolingFan5715 + coolingFan5712 + coolingFan5711 + coolingFan5713 + coolingFanF2 + coolingFanF3";
            table.Rows.Add(dr16);

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

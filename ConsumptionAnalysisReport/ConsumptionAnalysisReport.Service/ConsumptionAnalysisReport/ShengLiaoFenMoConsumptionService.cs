﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SqlServerDataAdapter;
using ConsumptionAnalysisReport.Infrastructure.Configuration;
using System.Data.SqlClient;

namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class ShengLiaoFenMoConsumptionService
    {
        public static DataTable GetShengLiaoFenMoDianHao(string mySelectTime)
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

            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            table.Columns.Add(dc5);

            //创建每一行

            DataRow dr1 = table.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "zc_nxjc_ychc_yfcf_clinker01";
            dr1["A1"] = "rawMaterialsGrind";
            dr1["A2"] = "rollingMachineActionRoller + rollingMachineSettledRoller";
            dr1["A3"] = "circulatingFan";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker02";
            dr2["A1"] = "rawMaterialsGrind";
            dr2["A2"] = "rollingMachineActionRoller + rollingMachineSettledRoller";
            dr2["A3"] = "rawMealGrindinSystemFan";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker03";
            dr3["A1"] = "rawMaterialsGrind";
            dr3["A2"] = "rawMealGrindingMainMotor";
            dr3["A3"] = "rawMealGrindinSystemFan";
            table.Rows.Add(dr3);


            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "zc_nxjc_ychc_lsf_clinker04";
            dr4["A1"] = "rawMaterialsGrind";
            dr4["A2"] = "rollingMachineSettledRoller01 + rollingMachineActionRoller01 + rollingMachineSettledRoller02 + rollingMachineActionRoller02";
            dr4["A3"] = "circulatingFan01 + circulatingFan02";
            table.Rows.Add(dr4);


            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "zc_nxjc_qtx_efc_clinker02";
            dr5["A1"] = "rawMaterialsGrind";
            dr5["A2"] = "rawMealGrindingMainMotor + rollingMachineSettledRoller";
            dr5["A3"] = "circulatingFan";
            table.Rows.Add(dr5);


            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "zc_nxjc_qtx_efc_clinker03";
            dr6["A1"] = "rawMaterialsGrind";
            dr6["A2"] = "rawMealGrindingMainMotor";
            dr6["A3"] = "circulatingFan";
            table.Rows.Add(dr6);


            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "zc_nxjc_qtx_tys_clinker04";
            dr7["A1"] = "rawMaterialsGrind";
            dr7["A2"] = "rawMealGrindingMainMotor";
            dr7["A3"] = "circulatingFan";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "zc_nxjc_qtx_tys_clinker05";
            dr8["A1"] = "rawMaterialsGrind";
            dr8["A2"] = "rawMealGrindingMainMotor + rollingMachineActionroller";
            dr8["A3"] = "circulatingFan";
            table.Rows.Add(dr8);


            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "zc_nxjc_znc_znf_clinker01";
            dr9["A1"] = "rawMaterialsGrind";
            dr9["A2"] = "rawMealGrindingMainMotor";
            dr9["A3"] = "rawMealGrindingFan";
            table.Rows.Add(dr9);


            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "zc_nxjc_znc_znf_clinker02";
            dr10["A1"] = "rawMaterialsGrind";
            dr10["A2"] = "rawMealGrindingMainMotor + rollingMachineSettledRoller";
            dr10["A3"] = "rawMealGrindingFan";
            table.Rows.Add(dr10);


            //DataRow dr11 = table.NewRow();
            //dr11["CompanyName"] = "六盘山水泥";
            //dr11["ProductionLine"] = "zc_nxjc_lpsc_lpsf_clinker01";
            //dr11["A1"] = "rawMaterialsGrind";
            //dr11["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            //dr11["A3"] = "circulatingFan";
            //table.Rows.Add(dr11);


            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker01";
            dr12["A1"] = "rawMaterialsGrind";
            dr12["A2"] = "rollingMachineActionRoller + rollingMachineSettledRoller";
            dr12["A3"] = "circulatingFan";
            table.Rows.Add(dr12);


            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "zc_nxjc_tsc_tsf_clinker02";
            dr13["A1"] = "rawMaterialsGrind";
            dr13["A2"] = "rawMealGrindingMainMotor";
            dr13["A3"] = "circulatingFan";
            table.Rows.Add(dr13);


            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "zc_nxjc_whsmc_whsmf_clinker01";
            dr14["A1"] = "rawMaterialsGrind";
            dr14["A2"] = "rawMealGrindingMainMotor";
            dr14["A3"] = "circulatingFan";
            table.Rows.Add(dr14);


            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "zc_nxjc_byc_byf_clinker01";
            dr15["A1"] = "rawMaterialsGrind";
            dr15["A2"] = "rawMealGrindingMainMotor";
            dr15["A3"] = "circulatingFan";
            table.Rows.Add(dr15);


            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "zc_nxjc_klqc_klqf_clinker01";
            dr16["A1"] = "rawMaterialsGrind";
            dr16["A2"] = "rollingMachineSettledRoller + rollingMachineActionRoller";
            dr16["A3"] = "circulatingFan";
            table.Rows.Add(dr16);
            return table;
        }
        
        /// <summary>
        /// 处理公式
        /// </summary>
        /// <param name="preFormula"></param>
        /// <param name="variableId"></param>
        /// <returns></returns>
        private static string DealWithFormula(string preFormula, string variableId)
        {
            if (preFormula.Contains('_'))
            {
                int num = preFormula.IndexOf('_');
                string subStr = preFormula.Substring(1, num - 1);
                return preFormula.Replace(subStr, variableId);
            }
            else
                return preFormula;
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

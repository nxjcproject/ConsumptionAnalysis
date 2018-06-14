using ConsumptionAnalysisReport.Infrastructure.Configuration;
using SqlServerDataAdapter;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class MShuiNiZhiBeiConsumptionService
    {
        public static DataTable GetConsumptionAnalysisTable(string m_SelectFirstTime, string m_SelectSecondTime, string m_selectThirdTime)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string mySql = @"select A.OrganizationID,B.Name,B.LevelCode,B.VariableID,B.LevelType,C.ValueFormula 
                                from (SELECT LevelCode FROM system_Organization WHERE Type = '水泥磨' ) M inner join system_Organization N
	                                on N.LevelCode Like M.LevelCode+'%' inner join tz_Formula A 
	                                on A.OrganizationID=N.OrganizationID inner join formula_FormulaDetail B
	                                on A.KeyID=B.KeyID and A.Type=2 left join balance_Energy_Template C 
	                                on B.VariableId+'_'+@consumptionType=C.VariableId   
                                order by OrganizationID,LevelCode";
            SqlParameter parameter = new SqlParameter("@consumptionType", "ElectricityConsumption");
            DataTable frameTable = dataFactory.Query(mySql, parameter);
            string preFormula = "";
            foreach (DataRow dr in frameTable.Rows)
            {
                if (dr["ValueFormula"] is DBNull)
                {
                    if (dr["VariableId"] is DBNull || dr["VariableId"].ToString().Trim() == "")
                    {
                        continue;
                    }
                    preFormula = DealWithFormula(preFormula, dr["VariableId"].ToString().Trim());
                    dr["ValueFormula"] = preFormula;
                }
                else
                {
                    preFormula = dr["ValueFormula"].ToString().Trim();
                }
            }

            //第一个月份的电耗           
            string dataSql1 = @"select B.OrganizationID,B.VariableId,SUM(B.FirstB) as FirstB,SUM(B.SecondB) as SecondB,SUM(B.ThirdB) as ThirdB,SUM(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB,SUM(B.PeakB) AS PeakB,SUM(B.ValleyB) AS ValleyB,SUM(B.FlatB) AS FlatB,
                                    SUM(CASE WHEN [A].[FirstWorkingTeam] = 'A班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'A班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'A班' THEN [B].[ThirdB] ELSE 0 END) AS teamA,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'B班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'B班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'B班' THEN [B].[ThirdB] ELSE 0 END) AS teamB,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'C班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'C班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'C班' THEN [B].[ThirdB] ELSE 0 END) AS teamC,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'D班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'D班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'D班' THEN [B].[ThirdB] ELSE 0 END) AS teamD
                                from tz_Balance A,balance_Energy B
                                where A.BalanceId=B.KeyId
                                and A.TimeStamp like (@m_SelectFirstTime+'%')                                 
                                group by B.OrganizationID,B.VariableId";
            SqlParameter[] parameters1 = { new SqlParameter("@m_SelectFirstTime", m_SelectFirstTime) };
            DataTable sourceData1 = dataFactory.Query(dataSql1, parameters1);
            string[] calColumns1 = new string[] { "TotalPeakValleyFlatB" };
            DataTable result1 = EnergyConsumption.EnergyConsumptionCalculate.CalculateByOrganizationId(sourceData1, frameTable, "ValueFormula", calColumns1);
            //筛选需要的分厂数据
            DataTable newDT1 = new DataTable();
            newDT1 = result1.Clone();
            DataRow[] dr_Formula1 = result1.Select("VariableId='cementGrind' or VariableId='hybridMaterialsPreparation' or VariableId='clinkerTransport'", "OrganizationID");
            for (int i = 0; i < dr_Formula1.Length; i++)
            {
                newDT1.ImportRow((DataRow)dr_Formula1[i]);
            }
            DataTable newDT11 = new DataTable();
            newDT11 = newDT1.Clone();
            DataRow[] dr_Formula11 = newDT1.Select("OrganizationID='zc_nxjc_qtx_efc_cementmill01' or OrganizationID like 'zc_nxjc_ychc_lsf%' or OrganizationID like 'zc_nxjc_qtx_tys%' or OrganizationID like 'zc_nxjc_tsc_tsf%' or OrganizationID like 'zc_nxjc_whsmc_whsmf%' or OrganizationID like 'zc_nxjc_byc_byf%' or OrganizationID like 'zc_nxjc_klqc_klqf%'", "OrganizationID");
            for (int i = 0; i < dr_Formula11.Length; i++)
            {
                newDT11.ImportRow((DataRow)dr_Formula11[i]);
            }

            //第二个月份的电耗
            string dataSql2 = @"select B.OrganizationID,B.VariableId,SUM(B.FirstB) as FirstB,SUM(B.SecondB) as SecondB,SUM(B.ThirdB) as ThirdB,SUM(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB,SUM(B.PeakB) AS PeakB,SUM(B.ValleyB) AS ValleyB,SUM(B.FlatB) AS FlatB,
                                    SUM(CASE WHEN [A].[FirstWorkingTeam] = 'A班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'A班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'A班' THEN [B].[ThirdB] ELSE 0 END) AS teamA,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'B班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'B班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'B班' THEN [B].[ThirdB] ELSE 0 END) AS teamB,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'C班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'C班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'C班' THEN [B].[ThirdB] ELSE 0 END) AS teamC,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'D班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'D班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'D班' THEN [B].[ThirdB] ELSE 0 END) AS teamD
                                from tz_Balance A,balance_Energy B
                                where A.BalanceId=B.KeyId
                                and A.TimeStamp like (@m_SelectSecondTime+'%')
                                --and B.OrganizationID like 'zc_nxjc_qtx_efc%'                                    
                                group by B.OrganizationID,B.VariableId";
            SqlParameter[] parameters2 = { new SqlParameter("@m_SelectSecondTime", m_SelectSecondTime) };
            DataTable sourceData2 = dataFactory.Query(dataSql2, parameters2);
            string[] calColumns2 = new string[] { "TotalPeakValleyFlatB" };
            DataTable result2 = EnergyConsumption.EnergyConsumptionCalculate.CalculateByOrganizationId(sourceData2, frameTable, "ValueFormula", calColumns2);

            //筛选需要的分厂数据
            DataTable newDT2 = new DataTable();
            newDT2 = result2.Clone();
            DataRow[] dr_Formula2 = result2.Select("VariableId='cementGrind' or VariableId='hybridMaterialsPreparation' or VariableId='clinkerTransport'", "OrganizationID");
            for (int i = 0; i < dr_Formula2.Length; i++)
            {
                newDT2.ImportRow((DataRow)dr_Formula2[i]);
            }
            DataTable newDT22 = new DataTable();
            newDT22 = newDT2.Clone();
            DataRow[] dr_Formula22 = newDT2.Select("OrganizationID='zc_nxjc_qtx_efc_cementmill01' or OrganizationID like 'zc_nxjc_ychc_lsf%' or OrganizationID like 'zc_nxjc_qtx_tys%' or OrganizationID like 'zc_nxjc_tsc_tsf%' or OrganizationID like 'zc_nxjc_whsmc_whsmf%' or OrganizationID like 'zc_nxjc_byc_byf%' or OrganizationID like 'zc_nxjc_klqc_klqf%'", "OrganizationID");
            for (int i = 0; i < dr_Formula22.Length; i++)
            {
                newDT22.ImportRow((DataRow)dr_Formula22[i]);
            }
            //第三个月份的电耗
            string dataSql3 = @"select B.OrganizationID,B.VariableId,SUM(B.FirstB) as FirstB,SUM(B.SecondB) as SecondB,SUM(B.ThirdB) as ThirdB,SUM(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB,SUM(B.PeakB) AS PeakB,SUM(B.ValleyB) AS ValleyB,SUM(B.FlatB) AS FlatB,
                                    SUM(CASE WHEN [A].[FirstWorkingTeam] = 'A班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'A班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'A班' THEN [B].[ThirdB] ELSE 0 END) AS teamA,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'B班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'B班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'B班' THEN [B].[ThirdB] ELSE 0 END) AS teamB,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'C班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'C班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'C班' THEN [B].[ThirdB] ELSE 0 END) AS teamC,
		                            SUM(CASE WHEN [A].[FirstWorkingTeam] = 'D班' THEN [B].[FirstB] WHEN [A].[SecondWorkingTeam] = 'D班' THEN [B].[SecondB] WHEN [A].[ThirdWorkingTeam] = 'D班' THEN [B].[ThirdB] ELSE 0 END) AS teamD
                                from tz_Balance A,balance_Energy B
                                where A.BalanceId=B.KeyId
                                and A.TimeStamp like (@m_selectThirdTime+'%')
                                --and B.OrganizationID like 'zc_nxjc_qtx_efc%'                                    
                                group by B.OrganizationID,B.VariableId";
            SqlParameter[] parameters3 = { new SqlParameter("@m_selectThirdTime", m_selectThirdTime) };
            DataTable sourceData3 = dataFactory.Query(dataSql3, parameters3);
            string[] calColumns3 = new string[] { "TotalPeakValleyFlatB" };
            DataTable result3 = EnergyConsumption.EnergyConsumptionCalculate.CalculateByOrganizationId(sourceData3, frameTable, "ValueFormula", calColumns3);

            //筛选需要的分厂数据
            DataTable newDT3 = new DataTable();
            newDT3 = result3.Clone();
            DataRow[] dr_Formula3 = result3.Select("VariableId='cementGrind' or VariableId='hybridMaterialsPreparation' or VariableId='clinkerTransport'", "OrganizationID");
            for (int i = 0; i < dr_Formula3.Length; i++)
            {
                newDT3.ImportRow((DataRow)dr_Formula3[i]);
            }
            DataTable newDT33 = new DataTable();
            newDT33 = newDT3.Clone();
            DataRow[] dr_Formula33 = newDT3.Select("OrganizationID='zc_nxjc_qtx_efc_cementmill01' or OrganizationID like 'zc_nxjc_ychc_lsf%' or OrganizationID like 'zc_nxjc_qtx_tys%' or OrganizationID like 'zc_nxjc_tsc_tsf%' or OrganizationID like 'zc_nxjc_whsmc_whsmf%' or OrganizationID like 'zc_nxjc_byc_byf%' or OrganizationID like 'zc_nxjc_klqc_klqf%'", "OrganizationID");
            for (int i = 0; i < dr_Formula33.Length; i++)
            {
                newDT33.ImportRow((DataRow)dr_Formula33[i]);
            }

            DataTable ResutTable = BuildEmptyTable();//最终返回的Table
            //将以上三个月分的数据填充到ResutTable
            DataRow dr1 = ResutTable.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "8#";
            dr1["A1"] = Convert.ToDecimal(newDT11.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A2"] = Convert.ToDecimal(newDT22.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A3"] = Convert.ToDecimal(newDT33.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A4"] = Convert.ToDecimal(newDT11.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A5"] = Convert.ToDecimal(newDT22.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A6"] = Convert.ToDecimal(newDT33.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A7"] = Convert.ToDecimal(newDT11.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A8"] = Convert.ToDecimal(newDT22.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["A9"] = Convert.ToDecimal(newDT33.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr1);

            DataRow dr2 = ResutTable.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "9#";
            dr2["A1"] = Convert.ToDecimal(newDT11.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A2"] = Convert.ToDecimal(newDT22.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A3"] = Convert.ToDecimal(newDT33.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A4"] = Convert.ToDecimal(newDT11.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A5"] = Convert.ToDecimal(newDT22.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A6"] = Convert.ToDecimal(newDT33.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A7"] = Convert.ToDecimal(newDT11.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A8"] = Convert.ToDecimal(newDT22.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["A9"] = Convert.ToDecimal(newDT33.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr2);

            DataRow dr3 = ResutTable.NewRow();
            dr3["CompanyName"] = "青铜峡水泥";
            dr3["ProductionLine"] = "1#";
            dr3["A1"] = Convert.ToDecimal(newDT11.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A2"] = Convert.ToDecimal(newDT22.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A3"] = Convert.ToDecimal(newDT33.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A4"] = Convert.ToDecimal(newDT11.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A5"] = Convert.ToDecimal(newDT22.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A6"] = Convert.ToDecimal(newDT33.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A7"] = Convert.ToDecimal(newDT11.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A8"] = Convert.ToDecimal(newDT22.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["A9"] = Convert.ToDecimal(newDT33.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr3);

            DataRow dr4 = ResutTable.NewRow();
            dr4["CompanyName"] = "青铜峡水泥";
            dr4["ProductionLine"] = "4#";
            dr4["A1"] = Convert.ToDecimal(newDT11.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A2"] = Convert.ToDecimal(newDT22.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A3"] = Convert.ToDecimal(newDT33.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A4"] = Convert.ToDecimal(newDT11.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A5"] = Convert.ToDecimal(newDT22.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A6"] = Convert.ToDecimal(newDT33.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A7"] = Convert.ToDecimal(newDT11.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A8"] = Convert.ToDecimal(newDT22.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["A9"] = Convert.ToDecimal(newDT33.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr4);

            DataRow dr5 = ResutTable.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "5#";
            dr5["A1"] = Convert.ToDecimal(newDT11.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A2"] = Convert.ToDecimal(newDT22.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A3"] = Convert.ToDecimal(newDT33.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A4"] = Convert.ToDecimal(newDT11.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A5"] = Convert.ToDecimal(newDT22.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A6"] = Convert.ToDecimal(newDT33.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A7"] = Convert.ToDecimal(newDT11.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A8"] = Convert.ToDecimal(newDT22.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["A9"] = Convert.ToDecimal(newDT33.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr5);

            DataRow dr6 = ResutTable.NewRow();
            dr6["CompanyName"] = "天水水泥";
            dr6["ProductionLine"] = "1#";
            dr6["A1"] = Convert.ToDecimal(newDT11.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A2"] = Convert.ToDecimal(newDT22.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A3"] = Convert.ToDecimal(newDT33.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A4"] = Convert.ToDecimal(newDT11.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A5"] = Convert.ToDecimal(newDT22.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A6"] = Convert.ToDecimal(newDT33.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A7"] = Convert.ToDecimal(newDT11.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A8"] = Convert.ToDecimal(newDT22.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["A9"] = Convert.ToDecimal(newDT33.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr6);

            DataRow dr7 = ResutTable.NewRow();
            dr7["CompanyName"] = "天水水泥";
            dr7["ProductionLine"] = "2#";
            dr7["A1"] = Convert.ToDecimal(newDT11.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A2"] = Convert.ToDecimal(newDT22.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A3"] = Convert.ToDecimal(newDT33.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A4"] = Convert.ToDecimal(newDT11.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A5"] = Convert.ToDecimal(newDT22.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A6"] = Convert.ToDecimal(newDT33.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A7"] = Convert.ToDecimal(newDT11.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A8"] = Convert.ToDecimal(newDT22.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["A9"] = Convert.ToDecimal(newDT33.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr7);

            DataRow dr8 = ResutTable.NewRow();
            dr8["CompanyName"] = "乌海赛马";
            dr8["ProductionLine"] = "1#";
            dr8["A1"] = Convert.ToDecimal(newDT11.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A2"] = Convert.ToDecimal(newDT22.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A3"] = Convert.ToDecimal(newDT33.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A4"] = Convert.ToDecimal(newDT11.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A5"] = Convert.ToDecimal(newDT22.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A6"] = Convert.ToDecimal(newDT33.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A7"] = Convert.ToDecimal(newDT11.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A8"] = Convert.ToDecimal(newDT22.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["A9"] = Convert.ToDecimal(newDT33.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr8);

            DataRow dr9 = ResutTable.NewRow();
            dr9["CompanyName"] = "白银水泥";
            dr9["ProductionLine"] = "1#";
            dr9["A1"] = Convert.ToDecimal(newDT11.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");//熟料输送
            dr9["A2"] = Convert.ToDecimal(newDT22.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["A3"] = Convert.ToDecimal(newDT33.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["A4"] = Convert.ToDecimal(newDT11.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["A5"] = Convert.ToDecimal(newDT22.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");//混合材
            dr9["A6"] = Convert.ToDecimal(newDT33.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["A7"] = Convert.ToDecimal(newDT11.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");//水泥粉磨
            dr9["A8"] = Convert.ToDecimal(newDT22.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["A9"] = Convert.ToDecimal(newDT33.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr9);

            DataRow dr10 = ResutTable.NewRow();
            dr10["CompanyName"] = "白银水泥";
            dr10["ProductionLine"] = "2#";
            dr10["A1"] = Convert.ToDecimal(newDT11.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A2"] = Convert.ToDecimal(newDT22.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A3"] = Convert.ToDecimal(newDT33.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A4"] = Convert.ToDecimal(newDT11.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A5"] = Convert.ToDecimal(newDT22.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A6"] = Convert.ToDecimal(newDT33.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A7"] = Convert.ToDecimal(newDT11.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A8"] = Convert.ToDecimal(newDT22.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["A9"] = Convert.ToDecimal(newDT33.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr10);

            DataRow dr11 = ResutTable.NewRow();
            dr11["CompanyName"] = "喀喇沁水泥";
            dr11["ProductionLine"] = "1#";
            dr11["A1"] = Convert.ToDecimal(newDT11.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A2"] = Convert.ToDecimal(newDT22.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A3"] = Convert.ToDecimal(newDT33.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A4"] = Convert.ToDecimal(newDT11.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A5"] = Convert.ToDecimal(newDT22.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A6"] = Convert.ToDecimal(newDT33.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A7"] = Convert.ToDecimal(newDT11.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A8"] = Convert.ToDecimal(newDT22.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["A9"] = Convert.ToDecimal(newDT33.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr11);

            DataRow dr12 = ResutTable.NewRow();
            dr12["CompanyName"] = "喀喇沁水泥";
            dr12["ProductionLine"] = "2#";
            dr12["A1"] = Convert.ToDecimal(newDT11.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A2"] = Convert.ToDecimal(newDT22.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A3"] = Convert.ToDecimal(newDT33.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A4"] = Convert.ToDecimal(newDT11.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A5"] = Convert.ToDecimal(newDT22.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A6"] = Convert.ToDecimal(newDT33.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A7"] = Convert.ToDecimal(newDT11.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A8"] = Convert.ToDecimal(newDT22.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["A9"] = Convert.ToDecimal(newDT33.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr12);

            return ResutTable;
        }

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

        public static DataTable BuildEmptyTable()
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

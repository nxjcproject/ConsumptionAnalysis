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
            dr1["1"] = Convert.ToDecimal(newDT11.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["2"] = Convert.ToDecimal(newDT22.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["3"] = Convert.ToDecimal(newDT33.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["4"] = Convert.ToDecimal(newDT11.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["5"] = Convert.ToDecimal(newDT22.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["6"] = Convert.ToDecimal(newDT33.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["7"] = Convert.ToDecimal(newDT11.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["8"] = Convert.ToDecimal(newDT22.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["9"] = Convert.ToDecimal(newDT33.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr1);

            DataRow dr2 = ResutTable.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "9#";
            dr2["1"] = Convert.ToDecimal(newDT11.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["2"] = Convert.ToDecimal(newDT22.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["3"] = Convert.ToDecimal(newDT33.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["4"] = Convert.ToDecimal(newDT11.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["5"] = Convert.ToDecimal(newDT22.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["6"] = Convert.ToDecimal(newDT33.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["7"] = Convert.ToDecimal(newDT11.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["8"] = Convert.ToDecimal(newDT22.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["9"] = Convert.ToDecimal(newDT33.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr2);

            DataRow dr3 = ResutTable.NewRow();
            dr3["CompanyName"] = "青铜峡水泥";
            dr3["ProductionLine"] = "1#";
            dr3["1"] = Convert.ToDecimal(newDT11.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["2"] = Convert.ToDecimal(newDT22.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["3"] = Convert.ToDecimal(newDT33.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["4"] = Convert.ToDecimal(newDT11.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["5"] = Convert.ToDecimal(newDT22.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["6"] = Convert.ToDecimal(newDT33.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["7"] = Convert.ToDecimal(newDT11.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["8"] = Convert.ToDecimal(newDT22.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["9"] = Convert.ToDecimal(newDT33.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr3);

            DataRow dr4 = ResutTable.NewRow();
            dr4["CompanyName"] = "青铜峡水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1"] = Convert.ToDecimal(newDT11.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["2"] = Convert.ToDecimal(newDT22.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["3"] = Convert.ToDecimal(newDT33.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["4"] = Convert.ToDecimal(newDT11.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["5"] = Convert.ToDecimal(newDT22.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["6"] = Convert.ToDecimal(newDT33.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["7"] = Convert.ToDecimal(newDT11.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["8"] = Convert.ToDecimal(newDT22.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["9"] = Convert.ToDecimal(newDT33.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr4);

            DataRow dr5 = ResutTable.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "5#";
            dr5["1"] = Convert.ToDecimal(newDT11.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["2"] = Convert.ToDecimal(newDT22.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["3"] = Convert.ToDecimal(newDT33.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["4"] = Convert.ToDecimal(newDT11.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["5"] = Convert.ToDecimal(newDT22.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["6"] = Convert.ToDecimal(newDT33.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["7"] = Convert.ToDecimal(newDT11.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["8"] = Convert.ToDecimal(newDT22.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["9"] = Convert.ToDecimal(newDT33.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr5);

            DataRow dr6 = ResutTable.NewRow();
            dr6["CompanyName"] = "天水水泥";
            dr6["ProductionLine"] = "1#";
            dr6["1"] = Convert.ToDecimal(newDT11.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["2"] = Convert.ToDecimal(newDT22.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["3"] = Convert.ToDecimal(newDT33.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["4"] = Convert.ToDecimal(newDT11.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["5"] = Convert.ToDecimal(newDT22.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["6"] = Convert.ToDecimal(newDT33.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["7"] = Convert.ToDecimal(newDT11.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["8"] = Convert.ToDecimal(newDT22.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["9"] = Convert.ToDecimal(newDT33.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr6);

            DataRow dr7 = ResutTable.NewRow();
            dr7["CompanyName"] = "天水水泥";
            dr7["ProductionLine"] = "2#";
            dr7["1"] = Convert.ToDecimal(newDT11.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["2"] = Convert.ToDecimal(newDT22.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["3"] = Convert.ToDecimal(newDT33.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["4"] = Convert.ToDecimal(newDT11.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["5"] = Convert.ToDecimal(newDT22.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["6"] = Convert.ToDecimal(newDT33.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["7"] = Convert.ToDecimal(newDT11.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["8"] = Convert.ToDecimal(newDT22.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["9"] = Convert.ToDecimal(newDT33.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr7);

            DataRow dr8 = ResutTable.NewRow();
            dr8["CompanyName"] = "乌海赛马";
            dr8["ProductionLine"] = "1#";
            dr8["1"] = Convert.ToDecimal(newDT11.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["2"] = Convert.ToDecimal(newDT22.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["3"] = Convert.ToDecimal(newDT33.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["4"] = Convert.ToDecimal(newDT11.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["5"] = Convert.ToDecimal(newDT22.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["6"] = Convert.ToDecimal(newDT33.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["7"] = Convert.ToDecimal(newDT11.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["8"] = Convert.ToDecimal(newDT22.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["9"] = Convert.ToDecimal(newDT33.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr8);

            DataRow dr9 = ResutTable.NewRow();
            dr9["CompanyName"] = "白银水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1"] = Convert.ToDecimal(newDT11.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");//熟料输送
            dr9["2"] = Convert.ToDecimal(newDT22.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["3"] = Convert.ToDecimal(newDT33.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["4"] = Convert.ToDecimal(newDT11.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["5"] = Convert.ToDecimal(newDT22.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");//混合材
            dr9["6"] = Convert.ToDecimal(newDT33.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["7"] = Convert.ToDecimal(newDT11.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");//水泥粉磨
            dr9["8"] = Convert.ToDecimal(newDT22.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["9"] = Convert.ToDecimal(newDT33.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr9);

            DataRow dr10 = ResutTable.NewRow();
            dr10["CompanyName"] = "白银水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1"] = Convert.ToDecimal(newDT11.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["2"] = Convert.ToDecimal(newDT22.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["3"] = Convert.ToDecimal(newDT33.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["4"] = Convert.ToDecimal(newDT11.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["5"] = Convert.ToDecimal(newDT22.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["6"] = Convert.ToDecimal(newDT33.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["7"] = Convert.ToDecimal(newDT11.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["8"] = Convert.ToDecimal(newDT22.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["9"] = Convert.ToDecimal(newDT33.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr10);

            DataRow dr11 = ResutTable.NewRow();
            dr11["CompanyName"] = "喀喇沁水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1"] = Convert.ToDecimal(newDT11.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["2"] = Convert.ToDecimal(newDT22.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["3"] = Convert.ToDecimal(newDT33.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["4"] = Convert.ToDecimal(newDT11.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["5"] = Convert.ToDecimal(newDT22.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["6"] = Convert.ToDecimal(newDT33.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["7"] = Convert.ToDecimal(newDT11.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["8"] = Convert.ToDecimal(newDT22.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["9"] = Convert.ToDecimal(newDT33.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr11);

            DataRow dr12 = ResutTable.NewRow();
            dr12["CompanyName"] = "喀喇沁水泥";
            dr12["ProductionLine"] = "2#";
            dr12["1"] = Convert.ToDecimal(newDT11.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["2"] = Convert.ToDecimal(newDT22.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["3"] = Convert.ToDecimal(newDT33.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["4"] = Convert.ToDecimal(newDT11.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["5"] = Convert.ToDecimal(newDT22.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["6"] = Convert.ToDecimal(newDT33.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["7"] = Convert.ToDecimal(newDT11.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["8"] = Convert.ToDecimal(newDT22.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["9"] = Convert.ToDecimal(newDT33.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
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
            return table;
        }

        public static DataTable GetFixConsumptionAnalysisTable()
        {
            //建表
            DataTable table = new DataTable();
            //建列
            DataColumn dc1 = new DataColumn("CompanyName", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("ProductionLine", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("1", Type.GetType("System.Double"));
            DataColumn dc4 = new DataColumn("2", Type.GetType("System.Double"));
            DataColumn dc5 = new DataColumn("3", Type.GetType("System.Double"));
            DataColumn dc6 = new DataColumn("4", Type.GetType("System.Double"));
            DataColumn dc7 = new DataColumn("5", Type.GetType("System.Double"));
            DataColumn dc8 = new DataColumn("6", Type.GetType("System.Double"));
            DataColumn dc9 = new DataColumn("7", Type.GetType("System.Double"));
            DataColumn dc10 = new DataColumn("8", Type.GetType("System.Double"));
            DataColumn dc11 = new DataColumn("9", Type.GetType("System.Double"));
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
            dr1["ProductionLine"] = "8#";
            dr1["1"] = 0.03;
            dr1["2"] = 0.04;
            dr1["3"] = DBNull.Value;
            dr1["4"] = DBNull.Value;
            dr1["5"] = DBNull.Value;
            dr1["6"] = DBNull.Value;
            dr1["7"] = 29.89;
            dr1["8"] = 26.97;
            dr1["9"] = DBNull.Value;
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "9#";
            dr2["1"] = 0.03;
            dr2["2"] = 0.04;
            dr2["3"] = DBNull.Value;
            dr2["4"] = DBNull.Value;
            dr2["5"] = DBNull.Value;
            dr2["6"] = DBNull.Value;
            dr2["7"] = 26.79;
            dr2["8"] = 26.4;
            dr2["9"] = DBNull.Value;
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "青铜峡水泥";
            dr3["ProductionLine"] = "1#";
            dr3["1"] = 0.19;
            dr3["2"] = 0.19;
            dr3["3"] = 0.71;
            dr3["4"] = 0.17;
            dr3["5"] = 0.16;
            dr3["6"] = 0.28;
            dr3["7"] = 38.79;
            dr3["8"] = 37.99;
            dr3["9"] = 57.83;
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "青铜峡水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1"] = 0.43;
            dr4["2"] = 0.44;
            dr4["3"] = 0.74;
            dr4["4"] = 0.08;
            dr4["5"] = 0.07;
            dr4["6"] = 0.04;
            dr4["7"] = 33.93;
            dr4["8"] = 33.33;
            dr4["9"] = 36.00;
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "5#";
            dr5["1"] = 0.45;
            dr5["2"] = 0.48;
            dr5["3"] = 0.74;
            dr5["4"] = 0.08;
            dr5["5"] = 0.10;
            dr5["6"] = 0.04;
            dr5["7"] = 35.36;
            dr5["8"] = 35.85;
            dr5["9"] = 37.37;
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "天水水泥";
            dr6["ProductionLine"] = "1#";
            dr6["1"] = 0.34;
            dr6["2"] = 0.31;
            dr6["3"] = 0.33;
            dr6["4"] = 0.16;
            dr6["5"] = 0.17;
            dr6["6"] = 0.19;
            dr6["7"] = 38.99;
            dr6["8"] = 41.23;
            dr6["9"] = 41.37;
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "天水水泥";
            dr7["ProductionLine"] = "2#";
            dr7["1"] = 0.31;
            dr7["2"] = 0.29;
            dr7["3"] = 0.26;
            dr7["4"] = 0.16;
            dr7["5"] = 0.17;
            dr7["6"] = 0.17;
            dr7["7"] = 27.19;
            dr7["8"] = 28.26;
            dr7["9"] = 27.87;
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "乌海赛马";
            dr8["ProductionLine"] = "1#";
            dr8["1"] = 0.2;
            dr8["2"] = 0.19;
            dr8["3"] = 0.24;
            dr8["4"] = 0.11;
            dr8["5"] = 0.09;
            dr8["6"] = 0.10;
            dr8["7"] = 31.43;
            dr8["8"] = 28;
            dr8["9"] = 33.23;
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "白银水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1"] = 0.15;
            dr9["2"] = 0.16;
            dr9["3"] = 0.19;
            dr9["4"] = 0.26;
            dr9["5"] = 0.26;
            dr9["6"] = 0.34;
            dr9["7"] = 35.01;
            dr9["8"] = 36.74;
            dr9["9"] = 35.40;
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "白银水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1"] = 0.15;
            dr10["2"] = 0.16;
            dr10["3"] = 0.17;
            dr10["4"] = 0.27;
            dr10["5"] = 0.26;
            dr10["6"] = 0.30;
            dr10["7"] = 34.92;
            dr10["8"] = 35.24;
            dr10["9"] = 34.63;
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "喀喇沁水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1"] = 0.21;
            dr11["2"] = 0.25;
            dr11["3"] = 0.33;
            dr11["4"] = 0.25;
            dr11["5"] = 0.25;
            dr11["6"] = 0.19;
            dr11["7"] = 36.3;
            dr11["8"] = 31.22;
            dr11["9"] = 35.25;
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "喀喇沁水泥";
            dr12["ProductionLine"] = "2#";
            dr12["1"] = 0.22;
            dr12["2"] = 0.23;
            dr12["3"] = DBNull.Value;
            dr12["4"] = 0.25;
            dr12["5"] = 0.22;
            dr12["6"] = DBNull.Value;
            dr12["7"] = 36.15;
            dr12["8"] = 30.31;
            dr12["9"] = DBNull.Value;
            table.Rows.Add(dr12);

            return table;
        }
    }
}

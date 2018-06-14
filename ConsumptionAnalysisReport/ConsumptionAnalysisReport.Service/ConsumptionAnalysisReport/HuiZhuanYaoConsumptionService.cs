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
    public class HuiZhuanYaoConsumptionService
    {
        public static DataTable GetHZYConsumptionTable(string m_SelectFirstTime, string m_SelectSecondTime, string m_selectThirdTime)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string mySql = @"select A.OrganizationID,B.Name,B.LevelCode,B.VariableID,B.LevelType,C.ValueFormula 
                                from (SELECT LevelCode FROM system_Organization WHERE Type = '熟料' ) M inner join system_Organization N
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
                                --and B.OrganizationID like 'zc_nxjc_qtx_efc%'                                    
                                group by B.OrganizationID,B.VariableId";
            SqlParameter[] parameters1 = { new SqlParameter("@m_SelectFirstTime", m_SelectFirstTime) };
            DataTable sourceData1 = dataFactory.Query(dataSql1, parameters1);
            string[] calColumns1 = new string[] { "TotalPeakValleyFlatB" };

            DataTable result1 = EnergyConsumption.EnergyConsumptionCalculate.CalculateByOrganizationId(sourceData1, frameTable, "ValueFormula", calColumns1);
            DataTable newDT1 = new DataTable();
            newDT1 = result1.Clone();
            DataRow[] dr_Formula1 = result1.Select("VariableId='rawMaterialsHomogenize' or VariableId='rawMaterialsGrind' or VariableId='coalPreparation' or VariableId='clinkerBurning' or VariableId='kilnSystem'", "OrganizationID");
            for (int i = 0; i < dr_Formula1.Length; i++)
            {
                newDT1.ImportRow((DataRow)dr_Formula1[i]);
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
            DataTable newDT2 = new DataTable();
            newDT2 = result2.Clone();
            DataRow[] dr_Formula2 = result2.Select("VariableId='rawMaterialsHomogenize' or VariableId='rawMaterialsGrind' or VariableId='coalPreparation' or VariableId='clinkerBurning' or VariableId='kilnSystem'", "OrganizationID");
            for (int i = 0; i < dr_Formula2.Length; i++)
            {
                newDT2.ImportRow((DataRow)dr_Formula2[i]);
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
            DataTable newDT3 = new DataTable();
            newDT3 = result3.Clone();
            DataRow[] dr_Formula3 = result3.Select("VariableId='rawMaterialsHomogenize' or VariableId='rawMaterialsGrind' or VariableId='coalPreparation' or VariableId='clinkerBurning' or VariableId='kilnSystem'", "OrganizationID");
            for (int i = 0; i < dr_Formula3.Length; i++)
            {
                newDT3.ImportRow((DataRow)dr_Formula3[i]);
            }

            DataTable ResutTable = BuildEmptyTable();//最终返回的Table
            //将以上三个月分的数据填充到ResutTable
            DataRow dr1 = ResutTable.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "1#";
            dr1["1month9"] = Convert.ToDecimal(newDT1.Rows[69]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["1month10"] = Convert.ToDecimal(newDT2.Rows[69]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["1month11"] = Convert.ToDecimal(newDT3.Rows[69]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["2month9"] = Convert.ToDecimal(newDT1.Rows[68]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["2month10"] = Convert.ToDecimal(newDT2.Rows[68]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr1["2month11"] = Convert.ToDecimal(newDT3.Rows[68]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["3month9"] = Convert.ToDecimal(newDT1.Rows[66]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["3month10"] = Convert.ToDecimal(newDT2.Rows[66]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["3month11"] = Convert.ToDecimal(newDT3.Rows[66]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["4month9"] = Convert.ToDecimal(newDT1.Rows[65]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["4month10"] = Convert.ToDecimal(newDT2.Rows[65]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["4month11"] = Convert.ToDecimal(newDT3.Rows[65]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["5month9"] = Convert.ToDecimal(newDT1.Rows[67]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["5month10"] = Convert.ToDecimal(newDT2.Rows[67]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr1["5month11"] = Convert.ToDecimal(newDT3.Rows[67]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr1);

            DataRow dr2 = ResutTable.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["1month9"] = Convert.ToDecimal(newDT1.Rows[54]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["1month10"] = Convert.ToDecimal(newDT2.Rows[54]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr2["1month11"] = Convert.ToDecimal(newDT3.Rows[54]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["2month9"] = Convert.ToDecimal(newDT1.Rows[53]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["2month10"] = Convert.ToDecimal(newDT2.Rows[53]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["2month11"] = Convert.ToDecimal(newDT3.Rows[53]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["3month9"] = Convert.ToDecimal(newDT1.Rows[51]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["3month10"] = Convert.ToDecimal(newDT2.Rows[51]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["3month11"] = Convert.ToDecimal(newDT3.Rows[51]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["4month9"] = Convert.ToDecimal(newDT1.Rows[50]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["4month10"] = Convert.ToDecimal(newDT2.Rows[50]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["4month11"] = Convert.ToDecimal(newDT3.Rows[50]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["5month9"] = Convert.ToDecimal(newDT1.Rows[52]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["5month10"] = Convert.ToDecimal(newDT2.Rows[52]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr2["5month11"] = Convert.ToDecimal(newDT3.Rows[52]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr2);

            DataRow dr3 = ResutTable.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["1month9"] = Convert.ToDecimal(newDT1.Rows[59]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["1month10"] = Convert.ToDecimal(newDT2.Rows[59]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["1month11"] = Convert.ToDecimal(newDT3.Rows[59]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["2month9"] = Convert.ToDecimal(newDT1.Rows[58]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["2month10"] = Convert.ToDecimal(newDT2.Rows[58]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["2month11"] = Convert.ToDecimal(newDT3.Rows[58]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["3month9"] = Convert.ToDecimal(newDT1.Rows[56]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr3["3month10"] = Convert.ToDecimal(newDT2.Rows[56]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["3month11"] = Convert.ToDecimal(newDT3.Rows[56]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["4month9"] = Convert.ToDecimal(newDT1.Rows[55]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["4month10"] = Convert.ToDecimal(newDT2.Rows[55]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["4month11"] = Convert.ToDecimal(newDT3.Rows[55]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["5month9"] = Convert.ToDecimal(newDT1.Rows[57]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["5month10"] = Convert.ToDecimal(newDT2.Rows[57]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr3["5month11"] = Convert.ToDecimal(newDT3.Rows[57]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr3);

            DataRow dr4 = ResutTable.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1month9"] = Convert.ToDecimal(newDT1.Rows[64]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["1month10"] = Convert.ToDecimal(newDT2.Rows[64]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["1month11"] = Convert.ToDecimal(newDT3.Rows[64]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr4["2month9"] = Convert.ToDecimal(newDT1.Rows[63]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["2month10"] = Convert.ToDecimal(newDT2.Rows[63]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["2month11"] = Convert.ToDecimal(newDT3.Rows[63]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["3month9"] = Convert.ToDecimal(newDT1.Rows[61]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["3month10"] = Convert.ToDecimal(newDT2.Rows[61]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["3month11"] = Convert.ToDecimal(newDT3.Rows[61]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["4month9"] = Convert.ToDecimal(newDT1.Rows[60]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["4month10"] = Convert.ToDecimal(newDT2.Rows[60]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["4month11"] = Convert.ToDecimal(newDT3.Rows[60]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["5month9"] = Convert.ToDecimal(newDT1.Rows[62]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["5month10"] = Convert.ToDecimal(newDT2.Rows[62]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr4["5month11"] = Convert.ToDecimal(newDT3.Rows[62]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr4);

            DataRow dr5 = ResutTable.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["1month9"] = Convert.ToDecimal(newDT1.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["1month10"] = Convert.ToDecimal(newDT2.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["1month11"] = Convert.ToDecimal(newDT3.Rows[19]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["2month9"] = Convert.ToDecimal(newDT1.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr5["2month10"] = Convert.ToDecimal(newDT2.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["2month11"] = Convert.ToDecimal(newDT3.Rows[18]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["3month9"] = Convert.ToDecimal(newDT1.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["3month10"] = Convert.ToDecimal(newDT2.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["3month11"] = Convert.ToDecimal(newDT3.Rows[16]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["4month9"] = Convert.ToDecimal(newDT1.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["4month10"] = Convert.ToDecimal(newDT2.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["4month11"] = Convert.ToDecimal(newDT3.Rows[15]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["5month9"] = Convert.ToDecimal(newDT1.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["5month10"] = Convert.ToDecimal(newDT2.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr5["5month11"] = Convert.ToDecimal(newDT3.Rows[17]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr5);

            DataRow dr6 = ResutTable.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["1month9"] = Convert.ToDecimal(newDT1.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["1month10"] = Convert.ToDecimal(newDT2.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["1month11"] = Convert.ToDecimal(newDT3.Rows[24]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["2month9"] = Convert.ToDecimal(newDT1.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["2month10"] = Convert.ToDecimal(newDT2.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["2month11"] = Convert.ToDecimal(newDT3.Rows[23]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["3month9"] = Convert.ToDecimal(newDT1.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["3month10"] = Convert.ToDecimal(newDT2.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr6["3month11"] = Convert.ToDecimal(newDT3.Rows[21]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["4month9"] = Convert.ToDecimal(newDT1.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["4month10"] = Convert.ToDecimal(newDT2.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["4month11"] = Convert.ToDecimal(newDT3.Rows[20]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["5month9"] = Convert.ToDecimal(newDT1.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["5month10"] = Convert.ToDecimal(newDT2.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr6["5month11"] = Convert.ToDecimal(newDT3.Rows[22]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr6);

            DataRow dr7 = ResutTable.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["1month9"] = Convert.ToDecimal(newDT1.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["1month10"] = Convert.ToDecimal(newDT2.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["1month11"] = Convert.ToDecimal(newDT3.Rows[29]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["2month9"] = Convert.ToDecimal(newDT1.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["2month10"] = Convert.ToDecimal(newDT2.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["2month11"] = Convert.ToDecimal(newDT3.Rows[28]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["3month9"] = Convert.ToDecimal(newDT1.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["3month10"] = Convert.ToDecimal(newDT2.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr7["3month11"] = Convert.ToDecimal(newDT3.Rows[26]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["4month9"] = Convert.ToDecimal(newDT1.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["4month10"] = Convert.ToDecimal(newDT2.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["4month11"] = Convert.ToDecimal(newDT3.Rows[25]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["5month9"] = Convert.ToDecimal(newDT1.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["5month10"] = Convert.ToDecimal(newDT2.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr7["5month11"] = Convert.ToDecimal(newDT3.Rows[27]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr7);

            DataRow dr8 = ResutTable.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["1month9"] = Convert.ToDecimal(newDT1.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["1month10"] = Convert.ToDecimal(newDT2.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["1month11"] = Convert.ToDecimal(newDT3.Rows[34]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr8["2month9"] = Convert.ToDecimal(newDT1.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["2month10"] = Convert.ToDecimal(newDT2.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["2month11"] = Convert.ToDecimal(newDT3.Rows[33]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["3month9"] = Convert.ToDecimal(newDT1.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["3month10"] = Convert.ToDecimal(newDT2.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["3month11"] = Convert.ToDecimal(newDT3.Rows[31]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["4month9"] = Convert.ToDecimal(newDT1.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["4month10"] = Convert.ToDecimal(newDT2.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["4month11"] = Convert.ToDecimal(newDT3.Rows[30]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["5month9"] = Convert.ToDecimal(newDT1.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["5month10"] = Convert.ToDecimal(newDT2.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr8["5month11"] = Convert.ToDecimal(newDT3.Rows[32]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr8);

            DataRow dr9 = ResutTable.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1month9"] = Convert.ToDecimal(newDT1.Rows[74]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["1month10"] = Convert.ToDecimal(newDT2.Rows[74]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["1month11"] = Convert.ToDecimal(newDT3.Rows[74]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["2month9"] = Convert.ToDecimal(newDT1.Rows[73]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["2month10"] = Convert.ToDecimal(newDT2.Rows[73]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["2month11"] = Convert.ToDecimal(newDT3.Rows[73]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["3month9"] = Convert.ToDecimal(newDT1.Rows[71]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["3month10"] = Convert.ToDecimal(newDT2.Rows[71]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["3month11"] = Convert.ToDecimal(newDT3.Rows[71]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["4month9"] = Convert.ToDecimal(newDT1.Rows[70]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr9["4month10"] = Convert.ToDecimal(newDT2.Rows[70]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["4month11"] = Convert.ToDecimal(newDT3.Rows[70]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["5month9"] = Convert.ToDecimal(newDT1.Rows[72]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["5month10"] = Convert.ToDecimal(newDT2.Rows[72]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr9["5month11"] = Convert.ToDecimal(newDT3.Rows[72]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr9);

            DataRow dr10 = ResutTable.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1month9"] = Convert.ToDecimal(newDT1.Rows[79]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["1month10"] = Convert.ToDecimal(newDT2.Rows[79]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["1month11"] = Convert.ToDecimal(newDT3.Rows[79]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["2month9"] = Convert.ToDecimal(newDT1.Rows[78]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["2month10"] = Convert.ToDecimal(newDT2.Rows[78]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["2month11"] = Convert.ToDecimal(newDT3.Rows[78]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["3month9"] = Convert.ToDecimal(newDT1.Rows[76]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["3month10"] = Convert.ToDecimal(newDT2.Rows[76]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["3month11"] = Convert.ToDecimal(newDT3.Rows[76]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr10["4month9"] = Convert.ToDecimal(newDT1.Rows[75]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["4month10"] = Convert.ToDecimal(newDT2.Rows[75]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["4month11"] = Convert.ToDecimal(newDT3.Rows[75]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["5month9"] = Convert.ToDecimal(newDT1.Rows[77]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["5month10"] = Convert.ToDecimal(newDT2.Rows[77]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr10["5month11"] = Convert.ToDecimal(newDT3.Rows[77]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr10);

            DataRow dr11 = ResutTable.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1month9"] = Convert.ToDecimal(newDT1.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["1month10"] = Convert.ToDecimal(newDT2.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["1month11"] = Convert.ToDecimal(newDT3.Rows[14]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["2month9"] = Convert.ToDecimal(newDT1.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["2month10"] = Convert.ToDecimal(newDT2.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["2month11"] = Convert.ToDecimal(newDT3.Rows[13]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr11["3month9"] = Convert.ToDecimal(newDT1.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["3month10"] = Convert.ToDecimal(newDT2.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["3month11"] = Convert.ToDecimal(newDT3.Rows[11]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["4month9"] = Convert.ToDecimal(newDT1.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["4month10"] = Convert.ToDecimal(newDT2.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["4month11"] = Convert.ToDecimal(newDT3.Rows[10]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["5month9"] = Convert.ToDecimal(newDT1.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["5month10"] = Convert.ToDecimal(newDT2.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr11["5month11"] = Convert.ToDecimal(newDT3.Rows[12]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr11);

            DataRow dr12 = ResutTable.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["1month9"] = Convert.ToDecimal(newDT1.Rows[39]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["1month10"] = Convert.ToDecimal(newDT2.Rows[39]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["1month11"] = Convert.ToDecimal(newDT3.Rows[39]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["2month9"] = Convert.ToDecimal(newDT1.Rows[38]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["2month10"] = Convert.ToDecimal(newDT2.Rows[38]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["2month11"] = Convert.ToDecimal(newDT3.Rows[38]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["3month9"] = Convert.ToDecimal(newDT1.Rows[36]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["3month10"] = Convert.ToDecimal(newDT2.Rows[36]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["3month11"] = Convert.ToDecimal(newDT3.Rows[36]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["4month9"] = Convert.ToDecimal(newDT1.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr12["4month10"] = Convert.ToDecimal(newDT2.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["4month11"] = Convert.ToDecimal(newDT3.Rows[35]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["5month9"] = Convert.ToDecimal(newDT1.Rows[37]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["5month10"] = Convert.ToDecimal(newDT2.Rows[37]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr12["5month11"] = Convert.ToDecimal(newDT3.Rows[37]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr12);

            DataRow dr13 = ResutTable.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["1month9"] = Convert.ToDecimal(newDT1.Rows[44]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["1month10"] = Convert.ToDecimal(newDT2.Rows[44]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["1month11"] = Convert.ToDecimal(newDT3.Rows[44]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["2month9"] = Convert.ToDecimal(newDT1.Rows[43]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["2month10"] = Convert.ToDecimal(newDT2.Rows[43]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["2month11"] = Convert.ToDecimal(newDT3.Rows[43]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr13["3month9"] = Convert.ToDecimal(newDT1.Rows[41]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["3month10"] = Convert.ToDecimal(newDT2.Rows[41]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["3month11"] = Convert.ToDecimal(newDT3.Rows[41]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["4month9"] = Convert.ToDecimal(newDT1.Rows[40]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["4month10"] = Convert.ToDecimal(newDT2.Rows[40]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["4month11"] = Convert.ToDecimal(newDT3.Rows[40]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["5month9"] = Convert.ToDecimal(newDT1.Rows[42]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["5month10"] = Convert.ToDecimal(newDT2.Rows[42]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr13["5month11"] = Convert.ToDecimal(newDT3.Rows[42]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr13);

            DataRow dr14 = ResutTable.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "1#";
            dr14["1month9"] = Convert.ToDecimal(newDT1.Rows[49]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr14["1month10"] = Convert.ToDecimal(newDT2.Rows[49]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["1month11"] = Convert.ToDecimal(newDT3.Rows[49]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["2month9"] = Convert.ToDecimal(newDT1.Rows[48]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["2month10"] = Convert.ToDecimal(newDT2.Rows[48]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["2month11"] = Convert.ToDecimal(newDT3.Rows[48]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["3month9"] = Convert.ToDecimal(newDT1.Rows[46]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["3month10"] = Convert.ToDecimal(newDT2.Rows[46]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["3month11"] = Convert.ToDecimal(newDT3.Rows[46]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["4month9"] = Convert.ToDecimal(newDT1.Rows[45]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["4month10"] = Convert.ToDecimal(newDT2.Rows[45]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["4month11"] = Convert.ToDecimal(newDT3.Rows[45]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["5month9"] = Convert.ToDecimal(newDT1.Rows[47]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["5month10"] = Convert.ToDecimal(newDT2.Rows[47]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr14["5month11"] = Convert.ToDecimal(newDT3.Rows[47]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr14);

            DataRow dr15 = ResutTable.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["1month9"] = Convert.ToDecimal(newDT1.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");//原料调配
            dr15["1month10"] = Convert.ToDecimal(newDT2.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["1month11"] = Convert.ToDecimal(newDT3.Rows[4]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["2month9"] = Convert.ToDecimal(newDT1.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");//生料粉磨  
            dr15["2month10"] = Convert.ToDecimal(newDT2.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["2month11"] = Convert.ToDecimal(newDT3.Rows[3]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["3month9"] = Convert.ToDecimal(newDT1.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");//煤粉制备
            dr15["3month10"] = Convert.ToDecimal(newDT2.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["3month11"] = Convert.ToDecimal(newDT3.Rows[1]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["4month9"] = Convert.ToDecimal(newDT1.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");//熟料烧成
            dr15["4month10"] = Convert.ToDecimal(newDT2.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["4month11"] = Convert.ToDecimal(newDT3.Rows[0]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["5month9"] = Convert.ToDecimal(newDT1.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");//废气处理
            dr15["5month10"] = Convert.ToDecimal(newDT2.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr15["5month11"] = Convert.ToDecimal(newDT3.Rows[2]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr15);

            DataRow dr16 = ResutTable.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["1month9"] = Convert.ToDecimal(newDT1.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");//4 3 5 2 1
            dr16["1month10"] = Convert.ToDecimal(newDT2.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["1month11"] = Convert.ToDecimal(newDT3.Rows[9]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["2month9"] = Convert.ToDecimal(newDT1.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["2month10"] = Convert.ToDecimal(newDT2.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["2month11"] = Convert.ToDecimal(newDT3.Rows[8]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["3month9"] = Convert.ToDecimal(newDT1.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["3month10"] = Convert.ToDecimal(newDT2.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["3month11"] = Convert.ToDecimal(newDT3.Rows[6]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["4month9"] = Convert.ToDecimal(newDT1.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["4month10"] = Convert.ToDecimal(newDT2.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["4month11"] = Convert.ToDecimal(newDT3.Rows[5]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["5month9"] = Convert.ToDecimal(newDT1.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["5month10"] = Convert.ToDecimal(newDT2.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            dr16["5month11"] = Convert.ToDecimal(newDT3.Rows[7]["TotalPeakValleyFlatB"]).ToString("0.00");
            ResutTable.Rows.Add(dr16);

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

            DataColumn dc3 = new DataColumn("1month9", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("1month10", Type.GetType("System.String"));
            DataColumn dc5 = new DataColumn("1month11", Type.GetType("System.String"));

            DataColumn dc6 = new DataColumn("2month9", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("2month10", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("2month11", Type.GetType("System.String"));

            DataColumn dc9 = new DataColumn("3month9", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("3month10", Type.GetType("System.String"));
            DataColumn dc11 = new DataColumn("3month11", Type.GetType("System.String"));

            DataColumn dc12 = new DataColumn("4month9", Type.GetType("System.String"));
            DataColumn dc13 = new DataColumn("4month10", Type.GetType("System.String"));
            DataColumn dc14 = new DataColumn("4month11", Type.GetType("System.String"));

            DataColumn dc15 = new DataColumn("5month9", Type.GetType("System.String"));
            DataColumn dc16 = new DataColumn("5month10", Type.GetType("System.String"));
            DataColumn dc17 = new DataColumn("5month11", Type.GetType("System.String"));

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
            table.Columns.Add(dc12);
            table.Columns.Add(dc13);
            table.Columns.Add(dc14);
            table.Columns.Add(dc15);
            table.Columns.Add(dc16);
            table.Columns.Add(dc17);
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

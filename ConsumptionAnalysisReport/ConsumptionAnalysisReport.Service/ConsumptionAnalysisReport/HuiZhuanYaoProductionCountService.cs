using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SqlServerDataAdapter;
using ConsumptionAnalysisReport.Infrastructure.Configuration;

namespace ConsumptionAnalysisReport.Service.ConsumptionAnalysisReport
{
    public class HuiZhuanYaoProductionCountService
    {
        public static DataTable GetConsumptionAnalysisTable(string searchDate, string[] myOganizationIds)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);

            Dictionary<string, string> m_ColumnIndexContrast = new Dictionary<string, string>();
            m_ColumnIndexContrast.Add("clinker_MixtureMaterialsOutput", "2");
            m_ColumnIndexContrast.Add("clinker_PulverizedCoalOutput", "5");
            m_ColumnIndexContrast.Add("clinker_ClinkerOutput", "8");
            m_ColumnIndexContrast.Add("MixtureMaterialsOutput_Checked", "1");
            m_ColumnIndexContrast.Add("PulverizedCoalOutput_Checked", "4");
            m_ColumnIndexContrast.Add("ClinkerOutput_Checked", "7");
            m_ColumnIndexContrast.Add("MixtureMaterialsOutput_Ratio", "3");
            m_ColumnIndexContrast.Add("PulverizedCoalOutput_Ratio", "6");
            m_ColumnIndexContrast.Add("ClinkerOutput_Ratio", "9");

            DataTable table = GetDCSProductionValue(searchDate, myOganizationIds, dataFactory);
            DataTable CheckedDataTable = GetCheckedProductionValue(searchDate, myOganizationIds, dataFactory);
            DataTable m_ResultTable = GetResultDataTable();
            if (table != null)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    string m_CompanyNameTemp = table.Rows[i]["CompanyName"].ToString();
                    string m_ProductLineTemp = table.Rows[i]["ProductionLine"].ToString();
                    string m_VariableIdTemp = table.Rows[i]["VariableId"].ToString();
                    DataRow[] m_ResultRowTemp = m_ResultTable.Select(string.Format("CompanyName = '{0}' and ProductionLine = '{1}'", m_CompanyNameTemp, m_ProductLineTemp));
                    if (m_ResultRowTemp.Length > 0)          //表示已经有改行
                    {
                        m_ResultRowTemp[0][m_ColumnIndexContrast[m_VariableIdTemp]] = table.Rows[i]["TotalPeakValleyFlatB"] != DBNull.Value ? (decimal)table.Rows[i]["TotalPeakValleyFlatB"] : 0.00m;
                    }
                    else
                    {
                        DataRow m_NewRowTemp = m_ResultTable.NewRow();
                        m_NewRowTemp["CompanyName"] = m_CompanyNameTemp;
                        m_NewRowTemp["ProductionLine"] = m_ProductLineTemp;
                        m_NewRowTemp[m_ColumnIndexContrast[m_VariableIdTemp]] = table.Rows[i]["TotalPeakValleyFlatB"] != DBNull.Value ? (decimal)table.Rows[i]["TotalPeakValleyFlatB"] : 0.00m;
                        m_ResultTable.Rows.Add(m_NewRowTemp);
                    }
                }
            }
            if (CheckedDataTable != null)
            {
                for (int i = 0; i < CheckedDataTable.Rows.Count; i++)
                {
                    string m_CompanyNameTemp = CheckedDataTable.Rows[i]["CompanyName"].ToString();
                    string m_ProductLineTemp = CheckedDataTable.Rows[i]["ProductionLine"].ToString();
                    string m_VariableIdTemp = CheckedDataTable.Rows[i]["VariableId"].ToString();
                    DataRow[] m_ResultRowTemp = m_ResultTable.Select(string.Format("CompanyName = '{0}' and ProductionLine = '{1}'", m_CompanyNameTemp, m_ProductLineTemp));
                    if (m_ResultRowTemp.Length > 0)          //表示已经有改行
                    {
                        m_ResultRowTemp[0][m_ColumnIndexContrast[m_VariableIdTemp]] = CheckedDataTable.Rows[i]["DataValue"] != DBNull.Value ? (decimal)CheckedDataTable.Rows[i]["DataValue"] : 0.00m;
                    }
                    else
                    {
                        DataRow m_NewRowTemp = m_ResultTable.NewRow();
                        m_NewRowTemp["CompanyName"] = m_CompanyNameTemp;
                        m_NewRowTemp["ProductionLine"] = m_ProductLineTemp;
                        m_NewRowTemp[m_ColumnIndexContrast[m_VariableIdTemp]] = CheckedDataTable.Rows[i]["DataValue"] != DBNull.Value ? (decimal)CheckedDataTable.Rows[i]["DataValue"] : 0.00m;
                        m_ResultTable.Rows.Add(m_NewRowTemp);
                    }
                }
            }
            for (int i = 0; i < m_ResultTable.Rows.Count; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    //string cc = m_ResultTable.Rows[i][(1 * j + 3).ToString()].ToString();
                    if (m_ResultTable.Rows[i][(3 * j + 1).ToString()] != DBNull.Value && m_ResultTable.Rows[i][(3 * j + 2).ToString()] != DBNull.Value)
                    {
                        m_ResultTable.Rows[i][(3 * j + 3).ToString()] = (100 * ((decimal)m_ResultTable.Rows[i][(3 * j + 1).ToString()] - (decimal)m_ResultTable.Rows[i][(3 * j + 2).ToString()]) / (decimal)m_ResultTable.Rows[i][(3 * j + 1).ToString()]).ToString("00.00") + "%";
                    }
                }
            }
            return m_ResultTable;
        }
        private static DataTable GetDCSProductionValue(string searchDate, string[] myOganizationIds, ISqlServerDataFactory dataFactory)
        {
            string m_VariableId = "'clinker_ClinkerOutput', 'clinker_MixtureMaterialsOutput','clinker_PulverizedCoalOutput'";
            string m_OrganizationCondition = "";
            for (int i = 0; i < myOganizationIds.Length; i++)
            {
                if (i == 0)
                {
                    m_OrganizationCondition = "'" + myOganizationIds[i] + "'";
                }
                else
                {
                    m_OrganizationCondition = m_OrganizationCondition + ",'" + myOganizationIds[i] + "'";
                }
            }
            string queryString = @"Select E.Name as CompanyName
	                                    ,F.Name as FacotoryName
	                                    ,replace(replace(replace(replace(D.Name,'号','#'),'窑',''),'熟料',''),'线','') as ProductionLine
	                                    ,B.VariableId
	                                    ,sum(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB
                                from tz_Balance A, balance_Energy B, system_Organization C, system_Organization D, system_Organization E, system_Organization F
                                where A.StaticsCycle = 'day'
                                and A.TimeStamp like '{0}%'
                                and A.BalanceId = B.KeyId
                                and B.ValueType = 'MaterialWeight'
                                and B.VariableId in ({2})
                                and B.OrganizationID = D.OrganizationID
                                and C.OrganizationID in ({1})
                                and D.levelCode like C.LevelCode + '%'
                                and D.LevelType = 'ProductionLine'
                                and charindex(E.LevelCode, D.LevelCode) > 0
                                and E.LevelType = 'Company'
                                and charindex(F.LevelCode, D.LevelCode) > 0
                                and F.LevelType = 'Factory'
                                group by E.Name, F.Name, D.Name, B.VariableId
                                order by E.Name, D.Name, B.VariableId";
            try
            {
                DataTable table = dataFactory.Query(string.Format(queryString, searchDate, m_OrganizationCondition, m_VariableId));
                return table;
            }
            catch
            {
                return null;
            }
        }
        private static DataTable GetCheckedProductionValue(string searchDate, string[] myOganizationIds, ISqlServerDataFactory dataFactory)
        {
            string m_MaxTimeCondtion = GetMaxTimeCondtion(searchDate, dataFactory);
            string m_VariableId = "'ClinkerOutput_Checked', 'MixtureMaterialsOutput_Checked','PulverizedCoalOutput_Checked'";
            string m_OrganizationCondition = "";
            for (int i = 0; i < myOganizationIds.Length; i++)
            {
                if (i == 0)
                {
                    m_OrganizationCondition = "'" + myOganizationIds[i] + "'";
                }
                else
                {
                    m_OrganizationCondition = m_OrganizationCondition + ",'" + myOganizationIds[i] + "'";
                }
            }
            string queryString = @"Select E.Name as CompanyName
	                                    ,F.Name as FacotoryName
	                                    ,replace(replace(replace(replace(D.Name,'号','#'),'窑',''),'熟料',''),'线','') as ProductionLine
	                                    ,B.VariableId
	                                    ,B.[DataValue] as [DataValue]
                                from system_EnergyDataManualInput B, system_Organization C, system_Organization D, system_Organization E, system_Organization F
                                where B.VariableId in ({2})
                                {0}
                                and B.OrganizationID = D.OrganizationID
                                and C.OrganizationID in ({1})
                                and D.levelCode like C.LevelCode + '%'
                                and D.LevelType = 'ProductionLine'
                                and charindex(E.LevelCode, D.LevelCode) > 0
                                and E.LevelType = 'Company'
                                and charindex(F.LevelCode, D.LevelCode) > 0
                                and F.LevelType = 'Factory'
                                order by E.Name, D.Name, B.VariableId";
            try
            {
                queryString = string.Format(queryString, m_MaxTimeCondtion, m_OrganizationCondition, m_VariableId);
                DataTable table = dataFactory.Query(queryString);
                return table;
            }
            catch
            {
                return null;
            }
        }
        private static string GetMaxTimeCondtion(string searchDate, ISqlServerDataFactory dataFactory)
        {
            DateTime m_StartTime = DateTime.Parse(searchDate + "-01");
            string queryString = @"Select A.VariableId, max(A.TimeStamp) as TimeStamp from system_EnergyDataManualInput A 
                                     where A.TimeStamp >= '{0}' and A.TimeStamp <= '{1}'
                                     group by A.VariableId";
            queryString = string.Format(queryString, m_StartTime.ToString("yyyy-MM-dd"), m_StartTime.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd"));
            string m_ReturnString = "";
            try
            {
                DataTable table = dataFactory.Query(queryString);
                if (table != null && table.Rows.Count > 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        if(i==0)
                        {
                            m_ReturnString = string.Format(" (B.VariableId = '{0}' and B.TimeStamp = '{1}')", table.Rows[i]["VariableId"].ToString(), table.Rows[i]["TimeStamp"].ToString());
                        }
                        else
                        {
                            m_ReturnString = m_ReturnString + string.Format(" or (B.VariableId = '{0}' and B.TimeStamp = '{1}')", table.Rows[i]["VariableId"].ToString(), table.Rows[i]["TimeStamp"].ToString());
                        }
                    }
                    m_ReturnString = " and (" + m_ReturnString + ")";
                }
                else
                {
                    m_ReturnString = string.Format(" and B.TimeStamp >= '{0}' and B.TimeStamp <= '{1}' ", m_StartTime.ToString("yyyy-MM-dd"), m_StartTime.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd"));
                }
                return m_ReturnString;
            }
            catch
            {
                m_ReturnString = string.Format(" and B.TimeStamp >= '{0}' and B.TimeStamp <= '{1}' ", m_StartTime.ToString("yyyy-MM-dd"), m_StartTime.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd"));
                return m_ReturnString;
            }
        }
        private static DataTable GetResultDataTable()
        {
            //最后的table
            DataTable m_resultTable = new DataTable();
            DataColumn dc1 = new DataColumn("CompanyName", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("ProductionLine", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("1", typeof(decimal));
            DataColumn dc4 = new DataColumn("4", typeof(decimal));
            DataColumn dc5 = new DataColumn("7", typeof(decimal));//盘库
            DataColumn dc6 = new DataColumn("2", typeof(decimal));
            DataColumn dc7 = new DataColumn("5", typeof(decimal));
            DataColumn dc8 = new DataColumn("8", typeof(decimal));//DCS
            DataColumn dc9 = new DataColumn("3", typeof(string));
            DataColumn dc10 = new DataColumn("6", typeof(string));
            DataColumn dc11 = new DataColumn("9", typeof(string));//增减比例
            m_resultTable.Columns.Add(dc1);
            m_resultTable.Columns.Add(dc2);
            m_resultTable.Columns.Add(dc3);
            m_resultTable.Columns.Add(dc4);
            m_resultTable.Columns.Add(dc5);
            m_resultTable.Columns.Add(dc6);
            m_resultTable.Columns.Add(dc7);
            m_resultTable.Columns.Add(dc8);
            m_resultTable.Columns.Add(dc9);
            m_resultTable.Columns.Add(dc10);
            m_resultTable.Columns.Add(dc11);

            return m_resultTable;
        }
        public static DataTable GetConsumptionAnalysisTable(string m_SelectTime)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string m_Sqlstr = @" Select C.Name AS CompanyName, M.Name as FactoryName, F.OrganizationID as FactoryOrgID, N.Name + B.Name AS VariableName, F.PeakB AS PeakB,F.ValleyB AS ValleyB,F.FlatB AS FlatB,(F.PeakB+F.ValleyB+F.FlatB) AS TotalProduction
                                        from tz_Material A, material_MaterialDetail B, system_Organization M
                                        left join system_Organization C on substring(M.LevelCode,1,Len(M.LevelCode) - 2) = C.LevelCode, system_Organization N, 
                                            (Select E.OrganizationID, E.VariableId, SUM(E.FirstB) AS FirstB,SUM(E.SecondB) AS SecondB,SUM(E.ThirdB) AS ThirdB,SUM(E.TotalPeakValleyFlatB) AS TotalPeakValleyFlatB,SUM(E.PeakB) AS PeakB,SUM(E.ValleyB) AS ValleyB,SUM(E.FlatB) AS FlatB,
                                    SUM(CASE WHEN [D].[FirstWorkingTeam] = 'A班' THEN [E].[FirstB] WHEN [D].[SecondWorkingTeam] = 'A班' THEN [E].[SecondB] WHEN [D].[ThirdWorkingTeam] = 'A班' THEN [E].[ThirdB] ELSE 0 END) AS A班,
		                            SUM(CASE WHEN [D].[FirstWorkingTeam] = 'B班' THEN [E].[FirstB] WHEN [D].[SecondWorkingTeam] = 'B班' THEN [E].[SecondB] WHEN [D].[ThirdWorkingTeam] = 'B班' THEN [E].[ThirdB] ELSE 0 END) AS B班,
		                            SUM(CASE WHEN [D].[FirstWorkingTeam] = 'C班' THEN [E].[FirstB] WHEN [D].[SecondWorkingTeam] = 'C班' THEN [E].[SecondB] WHEN [D].[ThirdWorkingTeam] = 'C班' THEN [E].[ThirdB] ELSE 0 END) AS C班,
		                            SUM(CASE WHEN [D].[FirstWorkingTeam] = 'D班' THEN [E].[FirstB] WHEN [D].[SecondWorkingTeam] = 'D班' THEN [E].[SecondB] WHEN [D].[ThirdWorkingTeam] = 'D班' THEN [E].[ThirdB] ELSE 0 END) AS D班
	                                            from tz_Balance D, balance_Energy E 
		                                        where D.TimeStamp like '{0}'+'%'		                                        
		                                        and D.StaticsCycle = 'day'
		                                        and D.BalanceId = E.KeyId
		                                        and E.ValueType = 'MaterialWeight'  
		                                        and D.OrganizationID in ('zc_nxjc_qtx_efc','zc_nxjc_ychc_lsf','zc_nxjc_ychc_yfcf','zc_nxjc_qtx_tys','zc_nxjc_znc_znf','zc_nxjc_lpsc_lpsf','zc_nxjc_tsc_tsf','zc_nxjc_whsmc_whsmf','zc_nxjc_byc_byf','zc_nxjc_klqc_klqf')
		                                        group by E.OrganizationID, E.VariableId) F
                                        where M.OrganizationID in ('zc_nxjc_qtx_efc','zc_nxjc_ychc_lsf','zc_nxjc_ychc_yfcf','zc_nxjc_qtx_tys','zc_nxjc_znc_znf','zc_nxjc_lpsc_lpsf','zc_nxjc_tsc_tsf','zc_nxjc_whsmc_whsmf','zc_nxjc_byc_byf','zc_nxjc_klqc_klqf') 
                                        and N.LevelCode like M.LevelCode + '%'
                                        and A.OrganizationID in (N.OrganizationID)
                                        and A.Enable = 1
                                        and A.State = 0
                                        and A.KeyID = B.KeyID
                                        and B.Visible = 1
                                        and A.OrganizationID = F.OrganizationID
                                        and B.VariableId = F.VariableId
                                        order by FactoryOrgID";
            m_Sqlstr = string.Format(m_Sqlstr, m_SelectTime);
            DataTable table = dataFactory.Query(m_Sqlstr);

            DataTable newdt = new DataTable();
            newdt = table.Clone();
            DataRow[] dr = table.Select("VariableName like '%生料产量' or VariableName like '%熟料产量' or VariableName like '%煤粉产量'", "FactoryOrgID");
            for (int i = 0; i < dr.Length; i++)
            {
                newdt.ImportRow((DataRow)dr[i]);
            }
            //盘库的table
            DataTable wareHouseTable = GetFixWarehouseValue(m_SelectTime);
            //最后的table
            DataTable m_resultTable = new DataTable();
            DataColumn dc1 = new DataColumn("CompanyName", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("ProductionLine", Type.GetType("System.String"));
            DataColumn dc3 = new DataColumn("1", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("4", Type.GetType("System.String"));
            DataColumn dc5 = new DataColumn("7", Type.GetType("System.String"));//盘库
            DataColumn dc6 = new DataColumn("2", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("5", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("8", Type.GetType("System.String"));//DCS
            DataColumn dc9 = new DataColumn("3", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("6", Type.GetType("System.String"));
            DataColumn dc11 = new DataColumn("9", Type.GetType("System.String"));//增减比例
            m_resultTable.Columns.Add(dc1);
            m_resultTable.Columns.Add(dc2);
            m_resultTable.Columns.Add(dc3);
            m_resultTable.Columns.Add(dc4);
            m_resultTable.Columns.Add(dc5);
            m_resultTable.Columns.Add(dc6);
            m_resultTable.Columns.Add(dc7);
            m_resultTable.Columns.Add(dc8);
            m_resultTable.Columns.Add(dc9);
            m_resultTable.Columns.Add(dc10);
            m_resultTable.Columns.Add(dc11);

            DataRow dr1 = m_resultTable.NewRow();
            dr1["CompanyName"] = "银川水泥";
            dr1["ProductionLine"] = "1#";
            dr1["1"] = wareHouseTable.Rows[0]["RawMaterial"];//生料
            dr1["4"] = wareHouseTable.Rows[0]["Coal"];//煤粉
            dr1["7"] = wareHouseTable.Rows[0]["Clinker"]; ;//熟料
            dr1["2"] = Convert.ToDecimal(newdt.Rows[40]["TotalProduction"]).ToString("0.00");//生料
            dr1["5"] = Convert.ToDecimal(newdt.Rows[39]["TotalProduction"]).ToString("0.00");//煤粉
            dr1["8"] = Convert.ToDecimal(newdt.Rows[41]["TotalProduction"]).ToString("0.00");//熟料
            dr1["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[0]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[40]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[40]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr1["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[0]["Coal"]) - Convert.ToDecimal(newdt.Rows[39]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[39]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr1["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[0]["Clinker"]) - Convert.ToDecimal(newdt.Rows[41]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[41]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr1);

            DataRow dr2 = m_resultTable.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["1"] = wareHouseTable.Rows[1]["RawMaterial"];//生料
            dr2["4"] = wareHouseTable.Rows[1]["Coal"];//煤粉
            dr2["7"] = wareHouseTable.Rows[1]["Clinker"]; ;//熟料
            dr2["2"] = Convert.ToDecimal(newdt.Rows[31]["TotalProduction"]).ToString("0.00");//生料
            dr2["5"] = Convert.ToDecimal(newdt.Rows[30]["TotalProduction"]).ToString("0.00");//煤粉
            dr2["8"] = Convert.ToDecimal(newdt.Rows[32]["TotalProduction"]).ToString("0.00");//熟料
            dr2["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[1]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[31]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[31]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr2["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[1]["Coal"]) - Convert.ToDecimal(newdt.Rows[30]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[30]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr2["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[1]["Clinker"]) - Convert.ToDecimal(newdt.Rows[32]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[32]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr2);

            DataRow dr3 = m_resultTable.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["1"] = wareHouseTable.Rows[2]["RawMaterial"];//生料
            dr3["4"] = wareHouseTable.Rows[2]["Coal"];//煤粉
            dr3["7"] = wareHouseTable.Rows[2]["Clinker"]; ;//熟料
            dr3["2"] = Convert.ToDecimal(newdt.Rows[34]["TotalProduction"]).ToString("0.00");//生料
            dr3["5"] = Convert.ToDecimal(newdt.Rows[33]["TotalProduction"]).ToString("0.00");//煤粉
            dr3["8"] = Convert.ToDecimal(newdt.Rows[35]["TotalProduction"]).ToString("0.00");//熟料
            dr3["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[2]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[34]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[34]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr3["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[2]["Coal"]) - Convert.ToDecimal(newdt.Rows[33]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[33]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr3["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[2]["Clinker"]) - Convert.ToDecimal(newdt.Rows[35]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[35]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr3);

            DataRow dr4 = m_resultTable.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1"] = wareHouseTable.Rows[3]["RawMaterial"];//生料
            dr4["4"] = wareHouseTable.Rows[3]["Coal"];//煤粉
            dr4["7"] = wareHouseTable.Rows[3]["Clinker"]; ;//熟料
            dr4["2"] = Convert.ToDecimal(newdt.Rows[37]["TotalProduction"]).ToString("0.00");//生料
            dr4["5"] = Convert.ToDecimal(newdt.Rows[36]["TotalProduction"]).ToString("0.00");//煤粉
            dr4["8"] = Convert.ToDecimal(newdt.Rows[38]["TotalProduction"]).ToString("0.00");//熟料
            dr4["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[3]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[37]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[37]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr4["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[3]["Coal"]) - Convert.ToDecimal(newdt.Rows[36]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[36]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr4["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[3]["Clinker"]) - Convert.ToDecimal(newdt.Rows[38]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[38]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr4);

            DataRow dr5 = m_resultTable.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["1"] = wareHouseTable.Rows[4]["RawMaterial"];//生料
            dr5["4"] = wareHouseTable.Rows[4]["Coal"];//煤粉
            dr5["7"] = wareHouseTable.Rows[4]["Clinker"]; ;//熟料
            dr5["2"] = Convert.ToDecimal(newdt.Rows[10]["TotalProduction"]).ToString("0.00");//生料
            dr5["5"] = Convert.ToDecimal(newdt.Rows[9]["TotalProduction"]).ToString("0.00");//煤粉
            dr5["8"] = Convert.ToDecimal(newdt.Rows[11]["TotalProduction"]).ToString("0.00");//熟料
            dr5["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[4]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[10]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[10]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr5["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[4]["Coal"]) - Convert.ToDecimal(newdt.Rows[9]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[9]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr5["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[4]["Clinker"]) - Convert.ToDecimal(newdt.Rows[11]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[11]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr5);

            DataRow dr6 = m_resultTable.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["1"] = wareHouseTable.Rows[5]["RawMaterial"];//生料
            dr6["4"] = wareHouseTable.Rows[5]["Coal"];//煤粉
            dr6["7"] = wareHouseTable.Rows[5]["Clinker"]; ;//熟料
            dr6["2"] = Convert.ToDecimal(newdt.Rows[13]["TotalProduction"]).ToString("0.00");//生料
            dr6["5"] = Convert.ToDecimal(newdt.Rows[12]["TotalProduction"]).ToString("0.00");//煤粉
            dr6["8"] = Convert.ToDecimal(newdt.Rows[14]["TotalProduction"]).ToString("0.00");//熟料
            dr6["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[5]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[13]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[13]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr6["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[5]["Coal"]) - Convert.ToDecimal(newdt.Rows[12]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[12]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr6["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[5]["Clinker"]) - Convert.ToDecimal(newdt.Rows[14]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[14]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr6);

            DataRow dr7 = m_resultTable.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["1"] = wareHouseTable.Rows[6]["RawMaterial"];//生料
            dr7["4"] = wareHouseTable.Rows[6]["Coal"];//煤粉
            dr7["7"] = wareHouseTable.Rows[6]["Clinker"]; ;//熟料
            dr7["2"] = Convert.ToDecimal(newdt.Rows[16]["TotalProduction"]).ToString("0.00");//生料
            dr7["5"] = Convert.ToDecimal(newdt.Rows[15]["TotalProduction"]).ToString("0.00");//煤粉
            dr7["8"] = Convert.ToDecimal(newdt.Rows[17]["TotalProduction"]).ToString("0.00");//熟料
            dr7["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[6]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[16]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[16]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr7["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[6]["Coal"]) - Convert.ToDecimal(newdt.Rows[15]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[15]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr7["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[6]["Clinker"]) - Convert.ToDecimal(newdt.Rows[17]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[17]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr7);

            DataRow dr8 = m_resultTable.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["1"] = wareHouseTable.Rows[7]["RawMaterial"];//生料
            dr8["4"] = wareHouseTable.Rows[7]["Coal"];//煤粉
            dr8["7"] = wareHouseTable.Rows[7]["Clinker"]; ;//熟料
            dr8["2"] = Convert.ToDecimal(newdt.Rows[19]["TotalProduction"]).ToString("0.00");//生料
            dr8["5"] = Convert.ToDecimal(newdt.Rows[18]["TotalProduction"]).ToString("0.00");//煤粉
            dr8["8"] = Convert.ToDecimal(newdt.Rows[20]["TotalProduction"]).ToString("0.00");//熟料
            dr8["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[7]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[19]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[19]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr8["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[7]["Coal"]) - Convert.ToDecimal(newdt.Rows[18]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[18]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr8["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[7]["Clinker"]) - Convert.ToDecimal(newdt.Rows[20]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[20]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr8);

            DataRow dr9 = m_resultTable.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1"] = wareHouseTable.Rows[8]["RawMaterial"];//生料
            dr9["4"] = wareHouseTable.Rows[8]["Coal"];//煤粉
            dr9["7"] = wareHouseTable.Rows[8]["Clinker"]; ;//熟料
            dr9["2"] = Convert.ToDecimal(newdt.Rows[43]["TotalProduction"]).ToString("0.00");//生料
            dr9["5"] = Convert.ToDecimal(newdt.Rows[42]["TotalProduction"]).ToString("0.00");//煤粉
            dr9["8"] = Convert.ToDecimal(newdt.Rows[44]["TotalProduction"]).ToString("0.00");//熟料
            dr9["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[8]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[43]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[43]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr9["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[8]["Coal"]) - Convert.ToDecimal(newdt.Rows[32]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[42]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr9["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[8]["Clinker"]) - Convert.ToDecimal(newdt.Rows[44]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[44]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr9);

            DataRow dr10 = m_resultTable.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1"] = wareHouseTable.Rows[9]["RawMaterial"];//生料
            dr10["4"] = wareHouseTable.Rows[9]["Coal"];//煤粉
            dr10["7"] = wareHouseTable.Rows[9]["Clinker"]; ;//熟料
            dr10["2"] = Convert.ToDecimal(newdt.Rows[46]["TotalProduction"]).ToString("0.00");//生料
            dr10["5"] = Convert.ToDecimal(newdt.Rows[45]["TotalProduction"]).ToString("0.00");//煤粉
            dr10["8"] = Convert.ToDecimal(newdt.Rows[47]["TotalProduction"]).ToString("0.00");//熟料
            dr10["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[9]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[46]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[46]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr10["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[9]["Coal"]) - Convert.ToDecimal(newdt.Rows[45]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[45]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr10["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[9]["Clinker"]) - Convert.ToDecimal(newdt.Rows[47]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[47]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr10);

            DataRow dr11 = m_resultTable.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1"] = wareHouseTable.Rows[10]["RawMaterial"];//生料
            dr11["4"] = wareHouseTable.Rows[10]["Coal"];//煤粉
            dr11["7"] = wareHouseTable.Rows[10]["Clinker"]; ;//熟料
            dr11["2"] = Convert.ToDecimal(newdt.Rows[7]["TotalProduction"]).ToString("0.00");//生料
            dr11["5"] = Convert.ToDecimal(newdt.Rows[6]["TotalProduction"]).ToString("0.00");//煤粉
            dr11["8"] = Convert.ToDecimal(newdt.Rows[8]["TotalProduction"]).ToString("0.00");//熟料
            dr11["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[10]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[7]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[7]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr11["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[10]["Coal"]) - Convert.ToDecimal(newdt.Rows[6]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[6]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr11["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[10]["Clinker"]) - Convert.ToDecimal(newdt.Rows[8]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[8]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr11);

            DataRow dr12 = m_resultTable.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["1"] = wareHouseTable.Rows[11]["RawMaterial"];//生料
            dr12["4"] = wareHouseTable.Rows[11]["Coal"];//煤粉
            dr12["7"] = wareHouseTable.Rows[11]["Clinker"]; ;//熟料
            dr12["2"] = Convert.ToDecimal(newdt.Rows[22]["TotalProduction"]).ToString("0.00");//生料
            dr12["5"] = Convert.ToDecimal(newdt.Rows[21]["TotalProduction"]).ToString("0.00");//煤粉
            dr12["8"] = Convert.ToDecimal(newdt.Rows[23]["TotalProduction"]).ToString("0.00");//熟料
            dr12["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[11]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[22]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[22]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr12["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[11]["Coal"]) - Convert.ToDecimal(newdt.Rows[21]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[21]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr12["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[11]["Clinker"]) - Convert.ToDecimal(newdt.Rows[23]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[23]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr12);

            DataRow dr13 = m_resultTable.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["1"] = wareHouseTable.Rows[12]["RawMaterial"];//生料
            dr13["4"] = wareHouseTable.Rows[12]["Coal"];//煤粉
            dr13["7"] = wareHouseTable.Rows[12]["Clinker"]; ;//熟料
            dr13["2"] = Convert.ToDecimal(newdt.Rows[25]["TotalProduction"]).ToString("0.00");//生料
            dr13["5"] = Convert.ToDecimal(newdt.Rows[24]["TotalProduction"]).ToString("0.00");//煤粉
            dr13["8"] = Convert.ToDecimal(newdt.Rows[26]["TotalProduction"]).ToString("0.00");//熟料
            dr13["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[12]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[25]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[25]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr13["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[12]["Coal"]) - Convert.ToDecimal(newdt.Rows[24]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[24]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr13["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[12]["Clinker"]) - Convert.ToDecimal(newdt.Rows[26]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[26]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr13);

            DataRow dr14 = m_resultTable.NewRow();
            dr14["CompanyName"] = "乌海水泥";
            dr14["ProductionLine"] = "1#";
            dr14["1"] = wareHouseTable.Rows[13]["RawMaterial"];//生料
            dr14["4"] = wareHouseTable.Rows[13]["Coal"];//煤粉
            dr14["7"] = wareHouseTable.Rows[13]["Clinker"]; ;//熟料
            dr14["2"] = Convert.ToDecimal(newdt.Rows[28]["TotalProduction"]).ToString("0.00");//生料
            dr14["5"] = Convert.ToDecimal(newdt.Rows[27]["TotalProduction"]).ToString("0.00");//煤粉
            dr14["8"] = Convert.ToDecimal(newdt.Rows[29]["TotalProduction"]).ToString("0.00");//熟料
            dr14["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[13]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[28]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[28]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr14["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[13]["Coal"]) - Convert.ToDecimal(newdt.Rows[27]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[27]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr14["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[13]["Clinker"]) - Convert.ToDecimal(newdt.Rows[29]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[29]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr14);

            DataRow dr15 = m_resultTable.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["1"] = wareHouseTable.Rows[14]["RawMaterial"];//生料
            dr15["4"] = wareHouseTable.Rows[14]["Coal"];//煤粉
            dr15["7"] = wareHouseTable.Rows[14]["Clinker"]; ;//熟料
            dr15["2"] = Convert.ToDecimal(newdt.Rows[1]["TotalProduction"]).ToString("0.00");//生料
            dr15["5"] = Convert.ToDecimal(newdt.Rows[0]["TotalProduction"]).ToString("0.00"); //煤粉
            dr15["8"] = Convert.ToDecimal(newdt.Rows[2]["TotalProduction"]).ToString("0.00");//熟料
            dr15["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[14]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[1]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[0]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr15["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[14]["Coal"]) - Convert.ToDecimal(newdt.Rows[0]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[1]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr15["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[14]["Clinker"]) - Convert.ToDecimal(newdt.Rows[2]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[2]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr15);

            DataRow dr16 = m_resultTable.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["1"] = wareHouseTable.Rows[15]["RawMaterial"];//生料
            dr16["4"] = wareHouseTable.Rows[15]["Coal"];//煤粉
            dr16["7"] = wareHouseTable.Rows[15]["Clinker"]; ;//熟料
            dr16["2"] = Convert.ToDecimal(newdt.Rows[4]["TotalProduction"]).ToString("0.00");//生料
            dr16["5"] = Convert.ToDecimal(newdt.Rows[3]["TotalProduction"]).ToString("0.00");//煤粉
            dr16["8"] = Convert.ToDecimal(newdt.Rows[5]["TotalProduction"]).ToString("0.00");//熟料
            dr16["3"] = ((Convert.ToDecimal(wareHouseTable.Rows[15]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[4]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[4]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr16["6"] = ((Convert.ToDecimal(wareHouseTable.Rows[15]["Coal"]) - Convert.ToDecimal(newdt.Rows[3]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[3]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr16["9"] = ((Convert.ToDecimal(wareHouseTable.Rows[15]["Clinker"]) - Convert.ToDecimal(newdt.Rows[5]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[5]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr16);
           
            return m_resultTable;
        }

        public static DataTable GetFixWarehouseValue(string m_selectTime)
        {
            DataTable table = new DataTable();
            DataColumn dc1 = new DataColumn("CheckWarehouseTime", Type.GetType("System.String"));//盘库时间
            DataColumn dc2 = new DataColumn("RawMaterial", Type.GetType("System.String"));//生料
            DataColumn dc3 = new DataColumn("Coal", Type.GetType("System.String"));//煤粉
            DataColumn dc4 = new DataColumn("Clinker", Type.GetType("System.String"));//熟料
            table.Columns.Add(dc1);
            table.Columns.Add(dc2);
            table.Columns.Add(dc3);
            table.Columns.Add(dc4);
            if (m_selectTime == "2017-11")
            {              
                DataRow dr1 = table.NewRow();
                dr1["CheckWarehouseTime"] = "2017-11";
                dr1["RawMaterial"] = "94328.1";
                dr1["Coal"] = "10134.8";
                dr1["Clinker"] = "59701.3";
                table.Rows.Add(dr1);

                DataRow dr2 = table.NewRow();
                dr2["CheckWarehouseTime"] = "2017-11";
                dr2["RawMaterial"] = "98860.9";
                dr2["Coal"] = "9682.6";
                dr2["Clinker"] = "62570.2";
                table.Rows.Add(dr2);

                DataRow dr3 = table.NewRow();
                dr3["CheckWarehouseTime"] = "2017-11";
                dr3["RawMaterial"] = "99052.4";
                dr3["Coal"] = "9712.0";
                dr3["Clinker"] = "62691.4";
                table.Rows.Add(dr3);

                DataRow dr4 = table.NewRow();
                dr4["CheckWarehouseTime"] = "2017-11";
                dr4["RawMaterial"] = "199939.5";
                dr4["Coal"] = "18221.6";
                dr4["Clinker"] = "126544.0";
                table.Rows.Add(dr4);

                DataRow dr5 = table.NewRow();
                dr5["CheckWarehouseTime"] = "2017-11";
                dr5["RawMaterial"] = "168264.0";
                dr5["Coal"] = "10742.0";
                dr5["Clinker"] = "70775.0";
                table.Rows.Add(dr5);

                DataRow dr6 = table.NewRow();
                dr6["CheckWarehouseTime"] = "2017-11";
                dr6["RawMaterial"] = "51937.0";
                dr6["Coal"] = "10649.0";
                dr6["Clinker"] = "71167.0";
                table.Rows.Add(dr6);

                DataRow dr7 = table.NewRow();
                dr7["CheckWarehouseTime"] = "2017-11";
                dr7["RawMaterial"] = "117833.0";
                dr7["Coal"] = "11774.0";
                dr7["Clinker"] = "75889.0";
                table.Rows.Add(dr7);

                DataRow dr8 = table.NewRow();
                dr8["CheckWarehouseTime"] = "2017-11";
                dr8["RawMaterial"] = "115099.0";
                dr8["Coal"] = "11122.0";
                dr8["Clinker"] = "71280.0";
                table.Rows.Add(dr8);

                DataRow dr9 = table.NewRow();
                dr9["CheckWarehouseTime"] = "2017-11";
                dr9["RawMaterial"] = "49367.0";
                dr9["Coal"] = "4567.6";
                dr9["Clinker"] = "3266.0";
                table.Rows.Add(dr9);

                DataRow dr10 = table.NewRow();
                dr10["CheckWarehouseTime"] = "2017-11";
                dr10["RawMaterial"] = "64448.0";
                dr10["Coal"] = "5973.9";
                dr10["Clinker"] = "42123.0";
                table.Rows.Add(dr10);

                DataRow dr11 = table.NewRow();
                dr11["CheckWarehouseTime"] = "2017-11";
                dr11["RawMaterial"] = "48458.6";
                dr11["Coal"] = "5130.0";
                dr11["Clinker"] = "30670.0";
                table.Rows.Add(dr11);

                DataRow dr12 = table.NewRow();
                dr12["CheckWarehouseTime"] = "2017-11";
                dr12["RawMaterial"] = "63184.1";
                dr12["Coal"] = "6982.4";
                dr12["Clinker"] = "40500.0";
                table.Rows.Add(dr12);

                DataRow dr13 = table.NewRow();
                dr13["CheckWarehouseTime"] = "2017-11";
                dr13["RawMaterial"] = "124210.0";
                dr13["Coal"] = "13729.0";
                dr13["Clinker"] = "79624.4";
                table.Rows.Add(dr13);

                DataRow dr14 = table.NewRow();
                dr14["CheckWarehouseTime"] = "2017-11";
                dr14["RawMaterial"] = "65091.2";
                dr14["Coal"] = "8815.6";
                dr14["Clinker"] = "41210.6";
                table.Rows.Add(dr14);

                DataRow dr15 = table.NewRow();
                dr15["CheckWarehouseTime"] = "2017-11";
                dr15["RawMaterial"] = "155150.4";
                dr15["Coal"] = "15236.0";
                dr15["Clinker"] = "100097.0";
                table.Rows.Add(dr15);

                DataRow dr16 = table.NewRow();
                dr16["CheckWarehouseTime"] = "2017-11";
                dr16["RawMaterial"] = "46305.0";
                dr16["Coal"] = "5716.0";
                dr16["Clinker"] = "30464.0";
                table.Rows.Add(dr16);
            }         
            return table;
        }

        public static DataTable GetFixConsumptionAnalysisTable()
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
            dr1["1"] = "94328.1";
            dr1["2"] = "130963.0";
            dr1["3"] = "-27.97%";
            dr1["4"] = "10134.8";
            dr1["5"] = "8461.5";
            dr1["6"] = "19.78%";
            dr1["7"] = "59701.3";
            dr1["8"] = "92053.8";
            dr1["9"] = "-35.15%";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["1"] = "98860.9";
            dr2["2"] = "90989.0";
            dr2["3"] = "8.65%";
            dr2["4"] = "9682.6";
            dr2["5"] = "11993.9";
            dr2["6"] = "-19.27%";
            dr2["7"] = "62570.2";
            dr2["8"] = "13551.0";
            dr2["9"] = "361.74%";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["1"] = "99052.4";
            dr3["2"] = "88761.0";
            dr3["3"] = "11.59%";
            dr3["4"] = "9712.0";
            dr3["5"] = "11812.1";
            dr3["6"] = "-17.78%";
            dr3["7"] = "62691.4";
            dr3["8"] = "63235.8";
            dr3["9"] = "-0.86%";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["1"] = "199939.5";
            dr4["2"] = "165138.0";
            dr4["3"] = "21.07%";
            dr4["4"] = "18221.6";
            dr4["5"] = "22806.7";
            dr4["6"] = "-20.10%";
            dr4["7"] = "126544.0";
            dr4["8"] = "122016.0";
            dr4["9"] = "3.71%";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["1"] = "168264.0";
            dr5["2"] = "158161.0";
            dr5["3"] = "6.39%";
            dr5["4"] = "10742.0";
            dr5["5"] = "12099.6";
            dr5["6"] = "-11.22%";
            dr5["7"] = "70775.0";
            dr5["8"] = "70950.6";
            dr5["9"] = "-0.25%";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["1"] = "51937.0";
            dr6["2"] = "45726.6";
            dr6["3"] = "13.58%";
            dr6["4"] = "10649.0";
            dr6["5"] = "11079.1";
            dr6["6"] = "-3.88%";
            dr6["7"] = "71167.0";
            dr6["8"] = "68715.0";
            dr6["9"] = "3.57%";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["1"] = "117833.0";
            dr7["2"] = "117619.3";
            dr7["3"] = "0.18%";
            dr7["4"] = "11774.0";
            dr7["5"] = "13667.1";
            dr7["6"] = "-13.85%";
            dr7["7"] = "75889.0";
            dr7["8"] = "74736.9";
            dr7["9"] = "1.54%";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["1"] = "115099.0";
            dr8["2"] = "113801.0";
            dr8["3"] = "1.14%";
            dr8["4"] = "11122.0";
            dr8["5"] = "11455.0";
            dr8["6"] = "-2.91%";
            dr8["7"] = "71280.0";
            dr8["8"] = "71250.3";
            dr8["9"] = "0.04%";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["1"] = "49367.0";
            dr9["2"] = "38311.5";
            dr9["3"] = "28.86%";
            dr9["4"] = "4537.6";
            dr9["5"] = "5451.4";
            dr9["6"] = "-16.21%";
            dr9["7"] = "32266.0";
            dr9["8"] = "26222.23";
            dr9["9"] = "23.05%";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["1"] = "64448.0";
            dr10["2"] = "54200.7";
            dr10["3"] = "18.91%";
            dr10["4"] = "5973.9";
            dr10["5"] = "7679.8";
            dr10["6"] = "-22.21%";
            dr10["7"] = "42123.0";
            dr10["8"] = "37155.36";
            dr10["9"] = "13.37%";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["1"] = "48458.6";
            dr11["2"] = "45490.5";
            dr11["3"] = "6.52%";
            dr11["4"] = "5130.0";
            dr11["5"] = "4380.9";
            dr11["6"] = "17.10%";
            dr11["7"] = "30670.0";
            dr11["8"] = "29540.0";
            dr11["9"] = "3.83%";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["1"] = "63184.1";
            dr12["2"] = "47691.0";
            dr12["3"] = "32.49%";
            dr12["4"] = "6982.4";
            dr12["5"] = "6398.5";
            dr12["6"] = "9.13%";
            dr12["7"] = "40500.0";
            dr12["8"] = "40985.0";
            dr12["9"] = "-1.18%";
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["1"] = "124210.0";
            dr13["2"] = "113292.0";
            dr13["3"] = "9.64%";
            dr13["4"] = "13729.0";
            dr13["5"] = "11113.0";
            dr13["6"] = "23.54%";
            dr13["7"] = "79624.4";
            dr13["8"] = "82183.1";
            dr13["9"] = "-3.11%";
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "1#";
            dr14["1"] = "65091.2";
            dr14["2"] = "56388.0";
            dr14["3"] = "15.43%";
            dr14["4"] = "8815.6";
            dr14["5"] = "5541.6";
            dr14["6"] = "59.08%";
            dr14["7"] = "41210.6";
            dr14["8"] = "44865.4";
            dr14["9"] = "-8.15%";
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["1"] = "155150.4";
            dr15["2"] = "136368.3";
            dr15["3"] = "13.77%";
            dr15["4"] = "15236.0";
            dr15["5"] = "18797.2";
            dr15["6"] = "-18.95%";
            dr15["7"] = "100097.0";
            dr15["8"] = "92676.6";
            dr15["9"] = "8.01%";
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["1"] = "46305.0";
            dr16["2"] = "35163.0";
            dr16["3"] = "31.69%";
            dr16["4"] = "5716.0";
            dr16["5"] = "5946.0";
            dr16["6"] = "-3.87%";
            dr16["7"] = "60464.0";
            dr16["8"] = "30743.0";
            dr16["9"] = "-0.91%";
            table.Rows.Add(dr16);

            return table;
        }
    }
}

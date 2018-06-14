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
            m_ColumnIndexContrast.Add("clinker_MixtureMaterialsOutput", "A2");
            m_ColumnIndexContrast.Add("clinker_PulverizedCoalOutput", "A5");
            m_ColumnIndexContrast.Add("clinker_ClinkerOutput", "A8");
            m_ColumnIndexContrast.Add("MixtureMaterialsOutput_Checked", "A1");
            m_ColumnIndexContrast.Add("PulverizedCoalOutput_Checked", "A4");
            m_ColumnIndexContrast.Add("ClinkerOutput_Checked", "A7");
            m_ColumnIndexContrast.Add("MixtureMaterialsOutput_Ratio", "A3");
            m_ColumnIndexContrast.Add("PulverizedCoalOutput_Ratio", "A6");
            m_ColumnIndexContrast.Add("ClinkerOutput_Ratio", "A9");

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
                    if (m_ResultTable.Rows[i][("A" + (3 * j + 1).ToString()).ToString()] != DBNull.Value && m_ResultTable.Rows[i][("A" + (3 * j + 2).ToString()).ToString()] != DBNull.Value)
                    {
                        m_ResultTable.Rows[i][("A" + (3 * j + 3).ToString()).ToString()] = (100 * ((decimal)m_ResultTable.Rows[i][("A" + (3 * j + 1).ToString()).ToString()] - (decimal)m_ResultTable.Rows[i][("A" + (3 * j + 2).ToString()).ToString()]) / (decimal)m_ResultTable.Rows[i][("A" + (3 * j + 1).ToString()).ToString()]).ToString("00.00") + "%";
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
                                group by D.LevelCode, E.Name, F.Name, D.Name, B.VariableId
                                order by D.LevelCode, E.Name, D.Name, B.VariableId";
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
            DataColumn dc3 = new DataColumn("A1", typeof(decimal));
            DataColumn dc4 = new DataColumn("A4", typeof(decimal));
            DataColumn dc5 = new DataColumn("A7", typeof(decimal));//盘库
            DataColumn dc6 = new DataColumn("A2", typeof(decimal));
            DataColumn dc7 = new DataColumn("A5", typeof(decimal));
            DataColumn dc8 = new DataColumn("A8", typeof(decimal));//DCS
            DataColumn dc9 = new DataColumn("A3", typeof(string));
            DataColumn dc10 = new DataColumn("A6", typeof(string));
            DataColumn dc11 = new DataColumn("A9", typeof(string));//增减比例
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
            DataColumn dc3 = new DataColumn("A1", Type.GetType("System.String"));
            DataColumn dc4 = new DataColumn("A4", Type.GetType("System.String"));
            DataColumn dc5 = new DataColumn("A7", Type.GetType("System.String"));//盘库
            DataColumn dc6 = new DataColumn("A2", Type.GetType("System.String"));
            DataColumn dc7 = new DataColumn("A5", Type.GetType("System.String"));
            DataColumn dc8 = new DataColumn("A8", Type.GetType("System.String"));//DCS
            DataColumn dc9 = new DataColumn("A3", Type.GetType("System.String"));
            DataColumn dc10 = new DataColumn("A6", Type.GetType("System.String"));
            DataColumn dc11 = new DataColumn("A9", Type.GetType("System.String"));//增减比例
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
            dr1["A1"] = wareHouseTable.Rows[0]["RawMaterial"];//生料
            dr1["A4"] = wareHouseTable.Rows[0]["Coal"];//煤粉
            dr1["A7"] = wareHouseTable.Rows[0]["Clinker"]; ;//熟料
            dr1["A2"] = Convert.ToDecimal(newdt.Rows[40]["TotalProduction"]).ToString("0.00");//生料
            dr1["A5"] = Convert.ToDecimal(newdt.Rows[39]["TotalProduction"]).ToString("0.00");//煤粉
            dr1["A8"] = Convert.ToDecimal(newdt.Rows[41]["TotalProduction"]).ToString("0.00");//熟料
            dr1["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[0]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[40]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[40]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr1["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[0]["Coal"]) - Convert.ToDecimal(newdt.Rows[39]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[39]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr1["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[0]["Clinker"]) - Convert.ToDecimal(newdt.Rows[41]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[41]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr1);

            DataRow dr2 = m_resultTable.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["A1"] = wareHouseTable.Rows[1]["RawMaterial"];//生料
            dr2["A4"] = wareHouseTable.Rows[1]["Coal"];//煤粉
            dr2["A7"] = wareHouseTable.Rows[1]["Clinker"]; ;//熟料
            dr2["A2"] = Convert.ToDecimal(newdt.Rows[31]["TotalProduction"]).ToString("0.00");//生料
            dr2["A5"] = Convert.ToDecimal(newdt.Rows[30]["TotalProduction"]).ToString("0.00");//煤粉
            dr2["A8"] = Convert.ToDecimal(newdt.Rows[32]["TotalProduction"]).ToString("0.00");//熟料
            dr2["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[1]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[31]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[31]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr2["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[1]["Coal"]) - Convert.ToDecimal(newdt.Rows[30]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[30]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr2["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[1]["Clinker"]) - Convert.ToDecimal(newdt.Rows[32]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[32]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr2);

            DataRow dr3 = m_resultTable.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["A1"] = wareHouseTable.Rows[2]["RawMaterial"];//生料
            dr3["A4"] = wareHouseTable.Rows[2]["Coal"];//煤粉
            dr3["A7"] = wareHouseTable.Rows[2]["Clinker"]; ;//熟料
            dr3["A2"] = Convert.ToDecimal(newdt.Rows[34]["TotalProduction"]).ToString("0.00");//生料
            dr3["A5"] = Convert.ToDecimal(newdt.Rows[33]["TotalProduction"]).ToString("0.00");//煤粉
            dr3["A8"] = Convert.ToDecimal(newdt.Rows[35]["TotalProduction"]).ToString("0.00");//熟料
            dr3["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[2]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[34]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[34]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr3["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[2]["Coal"]) - Convert.ToDecimal(newdt.Rows[33]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[33]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr3["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[2]["Clinker"]) - Convert.ToDecimal(newdt.Rows[35]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[35]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr3);

            DataRow dr4 = m_resultTable.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["A1"] = wareHouseTable.Rows[3]["RawMaterial"];//生料
            dr4["A4"] = wareHouseTable.Rows[3]["Coal"];//煤粉
            dr4["A7"] = wareHouseTable.Rows[3]["Clinker"]; ;//熟料
            dr4["A2"] = Convert.ToDecimal(newdt.Rows[37]["TotalProduction"]).ToString("0.00");//生料
            dr4["A5"] = Convert.ToDecimal(newdt.Rows[36]["TotalProduction"]).ToString("0.00");//煤粉
            dr4["A8"] = Convert.ToDecimal(newdt.Rows[38]["TotalProduction"]).ToString("0.00");//熟料
            dr4["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[3]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[37]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[37]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr4["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[3]["Coal"]) - Convert.ToDecimal(newdt.Rows[36]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[36]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr4["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[3]["Clinker"]) - Convert.ToDecimal(newdt.Rows[38]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[38]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr4);

            DataRow dr5 = m_resultTable.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["A1"] = wareHouseTable.Rows[4]["RawMaterial"];//生料
            dr5["A4"] = wareHouseTable.Rows[4]["Coal"];//煤粉
            dr5["A7"] = wareHouseTable.Rows[4]["Clinker"]; ;//熟料
            dr5["A2"] = Convert.ToDecimal(newdt.Rows[10]["TotalProduction"]).ToString("0.00");//生料
            dr5["A5"] = Convert.ToDecimal(newdt.Rows[9]["TotalProduction"]).ToString("0.00");//煤粉
            dr5["A8"] = Convert.ToDecimal(newdt.Rows[11]["TotalProduction"]).ToString("0.00");//熟料
            dr5["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[4]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[10]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[10]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr5["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[4]["Coal"]) - Convert.ToDecimal(newdt.Rows[9]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[9]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr5["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[4]["Clinker"]) - Convert.ToDecimal(newdt.Rows[11]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[11]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr5);

            DataRow dr6 = m_resultTable.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["A1"] = wareHouseTable.Rows[5]["RawMaterial"];//生料
            dr6["A4"] = wareHouseTable.Rows[5]["Coal"];//煤粉
            dr6["A7"] = wareHouseTable.Rows[5]["Clinker"]; ;//熟料
            dr6["A2"] = Convert.ToDecimal(newdt.Rows[13]["TotalProduction"]).ToString("0.00");//生料
            dr6["A5"] = Convert.ToDecimal(newdt.Rows[12]["TotalProduction"]).ToString("0.00");//煤粉
            dr6["A8"] = Convert.ToDecimal(newdt.Rows[14]["TotalProduction"]).ToString("0.00");//熟料
            dr6["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[5]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[13]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[13]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr6["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[5]["Coal"]) - Convert.ToDecimal(newdt.Rows[12]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[12]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr6["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[5]["Clinker"]) - Convert.ToDecimal(newdt.Rows[14]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[14]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr6);

            DataRow dr7 = m_resultTable.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["A1"] = wareHouseTable.Rows[6]["RawMaterial"];//生料
            dr7["A4"] = wareHouseTable.Rows[6]["Coal"];//煤粉
            dr7["A7"] = wareHouseTable.Rows[6]["Clinker"]; ;//熟料
            dr7["A2"] = Convert.ToDecimal(newdt.Rows[16]["TotalProduction"]).ToString("0.00");//生料
            dr7["A5"] = Convert.ToDecimal(newdt.Rows[15]["TotalProduction"]).ToString("0.00");//煤粉
            dr7["A8"] = Convert.ToDecimal(newdt.Rows[17]["TotalProduction"]).ToString("0.00");//熟料
            dr7["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[6]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[16]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[16]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr7["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[6]["Coal"]) - Convert.ToDecimal(newdt.Rows[15]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[15]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr7["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[6]["Clinker"]) - Convert.ToDecimal(newdt.Rows[17]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[17]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr7);

            DataRow dr8 = m_resultTable.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["A1"] = wareHouseTable.Rows[7]["RawMaterial"];//生料
            dr8["A4"] = wareHouseTable.Rows[7]["Coal"];//煤粉
            dr8["A7"] = wareHouseTable.Rows[7]["Clinker"]; ;//熟料
            dr8["A2"] = Convert.ToDecimal(newdt.Rows[19]["TotalProduction"]).ToString("0.00");//生料
            dr8["A5"] = Convert.ToDecimal(newdt.Rows[18]["TotalProduction"]).ToString("0.00");//煤粉
            dr8["A8"] = Convert.ToDecimal(newdt.Rows[20]["TotalProduction"]).ToString("0.00");//熟料
            dr8["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[7]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[19]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[19]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr8["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[7]["Coal"]) - Convert.ToDecimal(newdt.Rows[18]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[18]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr8["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[7]["Clinker"]) - Convert.ToDecimal(newdt.Rows[20]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[20]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr8);

            DataRow dr9 = m_resultTable.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["A1"] = wareHouseTable.Rows[8]["RawMaterial"];//生料
            dr9["A4"] = wareHouseTable.Rows[8]["Coal"];//煤粉
            dr9["A7"] = wareHouseTable.Rows[8]["Clinker"]; ;//熟料
            dr9["A2"] = Convert.ToDecimal(newdt.Rows[43]["TotalProduction"]).ToString("0.00");//生料
            dr9["A5"] = Convert.ToDecimal(newdt.Rows[42]["TotalProduction"]).ToString("0.00");//煤粉
            dr9["A8"] = Convert.ToDecimal(newdt.Rows[44]["TotalProduction"]).ToString("0.00");//熟料
            dr9["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[8]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[43]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[43]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr9["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[8]["Coal"]) - Convert.ToDecimal(newdt.Rows[32]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[42]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr9["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[8]["Clinker"]) - Convert.ToDecimal(newdt.Rows[44]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[44]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr9);

            DataRow dr10 = m_resultTable.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["A1"] = wareHouseTable.Rows[9]["RawMaterial"];//生料
            dr10["A4"] = wareHouseTable.Rows[9]["Coal"];//煤粉
            dr10["A7"] = wareHouseTable.Rows[9]["Clinker"]; ;//熟料
            dr10["A2"] = Convert.ToDecimal(newdt.Rows[46]["TotalProduction"]).ToString("0.00");//生料
            dr10["A5"] = Convert.ToDecimal(newdt.Rows[45]["TotalProduction"]).ToString("0.00");//煤粉
            dr10["A8"] = Convert.ToDecimal(newdt.Rows[47]["TotalProduction"]).ToString("0.00");//熟料
            dr10["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[9]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[46]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[46]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr10["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[9]["Coal"]) - Convert.ToDecimal(newdt.Rows[45]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[45]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr10["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[9]["Clinker"]) - Convert.ToDecimal(newdt.Rows[47]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[47]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr10);

            DataRow dr11 = m_resultTable.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["A1"] = wareHouseTable.Rows[10]["RawMaterial"];//生料
            dr11["A4"] = wareHouseTable.Rows[10]["Coal"];//煤粉
            dr11["A7"] = wareHouseTable.Rows[10]["Clinker"]; ;//熟料
            dr11["A2"] = Convert.ToDecimal(newdt.Rows[7]["TotalProduction"]).ToString("0.00");//生料
            dr11["A5"] = Convert.ToDecimal(newdt.Rows[6]["TotalProduction"]).ToString("0.00");//煤粉
            dr11["A8"] = Convert.ToDecimal(newdt.Rows[8]["TotalProduction"]).ToString("0.00");//熟料
            dr11["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[10]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[7]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[7]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr11["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[10]["Coal"]) - Convert.ToDecimal(newdt.Rows[6]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[6]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr11["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[10]["Clinker"]) - Convert.ToDecimal(newdt.Rows[8]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[8]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr11);

            DataRow dr12 = m_resultTable.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["A1"] = wareHouseTable.Rows[11]["RawMaterial"];//生料
            dr12["A4"] = wareHouseTable.Rows[11]["Coal"];//煤粉
            dr12["A7"] = wareHouseTable.Rows[11]["Clinker"]; ;//熟料
            dr12["A2"] = Convert.ToDecimal(newdt.Rows[22]["TotalProduction"]).ToString("0.00");//生料
            dr12["A5"] = Convert.ToDecimal(newdt.Rows[21]["TotalProduction"]).ToString("0.00");//煤粉
            dr12["A8"] = Convert.ToDecimal(newdt.Rows[23]["TotalProduction"]).ToString("0.00");//熟料
            dr12["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[11]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[22]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[22]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr12["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[11]["Coal"]) - Convert.ToDecimal(newdt.Rows[21]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[21]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr12["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[11]["Clinker"]) - Convert.ToDecimal(newdt.Rows[23]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[23]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr12);

            DataRow dr13 = m_resultTable.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["A1"] = wareHouseTable.Rows[12]["RawMaterial"];//生料
            dr13["A4"] = wareHouseTable.Rows[12]["Coal"];//煤粉
            dr13["A7"] = wareHouseTable.Rows[12]["Clinker"]; ;//熟料
            dr13["A2"] = Convert.ToDecimal(newdt.Rows[25]["TotalProduction"]).ToString("0.00");//生料
            dr13["A5"] = Convert.ToDecimal(newdt.Rows[24]["TotalProduction"]).ToString("0.00");//煤粉
            dr13["A8"] = Convert.ToDecimal(newdt.Rows[26]["TotalProduction"]).ToString("0.00");//熟料
            dr13["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[12]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[25]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[25]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr13["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[12]["Coal"]) - Convert.ToDecimal(newdt.Rows[24]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[24]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr13["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[12]["Clinker"]) - Convert.ToDecimal(newdt.Rows[26]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[26]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr13);

            DataRow dr14 = m_resultTable.NewRow();
            dr14["CompanyName"] = "乌海水泥";
            dr14["ProductionLine"] = "1#";
            dr14["A1"] = wareHouseTable.Rows[13]["RawMaterial"];//生料
            dr14["A4"] = wareHouseTable.Rows[13]["Coal"];//煤粉
            dr14["A7"] = wareHouseTable.Rows[13]["Clinker"]; ;//熟料
            dr14["A2"] = Convert.ToDecimal(newdt.Rows[28]["TotalProduction"]).ToString("0.00");//生料
            dr14["A5"] = Convert.ToDecimal(newdt.Rows[27]["TotalProduction"]).ToString("0.00");//煤粉
            dr14["A8"] = Convert.ToDecimal(newdt.Rows[29]["TotalProduction"]).ToString("0.00");//熟料
            dr14["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[13]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[28]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[28]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr14["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[13]["Coal"]) - Convert.ToDecimal(newdt.Rows[27]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[27]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr14["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[13]["Clinker"]) - Convert.ToDecimal(newdt.Rows[29]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[29]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr14);

            DataRow dr15 = m_resultTable.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["A1"] = wareHouseTable.Rows[14]["RawMaterial"];//生料
            dr15["A4"] = wareHouseTable.Rows[14]["Coal"];//煤粉
            dr15["A7"] = wareHouseTable.Rows[14]["Clinker"]; ;//熟料
            dr15["A2"] = Convert.ToDecimal(newdt.Rows[1]["TotalProduction"]).ToString("0.00");//生料
            dr15["A5"] = Convert.ToDecimal(newdt.Rows[0]["TotalProduction"]).ToString("0.00"); //煤粉
            dr15["A8"] = Convert.ToDecimal(newdt.Rows[2]["TotalProduction"]).ToString("0.00");//熟料
            dr15["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[14]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[1]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[0]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr15["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[14]["Coal"]) - Convert.ToDecimal(newdt.Rows[0]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[1]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr15["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[14]["Clinker"]) - Convert.ToDecimal(newdt.Rows[2]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[2]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr15);

            DataRow dr16 = m_resultTable.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["A1"] = wareHouseTable.Rows[15]["RawMaterial"];//生料
            dr16["A4"] = wareHouseTable.Rows[15]["Coal"];//煤粉
            dr16["A7"] = wareHouseTable.Rows[15]["Clinker"]; ;//熟料
            dr16["A2"] = Convert.ToDecimal(newdt.Rows[4]["TotalProduction"]).ToString("0.00");//生料
            dr16["A5"] = Convert.ToDecimal(newdt.Rows[3]["TotalProduction"]).ToString("0.00");//煤粉
            dr16["A8"] = Convert.ToDecimal(newdt.Rows[5]["TotalProduction"]).ToString("0.00");//熟料
            dr16["A3"] = ((Convert.ToDecimal(wareHouseTable.Rows[15]["RawMaterial"]) - Convert.ToDecimal(newdt.Rows[4]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[4]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr16["A6"] = ((Convert.ToDecimal(wareHouseTable.Rows[15]["Coal"]) - Convert.ToDecimal(newdt.Rows[3]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[3]["TotalProduction"]) * 100).ToString("0.00") + "%";
            dr16["A9"] = ((Convert.ToDecimal(wareHouseTable.Rows[15]["Clinker"]) - Convert.ToDecimal(newdt.Rows[5]["TotalProduction"])) / Convert.ToDecimal(newdt.Rows[5]["TotalProduction"]) * 100).ToString("0.00") + "%";
            m_resultTable.Rows.Add(dr16);
           
            return m_resultTable;
        }

        public static void ExportExcelFile(string myFileType, string myFileName, string myData)
        {
            if (myFileType == "xls")
            {
                UpDownLoadFiles.DownloadFile.ExportExcelFile(myFileName, myData);
            }
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
            dr1["ProductionLine"] = "1#";
            dr1["A1"] = "94328.1";
            dr1["A2"] = "130963.0";
            dr1["A3"] = "-27.97%";
            dr1["A4"] = "10134.8";
            dr1["A5"] = "8461.5";
            dr1["A6"] = "19.78%";
            dr1["A7"] = "59701.3";
            dr1["A8"] = "92053.8";
            dr1["A9"] = "-35.15%";
            table.Rows.Add(dr1);

            DataRow dr2 = table.NewRow();
            dr2["CompanyName"] = "银川水泥";
            dr2["ProductionLine"] = "2#";
            dr2["A1"] = "98860.9";
            dr2["A2"] = "90989.0";
            dr2["A3"] = "8.65%";
            dr2["A4"] = "9682.6";
            dr2["A5"] = "11993.9";
            dr2["A6"] = "-19.27%";
            dr2["A7"] = "62570.2";
            dr2["A8"] = "13551.0";
            dr2["A9"] = "361.74%";
            table.Rows.Add(dr2);

            DataRow dr3 = table.NewRow();
            dr3["CompanyName"] = "银川水泥";
            dr3["ProductionLine"] = "3#";
            dr3["A1"] = "99052.4";
            dr3["A2"] = "88761.0";
            dr3["A3"] = "11.59%";
            dr3["A4"] = "9712.0";
            dr3["A5"] = "11812.1";
            dr3["A6"] = "-17.78%";
            dr3["A7"] = "62691.4";
            dr3["A8"] = "63235.8";
            dr3["A9"] = "-0.86%";
            table.Rows.Add(dr3);

            DataRow dr4 = table.NewRow();
            dr4["CompanyName"] = "银川水泥";
            dr4["ProductionLine"] = "4#";
            dr4["A1"] = "199939.5";
            dr4["A2"] = "165138.0";
            dr4["A3"] = "21.07%";
            dr4["A4"] = "18221.6";
            dr4["A5"] = "22806.7";
            dr4["A6"] = "-20.10%";
            dr4["A7"] = "126544.0";
            dr4["A8"] = "122016.0";
            dr4["A9"] = "3.71%";
            table.Rows.Add(dr4);

            DataRow dr5 = table.NewRow();
            dr5["CompanyName"] = "青铜峡水泥";
            dr5["ProductionLine"] = "2#";
            dr5["A1"] = "168264.0";
            dr5["A2"] = "158161.0";
            dr5["A3"] = "6.39%";
            dr5["A4"] = "10742.0";
            dr5["A5"] = "12099.6";
            dr5["A6"] = "-11.22%";
            dr5["A7"] = "70775.0";
            dr5["A8"] = "70950.6";
            dr5["A9"] = "-0.25%";
            table.Rows.Add(dr5);

            DataRow dr6 = table.NewRow();
            dr6["CompanyName"] = "青铜峡水泥";
            dr6["ProductionLine"] = "3#";
            dr6["A1"] = "51937.0";
            dr6["A2"] = "45726.6";
            dr6["A3"] = "13.58%";
            dr6["A4"] = "10649.0";
            dr6["A5"] = "11079.1";
            dr6["A6"] = "-3.88%";
            dr6["A7"] = "71167.0";
            dr6["A8"] = "68715.0";
            dr6["A9"] = "3.57%";
            table.Rows.Add(dr6);

            DataRow dr7 = table.NewRow();
            dr7["CompanyName"] = "青铜峡水泥";
            dr7["ProductionLine"] = "4#";
            dr7["A1"] = "117833.0";
            dr7["A2"] = "117619.3";
            dr7["A3"] = "0.18%";
            dr7["A4"] = "11774.0";
            dr7["A5"] = "13667.1";
            dr7["A6"] = "-13.85%";
            dr7["A7"] = "75889.0";
            dr7["A8"] = "74736.9";
            dr7["A9"] = "1.54%";
            table.Rows.Add(dr7);

            DataRow dr8 = table.NewRow();
            dr8["CompanyName"] = "青铜峡水泥";
            dr8["ProductionLine"] = "5#";
            dr8["A1"] = "115099.0";
            dr8["A2"] = "113801.0";
            dr8["A3"] = "1.14%";
            dr8["A4"] = "11122.0";
            dr8["A5"] = "11455.0";
            dr8["A6"] = "-2.91%";
            dr8["A7"] = "71280.0";
            dr8["A8"] = "71250.3";
            dr8["A9"] = "0.04%";
            table.Rows.Add(dr8);

            DataRow dr9 = table.NewRow();
            dr9["CompanyName"] = "中宁水泥";
            dr9["ProductionLine"] = "1#";
            dr9["A1"] = "49367.0";
            dr9["A2"] = "38311.5";
            dr9["A3"] = "28.86%";
            dr9["A4"] = "4537.6";
            dr9["A5"] = "5451.4";
            dr9["A6"] = "-16.21%";
            dr9["A7"] = "32266.0";
            dr9["A8"] = "26222.23";
            dr9["A9"] = "23.05%";
            table.Rows.Add(dr9);

            DataRow dr10 = table.NewRow();
            dr10["CompanyName"] = "中宁水泥";
            dr10["ProductionLine"] = "2#";
            dr10["A1"] = "64448.0";
            dr10["A2"] = "54200.7";
            dr10["A3"] = "18.91%";
            dr10["A4"] = "5973.9";
            dr10["A5"] = "7679.8";
            dr10["A6"] = "-22.21%";
            dr10["A7"] = "42123.0";
            dr10["A8"] = "37155.36";
            dr10["A9"] = "13.37%";
            table.Rows.Add(dr10);

            DataRow dr11 = table.NewRow();
            dr11["CompanyName"] = "六盘山水泥";
            dr11["ProductionLine"] = "1#";
            dr11["A1"] = "48458.6";
            dr11["A2"] = "45490.5";
            dr11["A3"] = "6.52%";
            dr11["A4"] = "5130.0";
            dr11["A5"] = "4380.9";
            dr11["A6"] = "17.10%";
            dr11["A7"] = "30670.0";
            dr11["A8"] = "29540.0";
            dr11["A9"] = "3.83%";
            table.Rows.Add(dr11);

            DataRow dr12 = table.NewRow();
            dr12["CompanyName"] = "天水水泥";
            dr12["ProductionLine"] = "1#";
            dr12["A1"] = "63184.1";
            dr12["A2"] = "47691.0";
            dr12["A3"] = "32.49%";
            dr12["A4"] = "6982.4";
            dr12["A5"] = "6398.5";
            dr12["A6"] = "9.13%";
            dr12["A7"] = "40500.0";
            dr12["A8"] = "40985.0";
            dr12["A9"] = "-1.18%";
            table.Rows.Add(dr12);

            DataRow dr13 = table.NewRow();
            dr13["CompanyName"] = "天水水泥";
            dr13["ProductionLine"] = "2#";
            dr13["A1"] = "124210.0";
            dr13["A2"] = "113292.0";
            dr13["A3"] = "9.64%";
            dr13["A4"] = "13729.0";
            dr13["A5"] = "11113.0";
            dr13["A6"] = "23.54%";
            dr13["A7"] = "79624.4";
            dr13["A8"] = "82183.1";
            dr13["A9"] = "-3.11%";
            table.Rows.Add(dr13);

            DataRow dr14 = table.NewRow();
            dr14["CompanyName"] = "乌海赛马";
            dr14["ProductionLine"] = "1#";
            dr14["A1"] = "65091.2";
            dr14["A2"] = "56388.0";
            dr14["A3"] = "15.43%";
            dr14["A4"] = "8815.6";
            dr14["A5"] = "5541.6";
            dr14["A6"] = "59.08%";
            dr14["A7"] = "41210.6";
            dr14["A8"] = "44865.4";
            dr14["A9"] = "-8.15%";
            table.Rows.Add(dr14);

            DataRow dr15 = table.NewRow();
            dr15["CompanyName"] = "白银水泥";
            dr15["ProductionLine"] = "1#";
            dr15["A1"] = "155150.4";
            dr15["A2"] = "136368.3";
            dr15["A3"] = "13.77%";
            dr15["A4"] = "15236.0";
            dr15["A5"] = "18797.2";
            dr15["A6"] = "-18.95%";
            dr15["A7"] = "100097.0";
            dr15["A8"] = "92676.6";
            dr15["A9"] = "8.01%";
            table.Rows.Add(dr15);

            DataRow dr16 = table.NewRow();
            dr16["CompanyName"] = "喀喇沁水泥";
            dr16["ProductionLine"] = "1#";
            dr16["A1"] = "46305.0";
            dr16["A2"] = "35163.0";
            dr16["A3"] = "31.69%";
            dr16["A4"] = "5716.0";
            dr16["A5"] = "5946.0";
            dr16["A6"] = "-3.87%";
            dr16["A7"] = "60464.0";
            dr16["A8"] = "30743.0";
            dr16["A9"] = "-0.91%";
            table.Rows.Add(dr16);

            return table;
        }
    }
}

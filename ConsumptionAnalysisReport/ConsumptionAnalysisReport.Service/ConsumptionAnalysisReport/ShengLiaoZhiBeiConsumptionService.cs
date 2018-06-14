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
    public class ShengLiaoZhiBeiConsumptionService
    {
        public static DataTable GetRowMaterialProcessElectricityUsageByMonth(string searchDate, string[] myOganizationIds)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string m_VariableId = "rawMaterialsPreparation_ElectricityQuantity";
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
            string queryString = @"Select M.Name as CompanyName
                                        ,Z.OrganizationID as OrganizationId
                                        , case when Z.OrganizationID = M.OrganizationID then '合计' else replace(replace(replace(replace(N.Name,'号','#'),'窑',''),'熟料',''),'线','') end as ProductLine
                                        ,Z.runLevel
                                        , Z.PeakB
                                        , Z.ValleyB
                                        , Z.FlatB
                                        , Z.TotalPeakValleyFlatB
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.PeakB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as PeakBproportion
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.ValleyB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as ValleyBproportion
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.FlatB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as FlatBproportion                                       
                                        ,0 as runLevelChangedCount
                                        from system_Organization M, system_Organization N
                                        ,(Select B.OrganizationID as OrganizationID
	                                          ,D.LevelType
                                              ,sum(B.PeakB) as PeakB
                                              ,sum(B.ValleyB) as ValleyB
                                              ,sum(B.FlatB) as FlatB
	                                          ,sum(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB
                                              ,0.00 as runLevel
                                        from tz_Balance A, balance_Energy B, system_Organization C, system_Organization D
                                        where A.StaticsCycle = 'day'
                                        and A.TimeStamp like '{0}%'
                                        and A.BalanceId = B.KeyId
                                        and B.ValueType = 'ElectricityQuantity'
                                        and B.VariableId='{2}'
                                        and B.OrganizationID = D.OrganizationID
                                        and C.OrganizationID in ({1})
                                        and D.levelCode like C.LevelCode + '%'
                                        and D.LevelType = 'ProductionLine'
                                        group by B.OrganizationID, D.LevelType, B.VariableID
                                        union all
                                        Select E.OrganizationID as OrganizationID
	                                          ,E.LevelType
                                              ,sum(B.PeakB) as PeakB
                                              ,sum(B.ValleyB) as ValleyB
                                              ,sum(B.FlatB) as FlatB
	                                          ,sum(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB
                                              ,null as runLevel
                                        from tz_Balance A, balance_Energy B, system_Organization C, system_Organization D, system_Organization E
                                        where A.StaticsCycle = 'day'
                                        and A.TimeStamp like '{0}%'
                                        and A.BalanceId = B.KeyId
                                        and B.ValueType = 'ElectricityQuantity'
                                        and B.VariableId='{2}'
                                        and B.OrganizationID = D.OrganizationID
                                        and C.OrganizationID in ({1})
                                        and D.levelCode like C.LevelCode + '%'
                                        and D.LevelType = 'ProductionLine'
                                        and charindex(E.LevelCode, D.LevelCode) > 0
                                        and E.LevelType = 'Company'
                                        group by E.OrganizationID, E.LevelType) Z
                                        where N.OrganizationID = Z.OrganizationID
                                        and charindex(M.LevelCode, N.LevelCode) > 0
                                        and M.LevelType = 'Company'
                                        order by M.LevelCode, ProductLine";
            DataTable table = dataFactory.Query(string.Format(queryString, searchDate, m_OrganizationCondition, m_VariableId));
            GetEquipmentRunRate(ref table, searchDate, m_OrganizationCondition, dataFactory);
            return table;
        }
        public static void GetEquipmentRunRate(ref DataTable myResultTable, string searchDate, string myOrganizationCondition, ISqlServerDataFactory myDataFactory) 
        {
            Dictionary<string, string> m_EquipmentContrast = new Dictionary<string,string>();
            List<string> m_OrganzationIdDic = new List<string>();
            string m_EquipmentCommonId = "RawMaterialsGrind";
            string queryString = @"SELECT A.EquipmentId
                                          ,A.EquipmentName
                                          ,A.EquipmentCommonId
                                          ,A.VariableId
                                          ,A.Specifications
	                                      ,C.OrganizationID as OrganizationId
                                          ,A.ProductionLineId
                                          ,A.Enabled
                                      FROM equipment_EquipmentDetail A, system_Organization B, system_Organization C
                                      where A.Enabled = 1
                                      and A.EquipmentCommonId = '{1}'
                                      and A.OrganizationId = B.OrganizationId
                                      and C.OrganizationID in ({0})
									  and B.LevelCode like C.LevelCode + '%'
                                      and B.LevelType = 'Factory'
                                      order by A.OrganizationId, A.DisplayIndex";
            DataTable table = myDataFactory.Query(string.Format(queryString, myOrganizationCondition, m_EquipmentCommonId));
            if (table != null)
            {
                for (int i = 0; i < table.Rows.Count; i++)    //统计都计算哪些公司
                {
                    string m_CompanyOrganiazationIdTemp = table.Rows[i]["OrganizationId"].ToString();        //增加公司列表
                    if (!m_OrganzationIdDic.Contains(m_CompanyOrganiazationIdTemp))
                    {
                        m_OrganzationIdDic.Add(m_CompanyOrganiazationIdTemp);
                    }

                    string m_EquipmentIdTemp = table.Rows[i]["EquipmentId"].ToString();
                    string m_ProductionLineIdTemp = table.Rows[i]["ProductionLineId"].ToString(); 
                    if (!m_EquipmentContrast.ContainsKey(m_EquipmentIdTemp))
                    {
                        DataRow[] m_ProductionLineDataRowTemp = myResultTable.Select(string.Format("OrganizationID = '{0}'", m_ProductionLineIdTemp));
                        if (m_ProductionLineDataRowTemp.Length > 0)
                        {
                            m_EquipmentContrast.Add(m_EquipmentIdTemp, m_ProductionLineIdTemp);       //当基础表和组织机构表都有该组织机构的时候才进行计算
                        }
                    }
                }
                for (int i = 0; i < m_OrganzationIdDic.Count; i++)
                {
                    string m_StartDate = searchDate + "-01";
                    string m_EndDate = DateTime.Parse(m_StartDate).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                    DataTable m_RunRateResultTable = RunIndicators.EquipmentRunIndicators.GetEquipmentUtilizationByCommonId(new string[] { "运转率" }, m_EquipmentCommonId, m_OrganzationIdDic[i], m_StartDate, m_EndDate, myDataFactory);
                    if (m_RunRateResultTable != null)
                    {
                        for (int j = 0; j < m_RunRateResultTable.Rows.Count; j++)
                        {
                            string m_ResultEquipmentIdTemp = m_RunRateResultTable.Rows[j]["EquipmentId"].ToString();
                            if (m_EquipmentContrast.ContainsKey(m_ResultEquipmentIdTemp))
                            {
                                DataRow[] m_ProductionLineRows = myResultTable.Select(string.Format("OrganizationId = '{0}'", m_EquipmentContrast[m_ResultEquipmentIdTemp]));
                                for (int z = 0; z < m_ProductionLineRows.Length; z++)
                                {
                                    m_ProductionLineRows[z]["runLevel"] = (decimal)m_ProductionLineRows[z]["runLevel"] + (decimal)m_RunRateResultTable.Rows[j]["运转率"];
                                    m_ProductionLineRows[z]["runLevelChangedCount"] = (int)m_ProductionLineRows[z]["runLevelChangedCount"] + 1;
                                }
                            }
                        }
                    }
                }
                for (int i = 0; i < myResultTable.Rows.Count; i++)
                {
                    int m_RunLevelChangedCount = (int)myResultTable.Rows[i]["runLevelChangedCount"];
                    if (m_RunLevelChangedCount > 0)
                    {
                        myResultTable.Rows[i]["runLevel"] = 100 * (decimal)myResultTable.Rows[i]["runLevel"] / m_RunLevelChangedCount;
                    }

                }
            }
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

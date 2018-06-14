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
    public class ShengLiaoShuiNiConsumptionService
    {
        public static DataTable GetConsumptionAnalysisTable(string searchDate, string[] myOganizationIds)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string m_VariableId = "'rawMaterialsPreparation_ElectricityQuantity', 'cementPreparation_ElectricityQuantity'";
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
                                        ,replace(replace(replace(replace(N.Name,'号','#'),'窑',''),'熟料',''),'线','') as ProductLine
                                        , Z.PeakB
                                        , Z.ValleyB
                                        , Z.FlatB
                                        , Z.TotalPeakValleyFlatB
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.PeakB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as PeakBproportion
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.ValleyB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as ValleyBproportion
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.FlatB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as FlatBproportion
                                        ,Z.runLevel
                                        from system_Organization M, system_Organization N
                                        ,(Select E.OrganizationID as OrganizationID
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
                                        and B.VariableId in ({2})
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

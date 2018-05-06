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
    public class ShuiNiFenMoFengGuPingConsumptionService
    {
        public static DataTable GetShuiNiFenMoFengGuPingConsumption(string searchDate, string[] myOganizationIds)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string m_VariableId = "cementGrind_ElectricityQuantity";
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
                                        , case when Z.OrganizationID = M.OrganizationID then '合计' else replace(replace(replace(N.Name,'号','#'),'水泥',''),'磨','') end as ProductLine
                                        , Z.PeakB
                                        , Z.ValleyB
                                        , Z.FlatB
                                        , Z.TotalPeakValleyFlatB
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.PeakB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as PeakBproportion
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.ValleyB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as ValleyBproportion
                                        ,case when Z.TotalPeakValleyFlatB <> 0 then convert(varchar(32),convert(decimal(18,2),100 * Z.FlatB / Z.TotalPeakValleyFlatB)) + '%' else '0' end as FlatBproportion
                                        ,Z.runLevel
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
            Dictionary<string, string> m_EquipmentContrast = new Dictionary<string, string>();
            List<string> m_OrganzationIdDic = new List<string>();
            string m_EquipmentCommonId = "CementGrind";
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


        public static DataTable GetShuiNiFenMoFengGuPingConsumption(string searchDate)
        {
            string connectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory dataFactory = new SqlServerDataFactory(connectionString);
            string queryString = @"select VariableID, ProductLine, convert(bigint,PeakB) as PeakB, convert(bigint,ValleyB) as ValleyB, convert(bigint,FlatB) as FlatB , convert(bigint,TotalPeakValleyFlatB) as TotalPeakValleyFlatB
                     ,case When W.ProductLine like '1_'+'%' then '青铜峡水泥'  when W.ProductLine like '2_'+'%' then '中宁水泥' when W.ProductLine like '3_'+'%' then '天水水泥' when W.ProductLine like '4_'+'%' then '白银水泥' when W.ProductLine like '5_'+'%' then '六盘山水泥' end as CompanyName,
                     Convert(varchar(64),convert(decimal(5,2),(W.PeakB/TotalPeakValleyFlatB)*100))+'%' as PeakBproportion, Convert(varchar(64),convert(decimal(5,2),(W.ValleyB/TotalPeakValleyFlatB)*100))+'%' as ValleyBproportion, Convert(varchar(64),convert(decimal(5,2),(W.FlatB/TotalPeakValleyFlatB)*100))+'%' as FlatBproportion,0.00 As runLevel
                 from (Select   E.VariableID,SUM(E.TotalPeakValleyFlatB) AS TotalPeakValleyFlatB,SUM(E.PeakB) AS PeakB,SUM(E.ValleyB) AS ValleyB,SUM(E.FlatB) AS FlatB,
                 case E.OrganizationID when 'zc_nxjc_qtx_efc_cementmill01' then '1_1#' when 'zc_nxjc_qtx_efc_cementmill02' then '1_2#' when 'zc_nxjc_qtx_efc_cementmill03' then '1_3#' when 'zc_nxjc_qtx_tys_cementmill04' then '1_4#' when 'zc_nxjc_qtx_tys_cementmill05' then '1_5#'
												when 'zc_nxjc_znc_znf_cementmill01' then '2_1#' when 'zc_nxjc_znc_znf_cementmill02' then '2_2#' when 'zc_nxjc_znc_znf_cementmill03' then '2_3#' when 'zc_nxjc_znc_znf_cementmill04' then '2_4#'
												when 'zc_nxjc_tsc_tsf_cementmill01' then '3_1#' when 'zc_nxjc_tsc_tsf_cementmill02' then '3_2#'
												when 'zc_nxjc_byc_byf_cementmill01' then '4_1#' when 'zc_nxjc_byc_byf_cementmill02' then '4_2#'
												when 'zc_nxjc_lpsc_lpsf_cementmill01' then '5_1#' 	when 'zc_nxjc_lpsc_lpsf_cementmill02' then '5_2#'
												--else '6'
												end As ProductLine
	                                            from tz_Balance D, balance_Energy E 
		                                        where D.TimeStamp like '{0}'+'%'
		                                        and D.StaticsCycle = 'day'
		                                        and D.BalanceId = E.KeyId
		                                        and E.ValueType = 'ElectricityQuantity'
												and E.VariableId='cementGrind_ElectricityQuantity'
		                                        and  (E.OrganizationID = 'zc_nxjc_qtx_efc_cementmill02' or E.OrganizationID ='zc_nxjc_qtx_efc_cementmill01'
											         or E.OrganizationID = 'zc_nxjc_qtx_efc_cementmill03'
												or E.OrganizationID = 'zc_nxjc_qtx_tys_cementmill04'  or E.OrganizationID =  'zc_nxjc_qtx_tys_cementmill05' 
												or E.OrganizationID =  'zc_nxjc_znc_znf_cementmill01'or E.OrganizationID =  'zc_nxjc_znc_znf_cementmill02' or E.OrganizationID ='zc_nxjc_znc_znf_cementmill03' or E.OrganizationID ='zc_nxjc_znc_znf_cementmill04'
												or E.OrganizationID = 'zc_nxjc_tsc_tsf_cementmill01'or E.OrganizationID = 'zc_nxjc_tsc_tsf_cementmill02'  or E.OrganizationID =''  or E.OrganizationID =''
												or E.OrganizationID = 'zc_nxjc_byc_byf_cementmill01'  or E.OrganizationID ='zc_nxjc_byc_byf_cementmill02'  
												or E.OrganizationID = 'zc_nxjc_lpsc_lpsf_cementmill01'   or E.OrganizationID ='zc_nxjc_lpsc_lpsf_cementmill02')
		                                        group by case E.OrganizationID when 'zc_nxjc_qtx_efc_cementmill01' then '1_1#' when 'zc_nxjc_qtx_efc_cementmill02' then '1_2#' when 'zc_nxjc_qtx_efc_cementmill03' then '1_3#' when 'zc_nxjc_qtx_tys_cementmill04' then '1_4#' when 'zc_nxjc_qtx_tys_cementmill05' then '1_5#'
												when 'zc_nxjc_znc_znf_cementmill01' then '2_1#' when 'zc_nxjc_znc_znf_cementmill02' then '2_2#' when 'zc_nxjc_znc_znf_cementmill03' then '2_3#' when 'zc_nxjc_znc_znf_cementmill04' then '2_4#'
												when 'zc_nxjc_tsc_tsf_cementmill01' then '3_1#' when 'zc_nxjc_tsc_tsf_cementmill02' then '3_2#'
												when 'zc_nxjc_byc_byf_cementmill01' then '4_1#' when 'zc_nxjc_byc_byf_cementmill02' then '4_2#'
												when 'zc_nxjc_lpsc_lpsf_cementmill01' then '5_1#' 	when 'zc_nxjc_lpsc_lpsf_cementmill02' then '5_2#'
											    end
												,E.VariableID ) As W order by  ProductLine asc
									";
            DataTable table = dataFactory.Query(string.Format(queryString, searchDate));

            DataTable dtnew = new DataTable();
            dtnew = table.Clone();
            DataRow dr1 = dtnew.NewRow();
            dr1["PeakB"] = 0;
            dr1["ValleyB"] = 0;
            dr1["FlatB"] = 0;
            dr1["TotalPeakValleyFlatB"] = 0;
            DataRow dr2 = dtnew.NewRow();
            dr2["PeakB"] = 0;
            dr2["ValleyB"] = 0;
            dr2["FlatB"] = 0;
            dr2["TotalPeakValleyFlatB"] = 0;
            DataRow dr3 = dtnew.NewRow();
            dr3["PeakB"] = 0;
            dr3["ValleyB"] = 0;
            dr3["FlatB"] = 0;
            dr3["TotalPeakValleyFlatB"] = 0;

            DataRow dr4 = dtnew.NewRow();
            dr4["PeakB"] = 0;
            dr4["ValleyB"] = 0;
            dr4["FlatB"] = 0;
            dr4["TotalPeakValleyFlatB"] = 0;

            DataRow dr5 = dtnew.NewRow();
            dr5["PeakB"] = 0;
            dr5["ValleyB"] = 0;
            dr5["FlatB"] = 0;
            dr5["TotalPeakValleyFlatB"] = 0;
            int dtnewrow0Index = 0;
            int dtnewrow1Index = 0;
            int dtnewrow2Index = 0;
            int dtnewrow3Index = 0;
            int dtnewrow4Index = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {

                if (table.Rows[i]["CompanyName"].ToString() == "青铜峡水泥")
                {
                    dtnewrow0Index++;
                    dr1["productline"] = "合计";
                    dr1["TotalPeakValleyFlatB"] = (Convert.ToDecimal(dr1["TotalPeakValleyFlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["TotalPeakValleyFlatB"].ToString())).ToString();
                    dr1["PeakB"] = Convert.ToDecimal(dr1["PeakB"].ToString()) + Convert.ToDecimal(table.Rows[i]["PeakB"].ToString());
                    dr1["ValleyB"] = Convert.ToDecimal(dr1["ValleyB"].ToString()) + Convert.ToDecimal(table.Rows[i]["ValleyB"].ToString());
                    dr1["FlatB"] = Convert.ToDecimal(dr1["FlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["FlatB"].ToString());

                }
                else if (table.Rows[i]["CompanyName"].ToString() == "中宁水泥")
                {
                    dtnewrow1Index++;
                    dr2["productline"] = "合计";
                    dr2["TotalPeakValleyFlatB"] = Convert.ToDecimal(dr2["TotalPeakValleyFlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["TotalPeakValleyFlatB"].ToString());
                    dr2["PeakB"] = Convert.ToDecimal(dr2["PeakB"].ToString()) + Convert.ToDecimal(table.Rows[i]["PeakB"].ToString());
                    dr2["ValleyB"] = Convert.ToDecimal(dr2["ValleyB"].ToString()) + Convert.ToDecimal(table.Rows[i]["ValleyB"].ToString());
                    dr2["FlatB"] = Convert.ToDecimal(dr2["FlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["FlatB"].ToString());
                }
                else if (table.Rows[i]["CompanyName"].ToString() == "天水水泥")
                {
                    dtnewrow2Index++;
                    dr3["productline"] = "合计";
                    dr3["TotalPeakValleyFlatB"] = Convert.ToDecimal(dr3["TotalPeakValleyFlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["TotalPeakValleyFlatB"].ToString());
                    dr3["PeakB"] = Convert.ToDecimal(dr3["PeakB"].ToString()) + Convert.ToDecimal(table.Rows[i]["PeakB"].ToString());
                    dr3["ValleyB"] = Convert.ToDecimal(dr3["ValleyB"].ToString()) + Convert.ToDecimal(table.Rows[i]["ValleyB"].ToString());
                    dr3["FlatB"] = Convert.ToDecimal(dr3["FlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["FlatB"].ToString());

                }
                else if (table.Rows[i]["CompanyName"].ToString() == "白银水泥")
                {
                    dtnewrow3Index++;
                    dr4["productline"] = "合计";
                    dr4["TotalPeakValleyFlatB"] = Convert.ToDecimal(dr4["TotalPeakValleyFlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["TotalPeakValleyFlatB"].ToString());
                    dr4["PeakB"] = Convert.ToDecimal(dr4["PeakB"].ToString()) + Convert.ToDecimal(table.Rows[i]["PeakB"].ToString());
                    dr4["ValleyB"] = Convert.ToDecimal(dr4["ValleyB"].ToString()) + Convert.ToDecimal(table.Rows[i]["ValleyB"].ToString());
                    dr4["FlatB"] = Convert.ToDecimal(dr4["FlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["FlatB"].ToString());
                }
                else if (table.Rows[i]["CompanyName"].ToString() == "六盘山水泥")
                {
                    dtnewrow4Index++;
                    dr5["productline"] = "合计";
                    dr5["TotalPeakValleyFlatB"] = Convert.ToDecimal(dr5["TotalPeakValleyFlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["TotalPeakValleyFlatB"].ToString());
                    dr5["PeakB"] = Convert.ToDecimal(dr5["PeakB"].ToString()) + Convert.ToDecimal(table.Rows[i]["PeakB"].ToString());
                    dr5["ValleyB"] = Convert.ToDecimal(dr5["ValleyB"].ToString()) + Convert.ToDecimal(table.Rows[i]["ValleyB"].ToString());
                    dr5["FlatB"] = Convert.ToDecimal(dr5["FlatB"].ToString()) + Convert.ToDecimal(table.Rows[i]["FlatB"].ToString());
                }

            }
            if (dtnewrow0Index != 0)
            {
                dr1["PeakBproportion"] = ((Convert.ToDecimal(dr1["PeakB"].ToString()) / Convert.ToDecimal(dr1["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr1["ValleyBproportion"] = ((Convert.ToDecimal(dr1["ValleyB"].ToString()) / Convert.ToDecimal(dr1["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr1["FlatBproportion"] = ((Convert.ToDecimal(dr1["FlatB"].ToString()) / Convert.ToDecimal(dr1["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';

                dr1["CompanyName"] = "青铜峡水泥";
                dtnew.Rows.Add(dr1);
                DataRow row1 = table.NewRow();
                row1.ItemArray = dr1.ItemArray;
                table.Rows.InsertAt(row1, dtnewrow0Index);
            }
            if (dtnewrow1Index != 0)
            {
                dr2["PeakBproportion"] = ((Convert.ToDecimal(dr2["PeakB"].ToString()) / Convert.ToDecimal(dr2["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr2["ValleyBproportion"] = ((Convert.ToDecimal(dr2["ValleyB"].ToString()) / Convert.ToDecimal(dr2["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr2["FlatBproportion"] = ((Convert.ToDecimal(dr2["FlatB"].ToString()) / Convert.ToDecimal(dr2["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';

                dr2["CompanyName"] = "中宁水泥";
                dtnew.Rows.Add(dr2);
                DataRow row2 = table.NewRow();
                row2.ItemArray = dr2.ItemArray;
                table.Rows.InsertAt(row2, dtnewrow0Index + dtnewrow1Index + 1);
            }
            if (dtnewrow2Index != 0)
            {
                dr3["PeakBproportion"] = ((Convert.ToDecimal(dr3["PeakB"].ToString()) / Convert.ToDecimal(dr3["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr3["ValleyBproportion"] = ((Convert.ToDecimal(dr3["ValleyB"].ToString()) / Convert.ToDecimal(dr3["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr3["FlatBproportion"] = ((Convert.ToDecimal(dr3["FlatB"].ToString()) / Convert.ToDecimal(dr3["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';

                dr3["CompanyName"] = "天水水泥";
                dtnew.Rows.Add(dr3);
                DataRow row3 = table.NewRow();
                row3.ItemArray = dr3.ItemArray;
                table.Rows.InsertAt(row3, dtnewrow0Index + dtnewrow1Index + 1 + dtnewrow2Index + 1);
            }

            if (dtnewrow3Index != 0)
            {
                dr4["PeakBproportion"] = ((Convert.ToDecimal(dr4["PeakB"].ToString()) / Convert.ToDecimal(dr4["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr4["ValleyBproportion"] = ((Convert.ToDecimal(dr4["ValleyB"].ToString()) / Convert.ToDecimal(dr4["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr4["FlatBproportion"] = ((Convert.ToDecimal(dr4["FlatB"].ToString()) / Convert.ToDecimal(dr4["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr4["CompanyName"] = "白银水泥";
                dtnew.Rows.Add(dr4);
                DataRow row4 = table.NewRow();
                row4.ItemArray = dr4.ItemArray;
                table.Rows.InsertAt(row4, dtnewrow0Index + dtnewrow1Index + 1 + dtnewrow2Index + 1 + dtnewrow3Index + 1);
            }
            if (dtnewrow4Index != 0)
            {
                dr5["PeakBproportion"] = ((Convert.ToDecimal(dr5["PeakB"].ToString()) / Convert.ToDecimal(dr5["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr5["ValleyBproportion"] = ((Convert.ToDecimal(dr5["ValleyB"].ToString()) / Convert.ToDecimal(dr5["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr5["FlatBproportion"] = ((Convert.ToDecimal(dr5["FlatB"].ToString()) / Convert.ToDecimal(dr5["TotalPeakValleyFlatB"].ToString())) * 100).ToString("0.00") + '%';
                dr5["CompanyName"] = "六盘山水泥";
                dtnew.Rows.Add(dr5);
                DataRow row5 = table.NewRow();
                row5.ItemArray = dr5.ItemArray;
                table.Rows.InsertAt(row5, dtnewrow0Index + dtnewrow1Index + 1 + dtnewrow2Index + 1 + dtnewrow3Index + 1 + dtnewrow4Index + 1);
            }
            GetEquipmentRunRate(ref table, searchDate, dataFactory);
            return table;
        }
        public static void GetEquipmentRunRate(ref DataTable myResultTable, string searchDate, ISqlServerDataFactory myDataFactory)
        {
            Dictionary<string, string> m_OrganzationIdDic = new Dictionary<string, string>();
            //myResultTable.Columns.Add("runLevel", typeof(decimal));          //增加一列运转率
            string queryString = @"SELECT substring(A.EquipmentName,1,1) + '#' AS EquipmentName
                                          ,A.EquipmentCommonId
                                          ,A.VariableId
                                          ,A.Specifications
	                                      ,C.OrganizationID as OrganizationId
                                          ,A.ProductionLineId
	                                      ,Replace(C.Name,'公司','水泥') as OrganizationName
                                          ,A.Enabled
                                      FROM equipment_EquipmentDetail A, system_Organization B, system_Organization C
                                      where A.Enabled = 1
                                      and A.EquipmentCommonId = 'RawMaterialsGrind'
                                      and A.OrganizationId = B.OrganizationId
                                      and CHARINDEX(C.LevelCode, B.LevelCode) > 0
                                      and C.LevelType = 'Company'
                                      order by A.OrganizationId, A.DisplayIndex";
            DataTable table = myDataFactory.Query(string.Format(queryString, searchDate));
            if (table != null)
            {
                for (int i = 0; i < table.Rows.Count; i++)    //统计都计算哪些公司
                {
                    string m_OrganizationIdTemp = table.Rows[i]["OrganizationId"].ToString();
                    string m_OrganizationNameTemp = table.Rows[i]["OrganizationName"].ToString();
                    if (!m_OrganzationIdDic.ContainsKey(m_OrganizationNameTemp))
                    {
                        DataRow[] m_CompanyDataRowTemp = myResultTable.Select(string.Format("CompanyName = '{0}'", m_OrganizationNameTemp));
                        if (m_CompanyDataRowTemp.Length > 0)
                        {
                            m_OrganzationIdDic.Add(m_OrganizationNameTemp, m_OrganizationIdTemp);       //当基础表和组织机构表都有该组织机构的时候才进行计算
                        }
                    }
                }
                foreach (string myKey in m_OrganzationIdDic.Keys)
                {
                    string m_StartDate = searchDate + "-01";
                    string m_EndDate = DateTime.Parse(m_StartDate).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                    DataTable m_RunRateResultTable = RunIndicators.EquipmentRunIndicators.GetEquipmentUtilizationByCommonId(new string[] { "运转率" }, "CementGrind", m_OrganzationIdDic[myKey], m_StartDate, m_EndDate, myDataFactory);
                    if (m_RunRateResultTable != null)
                    {
                        for (int i = 0; i < m_RunRateResultTable.Rows.Count; i++)
                        {
                            string m_ProductionLine = m_RunRateResultTable.Rows[i]["EquipmentName"].ToString().Substring(0, 1) + "#";
                            DataRow[] m_ProductionLineRows = myResultTable.Select(string.Format("CompanyName = '{0}' and substring(productline,3,2) = '{1}'", myKey, m_ProductionLine));
                            for (int j = 0; j < m_ProductionLineRows.Length; j++)
                            {
                                m_ProductionLineRows[j]["runLevel"] = (decimal)m_RunRateResultTable.Rows[i]["运转率"];
                            }
                        }
                    }
                }
            }
        }
    }
}

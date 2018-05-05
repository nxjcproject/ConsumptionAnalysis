using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using ConsumptionAnalysisReport.Infrastructure.Configuration;
using SqlServerDataAdapter;
namespace ConsumptionAnalysisReport.Service.ComonFunction
{
    public class ElectricityConsumptionTemplate1
    {
        public static DataTable GetElectricityConsumption(DataTable myVariableFormulaTable, string myKeyColumn, string myStartDate, string myEndDate)
        {
            //string m_ElectricityConsumptionFlag = "_ElectricityConsumption";
            List<string> m_VariableIdList = new List<string>();
            //////////////////////构造表结构///////////////////////
            DataTable m_ElectricityConsumptionTable = new DataTable();
            int m_ValueZone = -1;
            for (int i = 0; i < myVariableFormulaTable.Columns.Count; i++)
            {
                if (myVariableFormulaTable.Columns[i].ColumnName == myKeyColumn)
                {
                    m_ValueZone = i + 1;
                    m_ElectricityConsumptionTable.Columns.Add(myVariableFormulaTable.Columns[i].ColumnName, typeof(string));
                }
                else if (m_ValueZone != -1)
                {
                    m_ElectricityConsumptionTable.Columns.Add(myVariableFormulaTable.Columns[i].ColumnName, typeof(decimal));
                }
                else
                {
                    m_ElectricityConsumptionTable.Columns.Add(myVariableFormulaTable.Columns[i].ColumnName, myVariableFormulaTable.Columns[i].DataType);
                }
            }
            if (m_ValueZone != -1)
            {
                for (int i = 0; i < myVariableFormulaTable.Rows.Count; i++)
                {
                    for (int j = m_ValueZone; j < myVariableFormulaTable.Columns.Count; j++)
                    {
                        if (myVariableFormulaTable.Rows[i][j] != DBNull.Value && myVariableFormulaTable.Rows[i][j].ToString() != "")
                        {
                            MatchCollection m_Collection = Regex.Matches(myVariableFormulaTable.Rows[i][j].ToString(), @"[\w\.]+(?:\s*\([\s\S]*?((?'op'\([\s\S]*?)*(?'-op'\)[\s\S]*?)*)*(?(op)(?!))\))?");
                            foreach (Match myItem in m_Collection)
                            {
                                if (!m_VariableIdList.Contains(myItem.Value))
                                {
                                    m_VariableIdList.Add(myItem.Value);
                                }
                            }
                        }
                    }
                }
            }

            string m_ConnectionString = ConnectionStringFactory.NXJCConnectionString;
            ISqlServerDataFactory m_DataFactory = new SqlServerDataFactory(m_ConnectionString);

            DataTable m_SourceTable = GetVariableValue(myVariableFormulaTable, m_VariableIdList, myKeyColumn, myStartDate, myEndDate, m_DataFactory);
            DataTable m_FormulaTable = GetConsumptionFormula(m_VariableIdList, m_DataFactory);
            Dictionary<string, string> m_KeyColumnName = GetProductionLineName(myVariableFormulaTable, myKeyColumn, m_DataFactory);
            DataTable m_ConsumptionValue = EnergyConsumption.EnergyConsumptionCalculate.CalculateByOrganizationId(m_SourceTable, m_FormulaTable, "ValueFormula", new string[] { "TotalPeakValleyFlatB" });
            for (int i = 0; i < myVariableFormulaTable.Rows.Count; i++)
            {
                DataRow m_NewValueRow = m_ElectricityConsumptionTable.NewRow();
                m_NewValueRow[0] = myVariableFormulaTable.Rows[i][0].ToString();   //赋值第一列的信息
                if (m_KeyColumnName.ContainsKey(myVariableFormulaTable.Rows[i][myKeyColumn].ToString()))//赋值第二列的产线名称
                {
                    m_NewValueRow[myKeyColumn] = m_KeyColumnName[myVariableFormulaTable.Rows[i][myKeyColumn].ToString()];
                }
                GetResult(ref m_NewValueRow, myVariableFormulaTable.Rows[i], myKeyColumn, m_ConsumptionValue, m_ValueZone, myVariableFormulaTable.Columns.Count, "TotalPeakValleyFlatB");
                m_ElectricityConsumptionTable.Rows.Add(m_NewValueRow);
            }

            return m_ElectricityConsumptionTable;
        }
        private static Dictionary<string, string> GetProductionLineName(DataTable myVariableFormulaTable, string myKeyColumn, ISqlServerDataFactory myDataFactory)
        {
            Dictionary<string, string> m_PruductionLineNameDic = new Dictionary<string,string>();
            string m_OganizationIds = "";
            for (int i = 0; i < myVariableFormulaTable.Rows.Count; i++)
            {
                if (i == 0)
                {
                    m_OganizationIds = "'" + myVariableFormulaTable.Rows[i][myKeyColumn].ToString() + "'";
                }
                else
                {
                    m_OganizationIds = m_OganizationIds + ",'" + myVariableFormulaTable.Rows[i][myKeyColumn].ToString() + "'";
                }
            }
            string dataSql = @"Select A.OrganizationID
                                  , substring(A.name,1,1) + '#' as ProductionLineName from system_Organization A
                                  where A.OrganizationID in ({0})";
            dataSql = string.Format(dataSql, m_OganizationIds);
            DataTable sourceData = myDataFactory.Query(dataSql);
            if (sourceData != null)
            {
                for (int i = 0; i < sourceData.Rows.Count; i++)
                {
                    if(!m_PruductionLineNameDic.ContainsKey(sourceData.Rows[i]["ProductionLineName"].ToString()))
                    {
                        m_PruductionLineNameDic.Add(sourceData.Rows[i]["OrganizationID"].ToString(), sourceData.Rows[i]["ProductionLineName"].ToString());
                    }
                }
            }
            return m_PruductionLineNameDic;
        }
        private static void GetResult(ref DataRow myValueRow, DataRow myVariableFormulaRow, string myKeyColumn, DataTable myConsumptionValue, int myValueZone, int myColumnCount, string myValueColumnName)
        {
            Dictionary<string, decimal> m_VariableValueDic = new Dictionary<string, decimal>();
            string m_OrganizationId = myVariableFormulaRow[myKeyColumn].ToString();
            DataRow[] m_ValueRows = myConsumptionValue.Select(string.Format("OrganizationID = '{0}'", m_OrganizationId));
            for (int i = 0; i < m_ValueRows.Length; i++)
            {
                if (!m_VariableValueDic.ContainsKey(m_ValueRows[i]["VariableId"].ToString()))
                {
                    m_VariableValueDic.Add(m_ValueRows[i]["VariableId"].ToString(), (decimal)m_ValueRows[i][myValueColumnName]);
                }
            }
            for (int i = myValueZone; i < myColumnCount; i++)
            {
                if (myVariableFormulaRow[i] != DBNull.Value && myVariableFormulaRow[i].ToString() != "")
                {
                    Regex reg = new Regex(@"[\w\.]+(?:\s*\([\s\S]*?((?'op'\([\s\S]*?)*(?'-op'\)[\s\S]*?)*)*(?(op)(?!))\))?");
                    string m_Formula = myVariableFormulaRow[i].ToString();
                    m_Formula = reg.Replace(m_Formula,
                       new MatchEvaluator(m =>
                       m_VariableValueDic.ContainsKey(m.Value) ? m_VariableValueDic[m.Value].ToString() : m.Value));
                    try
                    {
                        decimal m_ResultTemp = Decimal.Parse(myConsumptionValue.Compute(m_Formula, "").ToString());
                        myValueRow[i] = m_ResultTemp;
                    }
                    catch
                    {

                    }
                }
            }
        }
        private static DataTable GetVariableValue(DataTable myElectricityConsumptionTable, List<string> myVariableIdList, string myKeyColumn, string myStartDate, string myEndDate, ISqlServerDataFactory myDataFactory)
        {
            string m_OganizationIds = "";
            for (int i = 0; i < myElectricityConsumptionTable.Rows.Count; i++)
            {
                if (i == 0)
                {
                    m_OganizationIds = "'" + myElectricityConsumptionTable.Rows[i][myKeyColumn].ToString() + "'";
                }
                else
                {
                    m_OganizationIds = m_OganizationIds + ",'" + myElectricityConsumptionTable.Rows[i][myKeyColumn].ToString() + "'";
                }
            }
            string m_VariableIds = "";
            for (int i = 0; i < myVariableIdList.Count; i++)
            {
                if (i == 0)
                {
                    m_VariableIds = "'" + myVariableIdList[i] + "_ElectricityQuantity'";
                }
                else
                {
                    m_VariableIds = m_VariableIds + ",'" + myVariableIdList[i] + "_ElectricityQuantity'";
                }
            }

            string dataSql = @"select B.OrganizationID,B.VariableId, SUM(B.TotalPeakValleyFlatB) as TotalPeakValleyFlatB
                                from tz_Balance A,balance_Energy B
                                where A.BalanceId=B.KeyId
								and B.OrganizationID in ({0})
								and A.StaticsCycle = 'day'
                                and A.TimeStamp >= @StartDate
								and A.TimeStamp <= @EndDate
								and (B.VariableId in ({1}) or ValueType = 'MaterialWeight')
                                group by B.OrganizationID, B.VariableId";
            SqlParameter[] parameters = { new SqlParameter("@StartDate", myStartDate)
                                            ,new SqlParameter("@EndDate", myEndDate)};
            dataSql = string.Format(dataSql, m_OganizationIds, m_VariableIds);
            DataTable sourceData = myDataFactory.Query(dataSql, parameters);
            return sourceData;
        }
        private static DataTable GetConsumptionFormula(List<string> myVariableIdList, ISqlServerDataFactory myDataFactory)
        {
            string mySql = @"select A.OrganizationID,B.Name,B.LevelCode,B.VariableID,B.LevelType,C.ValueFormula
                                from (SELECT LevelCode FROM system_Organization WHERE LevelType = 'ProductionLine' ) M inner join system_Organization N
	                                on N.LevelCode Like M.LevelCode+'%' inner join tz_Formula A 
	                                on A.OrganizationID=N.OrganizationID inner join formula_FormulaDetail B
	                                on A.KeyID=B.KeyID and A.Type=2 left join balance_Energy_Template C 
	                                on B.VariableId+'_'+@consumptionType=C.VariableId   
                                order by OrganizationID,LevelCode";
            SqlParameter parameter = new SqlParameter("@consumptionType", "ElectricityConsumption");
            DataTable frameTable = myDataFactory.Query(mySql, parameter);
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
            //////////////////过滤多余的公式////////////////
            string m_VariableIds = "";
            for (int i = 0; i < myVariableIdList.Count; i++)
            {
                if (i == 0)
                {
                    m_VariableIds = "'" + myVariableIdList[i] + "'";
                }
                else
                {
                    m_VariableIds = m_VariableIds + ",'" + myVariableIdList[i] + "'";
                }
            }
            if (preFormula != null)
            {
                DataRow[] m_PreFormulaF = frameTable.Select(string.Format("VariableId in ({0})", m_VariableIds));
                if (m_PreFormulaF.Length > 0)
                {
                    return m_PreFormulaF.CopyToDataTable();
                }
                else
                {
                    return frameTable.Clone();
                }
            }
            else
            {
                return frameTable;
            }
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
    }
}

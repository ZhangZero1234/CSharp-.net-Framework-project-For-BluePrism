using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace CreateDataTable
{
    public class CreateData
    {
        public DataTable CreateTable(string contract_number,string contract_end_date,string contract_amount,string table_html,string sub_first_end_date,string sub_first_labor_amount,string sub_second_end_date,string sub_second_labor_amount)
        {
            DataTable dt = new DataTable();
            DataColumn contract_Number = new DataColumn("contract number", typeof(string));
            DataColumn contract_End_Date = new DataColumn("contract end date", typeof(string));
            DataColumn contract_Amount = new DataColumn("contract amount", typeof(string));
            DataColumn date_last_12 = new DataColumn("date last 12", typeof(string));
            DataColumn date_last_11 = new DataColumn("date last 11", typeof(string));
            DataColumn date_last_10 = new DataColumn("date last 10", typeof(string));
            DataColumn date_last_9 = new DataColumn("date last 9", typeof(string));
            DataColumn date_last_8 = new DataColumn("date last 8", typeof(string));
            DataColumn date_last_7 = new DataColumn("date last 7", typeof(string));
            DataColumn date_last_6 = new DataColumn("date last 6", typeof(string));
            DataColumn date_last_5 = new DataColumn("date last 5", typeof(string));
            DataColumn date_last_4 = new DataColumn("date last 4", typeof(string));
            DataColumn date_last_3 = new DataColumn("date last 3", typeof(string));
            DataColumn date_last_2 = new DataColumn("date last 2", typeof(string));
            DataColumn date_last_1 = new DataColumn("date last 1", typeof(string));
            DataColumn sub_First_End_Date = new DataColumn("sub 1 - end date", typeof(string));
            DataColumn sub_First_Labor_Amount = new DataColumn("sub 1 - labor amount", typeof(string));
            DataColumn sub_Second_End_Date = new DataColumn("sub 2 - end date", typeof(string));
            DataColumn sub_Second_Labor_Amount = new DataColumn("sub 2 - labor amount", typeof(string));
            dt.Columns.Add(contract_Number);
            dt.Columns.Add(contract_End_Date);
            dt.Columns.Add(contract_Amount);
            dt.Columns.Add(date_last_12);
            dt.Columns.Add(date_last_11);
            dt.Columns.Add(date_last_10);
            dt.Columns.Add(date_last_9);
            dt.Columns.Add(date_last_8);
            dt.Columns.Add(date_last_7);
            dt.Columns.Add(date_last_6);
            dt.Columns.Add(date_last_5);
            dt.Columns.Add(date_last_4);
            dt.Columns.Add(date_last_3);
            dt.Columns.Add(date_last_2);
            dt.Columns.Add(date_last_1);
            dt.Columns.Add(sub_First_End_Date);
            dt.Columns.Add(sub_First_Labor_Amount);
            dt.Columns.Add(sub_Second_End_Date);
            dt.Columns.Add(sub_Second_Labor_Amount);
            DataRow dr = dt.NewRow();

            dr["contract number"] = contract_number;

            //開始予定日2016年05月01日終了予定日2019年04月30日
            string contractEndDate = contract_end_date;
            contractEndDate = contractEndDate.Replace("終了予定日", "|");
            contractEndDate = contractEndDate.Split('|')[1].Trim();
            contractEndDate = contractEndDate.Replace("年", "-").Replace("月", "-").Replace("日", "");
            dr["contract end date"] = contractEndDate;


            //89,640円
            string contractAmount = contract_amount;
            contractAmount = contractAmount.Replace("円", "");
            dr["contract amount"] = contractAmount;

            //get list
            string htmlList = table_html;
            htmlList = htmlList.Replace("支払期日", "").Trim();
            htmlList = htmlList.Replace("支払金額", "").Trim();
            //ArrayList arrayList = new ArrayList();
            string[] buffer = htmlList.Split('円');
            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = buffer[i].Replace("日", "|").Split('|')[0].Replace("年", "-").Replace("月", "-");
            }
            //取后12个
            for (int j = 0; j < 12; j++)
            {
                //arrayList[j] = buffer[buffer.Length-12+j];
                dr["date last " + (12 - j)] = buffer[buffer.Length - 1 - 12 + j].Trim();
            }

            //1開始予定日2016年05月01日終了予定日2019年04月30日
            string subFirstEndDate = sub_first_end_date;
            subFirstEndDate = subFirstEndDate.Replace("終了予定日", "|");
            subFirstEndDate = subFirstEndDate.Split('|')[1].Trim();
            subFirstEndDate = subFirstEndDate.Replace("年", "-").Replace("月", "-").Replace("日", "");

            string subFirstLaborAmount = sub_first_labor_amount;
            subFirstLaborAmount = subFirstLaborAmount.Replace("円", "");
            dr["sub 1 - end date"] = subFirstEndDate;
            dr["sub 1 - labor amount"] = subFirstLaborAmount;


            //2開始予定日2016年05月01日終了予定日2019年04月30日
            string subSecondEndDate = sub_second_end_date;
            subSecondEndDate = subSecondEndDate.Replace("終了予定日", "|");
            subSecondEndDate = subSecondEndDate.Split('|')[1].Trim();
            subSecondEndDate = subSecondEndDate.Replace("年", "-").Replace("月", "-").Replace("日", "");
            string subSecondLaborAmount = sub_second_labor_amount;
            subSecondLaborAmount = subSecondLaborAmount.Replace("円", "");
            dr["sub 2 - end date"] = subSecondEndDate;
            dr["sub 2 - labor amount"] = subSecondLaborAmount;

            dt.Rows.Add(dr);
            return dt;
        }
        
    }
}

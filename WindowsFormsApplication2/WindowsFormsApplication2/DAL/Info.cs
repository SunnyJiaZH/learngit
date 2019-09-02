using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication2.DAL
{
    class Info
    {
        public string FileName  { get; set; }
        public string PublishURL { get; set; }
        public static List<Info> ToList(DataSet dataSet)
        {
            List<Info> List = new List<Info>();
            if (dataSet != null && dataSet.Tables.Count > 0)
            {
                foreach(DataRow row in dataSet.Tables[0].Rows)
                {
                    Info order = new Info();
                    if(dataSet.Tables[0].Columns.Contains("HostAssertID")&& !Convert.IsDBNull(row["HostAssertID"]))
                    {
                        order.FileName = (string)row["HostAssertID"];
                    }
                    if (dataSet.Tables[0].Columns.Contains("HostAssertID") && !Convert.IsDBNull(row["HostAssertID"]))
                    {
                        order.PublishURL = (string)row["HostAssertID"];
                    }
                    List.Add(order);
                }
            }
            return List;
        }
    }
}

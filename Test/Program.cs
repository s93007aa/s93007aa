using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace Test
{
    public class Program
    {
        static void Main(string[] args)
        {
            int BugWhere = 0;
            string company = args[0];
            string factory = args[1];
            //string company = "GRSI";
            //string factory = "RSV";

            try
            {
                string name = "ERP_WMS_Diff";
                string filePath = @"D:\ERP_WMS_inventory\excel\";
                string fileName = string.Format("{0}_{1}_{2}_{3}.xlsx", company, factory, name, DateTime.Now.ToString("yyyyMMddHHmm"));
                string filePathName = filePath + fileName;

                //刪除同檔名
                //excel.DeleteNoLockTemporaryFile(fileName);

                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件

                List<double> timeRecord = new List<double>();

                sw.Reset();//碼表歸零
                sw.Start();//碼表開始計時
                BugWhere++;//1
                var fabricData = RunFabricTotal(company, factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.Diff_Qty);
                timeRecord.Add(sw.Elapsed.TotalSeconds);

                BugWhere++;//2
                var fabricGradeData = RunFabricGrade(company, factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.Diff_Qty);
                timeRecord.Add(sw.Elapsed.TotalSeconds);

                BugWhere++;//3
                var trimData = RunTrimTotal(company, factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.Diff_Qty);
                timeRecord.Add(sw.Elapsed.TotalSeconds);

                BugWhere++;//4
                var trimGradeData = RunTrimGrade(company, factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.Diff_Qty);
                timeRecord.Add(sw.Elapsed.TotalSeconds);

                BugWhere++;//5
                var trimMPOData = RunTrimMPO(company, factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.Diff_Qty);
                timeRecord.Add(sw.Elapsed.TotalSeconds);

                BugWhere++;//6
                var ERPAData = RunErpADataSQL(company, factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.ERP_Qty);
                timeRecord.Add(sw.Elapsed.TotalSeconds);

                //////////////////////
        

                BugWhere++;//13
                if (company == "GRSI")
                {
                    var MS_OrderData = RunMS_Order(factory).OrderBy(x => x.MaterialNo).ThenBy(y => y.MS_Order);
                }
                sw.Stop();//碼錶停止

                if (!Directory.Exists(@"D:\ERP_WMS_inventory\countTime"))
                {
                    Directory.CreateDirectory(@"D:\ERP_WMS_inventory\countTime");
                }
                StreamWriter countTime = new StreamWriter(string.Format(@"D:\ERP_WMS_inventory\countTime\countTime_{0}.txt", DateTime.Now.ToString("yyyyMMdd")), true);//紀錄費時
                countTime.WriteLine(company + "_" + factory + "_" + DateTime.Now.ToString("HHmmss"));
                for (int i = 1; i < timeRecord.Count; i++)
                {
                    countTime.WriteLine("RunTrimMPO_" + (timeRecord[i] - timeRecord[i - 1]) + "s");
                }
                countTime.WriteLine("Total_" + sw.Elapsed.TotalSeconds.ToString() + "s");
                countTime.Close(); countTime.Dispose();
                //印出所花費的總豪秒數
                string result1 = sw.Elapsed.TotalMinutes.ToString();
                Console.WriteLine(result1);
                               
            }
            catch (Exception e)
            {
                if (!Directory.Exists(@"D:\ERP_WMS_inventory\log"))
                {
                    Directory.CreateDirectory(@"D:\ERP_WMS_inventory\log");
                }
                File.AppendAllText(@"D:\ERP_WMS_inventory\log\" + DateTime.Now.ToString("yyyyMMdd") + ".txt", "\r\n" + company + "_" + factory + DateTime.Now.ToString("HH:mm:ss") + ":" + e + "\r\nBug is here." + BugWhere);
            }
        }

        public static List<Data.Total> RunFabricTotal(string company, string factory)
        {
            List<Data.Total> ERPdata = new List<Data.Total>();
            List<Data.Total> WMSdata = new List<Data.Total>();
            List<Data.Total> MSdata = new List<Data.Total>();
            List<Data.Total> data = new List<Data.Total>();

          

            return data;
        }

        public static List<Data.FabricGrade> RunFabricGrade(string company, string factory)
        {
            List<Data.FabricGrade> ERPdata = new List<Data.FabricGrade>();
            List<Data.FabricGrade> WMSdata = new List<Data.FabricGrade>();
            List<Data.FabricGrade> MSdata = new List<Data.FabricGrade>();
            List<Data.FabricGrade> data = new List<Data.FabricGrade>();

            

            return data;
        }

        public static List<Data.Total> RunTrimTotal(string company, string factory)
        {
            List<Data.Total> ERPdata = new List<Data.Total>();
            List<Data.Total> WMSdata = new List<Data.Total>();
            List<Data.Total> MSdata = new List<Data.Total>();
            List<Data.Total> data = new List<Data.Total>();

            return data;
        }

        public static List<Data.TrimGrade> RunTrimGrade(string company, string factory)
        {
            List<Data.TrimGrade> ERPdata = new List<Data.TrimGrade>();
            List<Data.TrimGrade> WMSdata = new List<Data.TrimGrade>();
            List<Data.TrimGrade> MSdata = new List<Data.TrimGrade>();
            List<Data.TrimGrade> data = new List<Data.TrimGrade>();

            return data;
        }

        public static List<Data.TrimMPO> RunTrimMPO(string company, string factory)
        {
            List<Data.TrimMPO> ERPdata = new List<Data.TrimMPO>();
            List<Data.TrimMPO> WMSdata = new List<Data.TrimMPO>();
            List<Data.TrimMPO> MSdata = new List<Data.TrimMPO>();
            List<Data.TrimMPO> data = new List<Data.TrimMPO>();

            return data;
        }

        public static List<Data.ERP_A_Data> RunErpADataSQL(string company, string factory)
        {
            List<Data.ERP_A_Data> data = new List<Data.ERP_A_Data>();
                       
            return data;
        }

        public static List<Data.MS_OrderData> RunMS_Order(string factory)
        {
            List<Data.MS_OrderData> data = new List<Data.MS_OrderData>();

            return data;
        }
    }

    public class Data
    {
        public class Total
        {
            public string MaterialNo { get; set; }
            public decimal ERP_Qty { get; set; }
            public decimal WMS_Qty { get; set; }
            public decimal MS_Qty { get; set; }
            public decimal Diff_Qty { get; set; }
        }
        public class FabricGrade
        {
            public string MaterialNo { get; set; }
            public string ColorID { get; set; }
            public string ERP_Color { get; set; }
            public string WMS_Color { get; set; }
            public string Grade { get; set; }
            public decimal ERP_Qty { get; set; }
            public decimal WMS_Qty { get; set; }
            public decimal MS_Qty { get; set; }
            public decimal Diff_Qty { get; set; }

        }
        public class TrimGrade
        {
            public string MaterialNo { get; set; }
            public string ColorID { get; set; }
            public string ERP_Color { get; set; }
            public string WMS_Color { get; set; }
            public string Size { get; set; }
            public string Style { get; set; }
            public string Grade { get; set; }
            public decimal ERP_Qty { get; set; }
            public decimal WMS_Qty { get; set; }
            public decimal MS_Qty { get; set; }
            public decimal Diff_Qty { get; set; }
        }
        public class TrimMPO
        {
            public string MaterialNo { get; set; }
            public string ColorID { get; set; }
            public string ERP_Color { get; set; }
            public string WMS_Color { get; set; }
            public string Size { get; set; }
            public string Style { get; set; }
            public string MPO { get; set; }
            public string Grade { get; set; }
            public decimal ERP_Qty { get; set; }
            public decimal WMS_Qty { get; set; }
            public decimal MS_Qty { get; set; }
            public decimal Diff_Qty { get; set; }
        }

        public class ERP_A_Data
        {
            public string MaterialNo { get; set; }
            public string ColorID { get; set; }
            public string ERP_Color { get; set; }
            public string Size { get; set; }
            public string Style { get; set; }
            public string MPO { get; set; }
            public string SERIAL { get; set; }
            public decimal ERP_Qty { get; set; }
        }

        public class MS_OrderData
        {
            public string MS_Order { get; set; }
            public string MaterialNo { get; set; }
            public string Type { get; set; }
            public decimal MS_Qty { get; set; }
            public string IsChangeSite { get; set; }
            public string CreateBy { get; set; }
            public string CreateTime { get; set; }
            public string FromSite { get; set; }
            public string ToSite { get; set; }
            public string FromLocation { get; set; }
            public string ToLocation { get; set; }
            public string FromMPO { get; set; }
            public string ToMPO { get; set; }
            public string FromColor { get; set; }
            public string ToColor { get; set; }
            public string FromSize { get; set; }
            public string ToSize { get; set; }
            public string FromStyle { get; set; }
            public string ToStyle { get; set; }
            public string status { get; set; }
        }
    }
}

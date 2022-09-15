using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using System.Data;
using System.Web.Configuration;
using RSI.Data;

namespace InputSomeThing
{
    class Program
    {
        static void Main(string[] args)
        {
            //startRSI();
            //parsePDF();

            DataTable dt = getSizeData("428", "1472,1473,1476,1477");
            var arr_lusthids = dt.AsEnumerable().Select(s => s.Field<long>("lusthid")).Distinct().OrderBy(o => o).ToList();

            DataTable dtCompare = getSizeData("426", "1470,1471");
            var arr_Comparelusthids = dtCompare.AsEnumerable().Select(s => s.Field<long>("lusthid")).Distinct().OrderBy(o => o).ToList();

            int PageCount = arr_lusthids.Count > arr_Comparelusthids.Count ? arr_lusthids.Count : arr_Comparelusthids.Count;

            for (int h = 0; h < PageCount; h++)
            {
                string lusthid = "";
                string lusthid_compare = "";

                try { lusthid = arr_lusthids[h].ToString(); }
                catch (Exception ex) { lusthid = "0"; }

                try { lusthid_compare = arr_Comparelusthids[h].ToString(); }
                catch (Exception ex) { lusthid_compare = "0"; }

                DataTable dtLu_SizeTable_Header = getSizeHeader(lusthid);
                DataTable dtLu_SizeTableCompare_Header = getSizeHeader(lusthid_compare);
            }

            Console.ReadLine();
        }

        static DataTable getSizeData(string _luhid, string _lusthid)
        {
            DataTable QueryResult = new DataTable();
            List<DatabaseParameter> lstParameter = new List<DatabaseParameter>();
            string sSql = @"select a.*,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15 
                from PDFTAG.dbo.Lu_SizeTable a
                join  PDFTAG.dbo.Lu_SizeTable_Header b on a.lusthid=b.lusthid
                where 1=1
                and a.luhid = "+ _luhid + @"
                and a.lusthid in (" + _lusthid + @")
                order by lusthid,rowid asc ";
            using (Database db = DatabaseFactory.GetDatabase("DB"))
            {
                QueryResult = db.ExecuteQuery(sSql, lstParameter);
            }
            return QueryResult;
        }

        static DataTable getSizeHeader(string _lusthid)
        {
            DataTable QueryResult = new DataTable();
            List<DatabaseParameter> lstParameter = new List<DatabaseParameter>();
            string sSql = @"select *
                from PDFTAG.dbo.Lu_SizeTable_Header a
                where 1=1
                and a.lusthid in (" + _lusthid + @") ";
            using (Database db = DatabaseFactory.GetDatabase("DB"))
            {
                QueryResult = db.ExecuteQuery(sSql, lstParameter);
            }
            return QueryResult;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="org"></param>
        /// <param name="newText"></param>
        /// <param name="note"></param>
        /// <param name="IsMapping"></param>
        /// <returns></returns>
        static string Compare(string org, string newText, string note, bool IsMapping = true)
        {
            if (org == newText)
                return org;
            else
            {
                if (IsMapping)
                    return "<font>原:" + org + "</font><br><font color='red'>修:" + newText + "</font><br><font color='blue'>中:" + note + "</font>";
                else
                    return "<font color='red'>修:無對應</font>";
            }
        }
        
        /// <summary>
        /// PDF轉文字
        /// </summary>
        static void parsePDF()
        {
            string sPdfPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/882390~590915~98734_Salutation_1_4_ZIP_BLACK_LYCR~000649195_20210916.pdf";
            string sTxtPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/882390~590915~98734_Salutation_1_4_ZIP_BLACK_LYCR~000649195_20210916.txt";
            //string sPdfPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/LW1DNQS BOM 1129.pdf";
            //string sTxtPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/LW1DNQS BOM 1129.txt";

            // Initialize license object
            Aspose.Pdf.License license = new Aspose.Pdf.License();
            try
            {
                // Set license
                license.SetLicense("C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/Aspose.Pdf.lic");
            }
            catch (Exception)
            {
                // something went wrong
                throw;
            }

            // Open document
            Document pdfDocument = new Document(sPdfPath);

            StringBuilder sb = new StringBuilder();
            foreach (var page in pdfDocument.Pages)
            {
                Aspose.Pdf.Text.TableAbsorber absorber = new Aspose.Pdf.Text.TableAbsorber();
                absorber.Visit(page);
                foreach (AbsorbedTable table in absorber.TableList)
                {
                    //Console.WriteLine("Table");
                    foreach (AbsorbedRow row in table.RowList)
                    {

                        sb.Append("@Row: ");

                        bool isHeader = false;
                        if (row.CellList.Any(a => a.TextFragments.Any(t => t.Text.Contains("Season:"))))
                            isHeader = true;

                        foreach (AbsorbedCell cell in row.CellList)
                        {
                            string cellText = "";
                            foreach (TextFragment fragment in cell.TextFragments)
                            {
                                if (fragment.BaselinePosition.YIndent <= 11)
                                {
                                    //Remove "Modified Date:Aug 11, 2020.."
                                    continue;
                                }

                                foreach (TextSegment seg in fragment.Segments)
                                {
                                    //if(seg.Text== "Standard")
                                    //{

                                    //}
                                    Console.Write(seg.Text + "__");
                                    cellText += (isHeader ? (seg.Text.EndsWith(":") ? "&" : "") : "") + seg.Text + " ";
                                }
                                //sb.Append(seg.Text + " | ");
                                //Console.Write($"{sb.ToString()}|");                              
                            }
                            //if (cellText.StartsWith("Modified D")) cellText = "";
                            Console.Write(" | ");
                            sb.Append(cellText + " | ");
                        }

                        Console.WriteLine();
                    }

                    sb.AppendLine("----");
                }
            }

            System.IO.File.WriteAllText(sTxtPath, sb.ToString());
        }
       
        /// <summary>
        /// 50週年程式小短片
        /// </summary>
        static void startRSI()
        {
            //string Statement0 = Console.ReadLine();
            //string Statement1 = Console.ReadLine();
            //string Statement2 = Console.ReadLine();
            //string Statement3 = Console.ReadLine();
            //string Statement4 = Console.ReadLine();
            //string Statement5 = Console.ReadLine();
            //string Statement6 = Console.ReadLine();
            string Statement7 = Console.ReadLine();
            //Console.WriteLine("某天信源的下班時間"); Thread.Sleep(2000);
            //Console.WriteLine("IT 一號:下班了下班了"); Thread.Sleep(2000);
            //Console.WriteLine("IT 二號:欸欸別忘了喔"); Thread.Sleep(2000);
            //Console.WriteLine("IT 一號表示疑惑:今天是什麼日子嗎?"); Thread.Sleep(2000);
            //Console.WriteLine("IT 二號:今天要去吃飯啊，吃免錢的你忘了喔"); Thread.Sleep(2000);
            //Console.WriteLine("IT一號:對欸50週年"); Thread.Sleep(2000);
            //Console.WriteLine("祝信源\n"); Thread.Sleep(500);

            Console.WriteLine("■■■■■■                      ■■■■■■■■                  ■■■■■■■■■■"); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■                  ■■■■■■■■■■                ■■■■■■■■■■"); Thread.Sleep(100);
            Console.WriteLine("■          ■■              ■■■             ■■                     ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■           ■■              ■■■             ■■                    ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■          ■■                 ■■■                                   ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■                      ■■■                                ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■                          ■■■                              ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■ ■■                                ■■■                             ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■  ■■                                  ■■■                          ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■   ■■                                   ■■■                        ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■     ■■                  ■■             ■■■                      ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■       ■■                 ■■             ■■■                     ■■■■     "); Thread.Sleep(100);
            Console.WriteLine("■         ■■                 ■■■■■■■■■■                ■■■■■■■■■■"); Thread.Sleep(100);
            Console.WriteLine("■           ■■                ■■■■■■■■■                 ■■■■■■■■■■"); Thread.Sleep(100);

            Console.WriteLine("");

            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■      ■    ■■■■■■          ■"); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■      ■    ■        ■        ■"); Thread.Sleep(100);
            Console.WriteLine("■■                  ■■■      ■■■            ■   ■   ■      ■■■■■■■  "); Thread.Sleep(100);
            Console.WriteLine("■■                  ■■■      ■■■    ■■■  ■ ■■■ ■     ■     ■        "); Thread.Sleep(100);
            Console.WriteLine("■■                  ■■■      ■■■        ■  ■   ■   ■    ■      ■        "); Thread.Sleep(100);
            Console.WriteLine("■■                  ■■■      ■■■       ■   ■ ■■■ ■            ■        "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■■■    ■■■      ■■■     ■     ■        ■       ■■■■■■    "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■■■    ■■■      ■■■    ■■■  ■ ■■■ ■       ■   ■        "); Thread.Sleep(100);
            Console.WriteLine("              ■■    ■■■      ■■■        ■  ■ ■  ■ ■       ■   ■        "); Thread.Sleep(100);
            Console.WriteLine("              ■■    ■■■      ■■■       ■   ■ ■  ■ ■       ■   ■        "); Thread.Sleep(100);
            Console.WriteLine("              ■■    ■■■      ■■■      ■    ■ ■■■ ■    ■■■■■■■■■"); Thread.Sleep(100);
            Console.WriteLine("              ■■    ■■■      ■■■     ■     ■        ■            ■        "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■    ■■                            ■        "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■   ■   ■■■■■■■■            ■        "); Thread.Sleep(100);

            Console.WriteLine("");

            Console.WriteLine("      ■■            ■■■■■■■■■        ■    ■                  ■        "); Thread.Sleep(100);
            Console.WriteLine("     ■ ■            ■■■■■■■■■        ■    ■           ■   ■■■  ■  "); Thread.Sleep(100);
            Console.WriteLine("    ■  ■            ■■          ■■        ■    ■          ■ ■ ■  ■ ■ ■ "); Thread.Sleep(100);
            Console.WriteLine("  ■■■■■■■      ■■          ■■    ■  ■  ■■■■      ■■  ■■■  ■■"); Thread.Sleep(100);
            Console.WriteLine(" ■     ■            ■■          ■■    ■■■■  ■  ■        ■  ■  ■   ■ "); Thread.Sleep(100);
            Console.WriteLine("■      ■            ■■          ■■    ■  ■    ■  ■       ■■ ■■■  ■■"); Thread.Sleep(100);
            Console.WriteLine("        ■            ■■■■■■■■■    ■  ■    ■  ■              ■        "); Thread.Sleep(100);
            Console.WriteLine("        ■            ■■■■■■■■■        ■    ■  ■      ■■■■■■■■■"); Thread.Sleep(100);
            Console.WriteLine("      ■■■          ■■          ■■        ■■■■■■■            ■        "); Thread.Sleep(100);
            Console.WriteLine("        ■            ■■          ■■        ■   ■                ■ ■ ■       "); Thread.Sleep(100);
            Console.WriteLine("        ■            ■■          ■■        ■  ■ ■             ■  ■  ■    "); Thread.Sleep(100);
            Console.WriteLine("        ■            ■■          ■■        ■ ■   ■           ■   ■   ■   "); Thread.Sleep(100);
            Console.WriteLine("        ■            ■■■■■■■■■        ■ ■     ■        ■    ■    ■  "); Thread.Sleep(100);
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■        ■ ■       ■     ■     ■     ■"); Thread.Sleep(500);

            showRSI();
            for (int i = 0; i < 14; i++) { Console.WriteLine(""); Thread.Sleep(100); }
            Console.WriteLine("                              科資中心祝福RSI  50歲生日快樂"); Thread.Sleep(100);
            for (int i = 0; i < 14; i++) { Console.WriteLine(""); Thread.Sleep(100); }



            Console.Read();
        }
        
        /// <summary>
        /// 可變換字體顏色及背景色
        /// </summary>
        static void showRSI()
        {
            Type type = typeof(ConsoleColor);
            //Console.ForegroundColor = ConsoleColor.White;
            //foreach (var name in Enum.GetNames(type))
            //{
            //    if ((ConsoleColor)Enum.Parse(type, name) == ConsoleColor.White) continue;
            //    Console.BackgroundColor = (ConsoleColor)Enum.Parse(type, name);
            //    strRSI();
            //}
            //Console.BackgroundColor = ConsoleColor.Black;
            foreach (var name in Enum.GetNames(type))
            {
                if ((ConsoleColor)Enum.Parse(type, name) == ConsoleColor.Black) continue;
                Console.ForegroundColor = (ConsoleColor)Enum.Parse(type, name);
                strRSI();
            }
        }

        /// <summary>
        /// 慶祝文字
        /// </summary>
        static void strRSI()
        {
            Console.WriteLine("■■■■■■                      ■■■■■■■■                  ■■■■■■■■■■");
            Console.WriteLine("■■■■■■■                  ■■■■■■■■■■                ■■■■■■■■■■");
            Console.WriteLine("■          ■■              ■■■             ■■                     ■■■■      ");
            Console.WriteLine("■           ■■              ■■■             ■■                    ■■■■      ");
            Console.WriteLine("■          ■■                 ■■■                                   ■■■■      ");
            Console.WriteLine("■■■■■■■                      ■■■                                ■■■■      ");
            Console.WriteLine("■■■■■■                          ■■■                              ■■■■      ");
            Console.WriteLine("■ ■■                                ■■■                             ■■■■      ");
            Console.WriteLine("■  ■■                                  ■■■                          ■■■■      ");
            Console.WriteLine("■   ■■                                   ■■■                        ■■■■      ");
            Console.WriteLine("■     ■■                  ■■             ■■■                      ■■■■      ");
            Console.WriteLine("■       ■■                 ■■             ■■■                     ■■■■      ");
            Console.WriteLine("■         ■■                 ■■■■■■■■■■                ■■■■■■■■■■");
            Console.WriteLine("■           ■■                ■■■■■■■■■                 ■■■■■■■■■■");

            Console.WriteLine("\n");

            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■      ■    ■■■■■■          ■          ");
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■      ■    ■        ■        ■            ");
            Console.WriteLine("■■                  ■■■      ■■■            ■   ■   ■      ■■■■■■■  ");
            Console.WriteLine("■■                  ■■■      ■■■    ■■■  ■ ■■■ ■     ■     ■        ");
            Console.WriteLine("■■                  ■■■      ■■■        ■  ■   ■   ■    ■      ■        ");
            Console.WriteLine("■■                  ■■■      ■■■       ■   ■ ■■■ ■            ■        ");
            Console.WriteLine("■■■■■■■■■    ■■■      ■■■     ■     ■        ■       ■■■■■■   ");
            Console.WriteLine("■■■■■■■■■    ■■■      ■■■    ■■■  ■ ■■■ ■       ■   ■        ");
            Console.WriteLine("              ■■    ■■■      ■■■        ■  ■ ■  ■ ■       ■   ■        ");
            Console.WriteLine("              ■■    ■■■      ■■■       ■   ■ ■  ■ ■       ■   ■        ");
            Console.WriteLine("              ■■    ■■■      ■■■      ■    ■ ■■■ ■    ■■■■■■■■■");
            Console.WriteLine("              ■■    ■■■      ■■■     ■     ■        ■            ■        ");
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■    ■■                            ■        ");
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■   ■   ■■■■■■■■            ■        "); 

            Console.WriteLine("");

            Console.WriteLine("      ■■            ■■■■■■■■■        ■    ■                  ■        ");
            Console.WriteLine("     ■ ■            ■■■■■■■■■        ■    ■           ■   ■■■  ■  ");
            Console.WriteLine("    ■  ■            ■■          ■■        ■    ■          ■ ■ ■  ■ ■ ■ ");
            Console.WriteLine("  ■■■■■■■      ■■          ■■    ■  ■  ■■■■      ■■  ■■■  ■■");
            Console.WriteLine(" ■     ■            ■■          ■■    ■■■■  ■  ■        ■  ■  ■   ■ ");
            Console.WriteLine("■      ■            ■■          ■■    ■  ■    ■  ■       ■■ ■■■  ■■");
            Console.WriteLine("        ■            ■■■■■■■■■    ■  ■    ■  ■              ■        ");
            Console.WriteLine("        ■            ■■■■■■■■■        ■    ■  ■      ■■■■■■■■■");
            Console.WriteLine("      ■■■          ■■          ■■        ■■■■■■■            ■        ");
            Console.WriteLine("        ■            ■■          ■■        ■   ■                ■ ■ ■       ");
            Console.WriteLine("        ■            ■■          ■■        ■  ■ ■             ■  ■  ■    ");
            Console.WriteLine("        ■            ■■          ■■        ■ ■   ■           ■   ■   ■   ");
            Console.WriteLine("        ■            ■■■■■■■■■        ■ ■     ■        ■    ■    ■  ");
            Console.WriteLine("■■■■■■■■■    ■■■■■■■■■        ■ ■       ■     ■     ■     ■"); Thread.Sleep(500);
        }
    }
}

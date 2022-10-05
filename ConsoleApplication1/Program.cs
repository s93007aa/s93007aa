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
using System.IO;
using RSI.Data;
using Dapper;

namespace InputSomeThing
{
    class Program
    {
        static void Main(string[] args)
        {
            //startRSI();
            //parsePDF();
            //parsePDF_UA();
            //readPDFtxt();
                        
            Console.WriteLine("DONE");
            Console.ReadLine();
        }

        /// <summary>
        /// LeetCode 1337.The K Weakest Rows in a Matrix
        /// </summary>
        /// <param name="mat"></param>
        /// <param name="k"></param>
        /// <returns></returns>
        static int[] KWeakestRows(int[][] mat, int k)
        {
            int[] outPut = new int[k];
            int[] sumArr = new int[mat.Length];
            Dictionary<int, int> dic = new Dictionary<int, int>();
            for (int i = 0; i < mat.Length; i++)
            {
                int sumInt = 0;
                foreach (var tmpInt in mat[i])
                {
                    sumInt += tmpInt;
                }
                sumArr[i] = sumInt;
                dic.Add(i, sumInt);
            }
            Array.Sort(sumArr);
            for (int i = 0; i < k; i++)
            {
                outPut[i] = dic.FirstOrDefault(x => x.Value == sumArr[i]).Key;
                dic.Remove(dic.FirstOrDefault(x => x.Value == sumArr[i]).Key);
            }
            return outPut;
        }

        /// <summary>
        /// LeetCode 412.Fizz Buzz
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        static IList<string> FizzBuzz(int n)
        {
            List<string> outPut = new List<string>();
            for (int i = 1; i <= n; i++)
            {
                string outStr = "";
                if (i % 3 == 0)
                    outStr += "Fizz";
                if (i % 5 == 0)
                    outStr += "Buzz";
                if (outStr == "")
                    outStr = i.ToString();
                outPut.Add(outStr);
            }            
            return outPut;

        }

        /// <summary>
        /// LeetCode 383.Ransom Note
        /// </summary>
        /// <param name="ransomNote"></param>
        /// <param name="magazine"></param>
        /// <returns></returns>
        static bool CanConstruct(string ransomNote, string magazine)
        {
            Dictionary<int, char> getChar = new Dictionary<int, char>();
            int i = 0;
            foreach (var strChar in magazine)
            {
                getChar.Add(i, strChar);
                i++;
            }
            foreach (var strChar in ransomNote)
            {
                if (!getChar.ContainsValue(strChar))
                    return false;
                else
                {                    
                    getChar.Remove(getChar.FirstOrDefault(x => x.Value == strChar).Key);
                }
            }
            return true;
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
                --and a.lusthid in (" + _lusthid + @")
                and trim(H1) <> 'Requested' 
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
                and a.lusthid in ('" + _lusthid + @"') ";
            using (Database db = DatabaseFactory.GetDatabase("DB"))
            {
                QueryResult = db.ExecuteQuery(sSql, lstParameter);
            }
            return QueryResult;
        }

        static DataTable getDistinctSizeHeader(string _luhid, string _lusthid)
        {
            DataTable QueryResult = new DataTable();
            List<DatabaseParameter> lstParameter = new List<DatabaseParameter>();
            string sSql = @"select distinct H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15
                from PDFTAG.dbo.Lu_SizeTable_Header a
                where 1=1
                and a.lusthid in (" + _lusthid + @") ";
            using (Database db = DatabaseFactory.GetDatabase("DB"))
            {
                QueryResult = db.ExecuteQuery(sSql, lstParameter);
            }
            return QueryResult;
        }

        static List<string> getListSizeHeader(DataTable _sizeHeader)
        {
            List<string> sizelist = new List<string>();

            foreach (DataRow drSize in _sizeHeader.Rows)
            {
                for (int i = 1; i <= 15; i++)
                {
                    string sH = drSize["H" + i].ToString();

                    if (string.IsNullOrEmpty(sH)) { break; }
                    else if (sizelist.Contains(sH)) { continue; }

                    sizelist.Add(sH);
                }
            }
            return sizelist;
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
        /// 讀PDF的TXT寫到DB裡
        /// </summary>
        static void readPDFtxt()
        {
            DataTable dt = new DataTable();
            
            StringBuilder sbLog = new StringBuilder();
            
            string sLine = "";
            string sNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            bool isMaterialColorReport = false;
            bool isLu_BOMGarmentcolor = true;
            int iRow = 1;
            int iRowid = 1;
            long luhid = 0;
            string type = "";
            string sampleStep = "";
            string sampleSize = "";
            string sSql = "";
            SQLHelper sql = new SQLHelper();

            using (System.Data.SqlClient.SqlCommand cm = new System.Data.SqlClient.SqlCommand(sSql, sql.getDbcn()))
            {
                using (StreamReader data = new StreamReader("C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/882390~590915~98734_Salutation_1_4_ZIP_BLACK_LYCR~000649195_20210916.txt"))
                {
                    #region Lu_SizeTable_1

                    sSql = @"select * from PDFTAG.dbo.Lu_SizeTable_Header where lusthid=@lusthid";

                    cm.CommandText = sSql;
                    cm.Parameters.Clear();
                    cm.Parameters.AddWithValue("@lusthid", 2677);
                    DataTable dtLu_SizeTable_Header = new DataTable();
                    using (System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(cm))
                    {
                        da.Fill(dtLu_SizeTable_Header);
                    }

                    if (dtLu_SizeTable_Header.Rows.Count > 0)
                    {
                        List<string> arrSizeHeaders = new List<string>(){
                                        "0", "2", "4", "6", "8", "10", "12", "14", "16", "18", "20", "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL"};
                        //"0", "2", "4", "6", "8", "10", "12", "14", "16", "18", "20", "XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL", "NEW" };

                        List<Lu_SizeHeaderDto> arrSizeHeaderIdxs = new List<Lu_SizeHeaderDto>();

                        for (int h = 1; h <= 15; h++)
                        {
                            if (arrSizeHeaders.Contains(dtLu_SizeTable_Header.Rows[0]["H" + h].ToString().Trim().ToUpper()))
                            {
                                arrSizeHeaderIdxs.Add(new Lu_SizeHeaderDto { Idx = h, Name = dtLu_SizeTable_Header.Rows[0]["H" + h].ToString().Trim() });
                            }
                            else if (string.IsNullOrEmpty(dtLu_SizeTable_Header.Rows[0]["SAMPLESIZE"].ToString()))
                            {
                                string[] splitSampleSize = dtLu_SizeTable_Header.Rows[0]["SAMPLESIZE"].ToString().Split(',');
                                for (int i = 0; i < splitSampleSize.Count(); i++)
                                {
                                    arrSizeHeaderIdxs.Add(new Lu_SizeHeaderDto { Idx = i + 1, Name = splitSampleSize[i].Trim().ToUpper() });
                                }
                                break;
                            }
                        }



                        sSql = @"insert into PDFTAG.dbo.Lu_SizeTable_Header_1 
(pipid,REQUESTNAME,SAMPLESIZE,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15) 
values 
(@pipid,@requestName,@sampleSize,@H1,@H2,@H3,@H4,@H5,@H6,@H7,@H8,@H9,@H10,@H11,@H12,@H13,@H14,@H15); SELECT SCOPE_IDENTITY();";

                        cm.CommandText = sSql;
                        cm.Parameters.Clear();
                        //cm.Parameters.AddWithValue("@pipid", pipid);
                        cm.Parameters.AddWithValue("@requestName", sampleStep);
                        cm.Parameters.AddWithValue("@sampleSize", sampleSize);
                        for (int h = 0; h <= 14; h++)
                        {
                            //var res = arrSizeHeaderIdxs.FirstOrDefault(x => x.Idx == h);

                            if (h < arrSizeHeaderIdxs.Count)
                            {
                                cm.Parameters.AddWithValue("@H" + (h + 1), arrSizeHeaderIdxs[h].Name);
                            }
                            else
                                cm.Parameters.AddWithValue("@H" + (h + 1), "");
                        }

                        //long lusthid_1 = Convert.ToInt64(cm.ExecuteScalar().ToString());

                        sSql = @"insert into PDFTAG.dbo.Lu_SizeTable_1(lustid_relation,luhid,rowid,codeid,Name,Criticality,TolA,TolB,HTMInstruction,lusthid,A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,isEdit) ";

                        sSql += @" select lustid as lustid_relation,luhid,rowid,codeid,Name,Criticality,TolA,TolB,HTMInstruction ";
                        //sSql += @"  ,'" + lusthid_1 + "' as lusthid ";

                        for (int h = 0; h <= 14; h++)
                        {
                            if (h < arrSizeHeaderIdxs.Count)
                                sSql += "  ,A" + arrSizeHeaderIdxs[h].Idx + " as A" + (h + 1);
                            else
                                sSql += "  ,null as A" + (h + 1);
                        }
                        sSql += @"  ,0 as isEdit ";

                        sSql += @" from PDFTAG.dbo.Lu_SizeTable ";
                        sSql += @" where luhid=@luhid and lusthid=@lusthid";


                        //sbLog.AppendLine("[Insert into PDFTAG.dbo.Lu_SizeTable_1] pipid=" + pipid + " sSql=" + sSql.Replace("@luhid", luhid.ToString()).Replace("@lusthid", lusthid.ToString()));
                        cm.CommandText = sSql;
                        cm.Parameters.Clear();
                        cm.Parameters.AddWithValue("@luhid", luhid);
                        cm.Parameters.AddWithValue("@lusthid", 2677);
                        //cm.ExecuteNonQuery();

                    }
                    #endregion
                }
            }
        }
                
        /// <summary>
        /// PDF轉文字
        /// </summary>
        static void parsePDF()
        {
            List<string> listSampleStep = new List<string>();
            List<string> listSampleSize = new List<string>();

            //string sPdfPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/882390~590915~98734_Salutation_1_4_ZIP_BLACK_LYCR~000649195_20210916.pdf";
            //string sTxtPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/882390~590915~98734_Salutation_1_4_ZIP_BLACK_LYCR~000649195_20210916.txt";
            //string sPdfPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/LW1DNQS BOM 1129.pdf";
            //string sTxtPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/LW1DNQS BOM 1129.txt";
            string sPdfPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/Semi Fitted NEW CONCEPT E...5 Semi Fitted NEW CONCEPT ESSENCE TANK CRAFT - TBD 000699495 Concept-en.pdf";
            string sTxtPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/Semi Fitted NEW CONCEPT E...5 Semi Fitted NEW CONCEPT ESSENCE TANK CRAFT - TBD 000699495 Concept-en.txt";

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
                            bool getSamplStep = false;
                            bool getSamplSize = false;
                            foreach (TextFragment fragment in cell.TextFragments)
                            {
                                //if (fragment.BaselinePosition.YIndent <= 11)
                                //{
                                //    //Remove "Modified Date:Aug 11, 2020.."
                                //    continue;
                                //}

                                
                                foreach (TextSegment seg in fragment.Segments)
                                {
                                    //if(seg.Text== "Standard")
                                    //{

                                    //}
                                    cellText += seg.Text + " ";
                                    //cellText += (isHeader ? (seg.Text.EndsWith(":") ? "&" : "") : "") + seg.Text + " ";
                                    //if (getSamplStep) { listSampleStep.Add(seg.Text); getSamplStep = false; }
                                    //else if (getSamplSize) { listSampleSize.Add(seg.Text); getSamplSize = false; }
                                    //if (seg.Text == "Sample") { getSamplStep = true; }
                                    //else if (seg.Text == "Sample Size") { getSamplSize = true; }
                                }
                                //sb.Append(seg.Text + " | ");
                                //Console.Write($"{sb.ToString()}|");                              
                            }
                            //if (cellText.StartsWith("Modified D")) cellText = "";
                            sb.Append(cellText + " | ");
                        }

                        //Console.WriteLine();
                    }

                    sb.AppendLine("----");
                }
            }

            System.IO.File.WriteAllText(sTxtPath, sb.ToString());

            //string sLine = "";
            //string sNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            //bool isMaterialColorReport = false;
            //bool isLu_BOMGarmentcolor = true;
            //string sampleStep = "";
            //string sampleSize = "";
            //int countHeader = -1;

            //using (StreamReader data = new StreamReader(sTxtPath))
            //{
            //    while (!data.EndOfStream)
            //    {
            //        sLine = data.ReadLine();
            //        if (sLine.Contains("BOM Comments"))
            //            isMaterialColorReport = true;
            //        else if (sLine.Contains("@Row: #  | Name  | Criticality"))
            //            isMaterialColorReport = false;
            //        else if (sLine.Contains("@Row: Request Name"))
            //        {
            //            countHeader++;//找到樣衣尺寸
            //        }

            //        if (isMaterialColorReport) { }
            //        else
            //        {
            //            string[] arrRow = sLine.Split(new string[] { "@Row:" }, StringSplitOptions.None);
            //            if (sLine.Contains("@Row: #  | Name"))
            //            {
            //                for (int i = 0; i < arrRow.Length; i++)
            //                {
            //                    string sRowLine = arrRow[i].Trim().TrimStart('|').TrimEnd('|');
            //                    if (i == 1)
            //                    {
            //                        string[] arrRowValue = sRowLine.Split(new string[] { "|" }, StringSplitOptions.None);
            //                        int idxOfHTMInstruction = 0;
            //                        for (int h = 0; h < 20; h++)
            //                        {
            //                            if (h < arrRowValue.Length)
            //                            {
            //                                if (arrRowValue[h].Replace(" ", "").Trim().ToLower() == "htminstruction")
            //                                {
            //                                    idxOfHTMInstruction = h;
            //                                }
            //                            }
            //                        }
            //                        bool flagSample = false;
            //                        bool flagHTM = false;
            //                        for (int h = 1; h <= 15; h++)
            //                        {
            //                            int iHIdex = (idxOfHTMInstruction == 5 ? (h + 5) : (h + 4));
            //                            string val = "";
            //                            if (iHIdex < arrRowValue.Length)
            //                            {
            //                                val = arrRowValue[iHIdex].ToString().Trim();
            //                            }
            //                            if (val == "New")
            //                            {
            //                                flagSample = true;
            //                                sampleStep = listSampleStep[countHeader];
            //                                sampleSize = listSampleSize[countHeader];
            //                                //countHeader++;
            //                            }

            //                            if (val.Replace(" ", "").Trim().ToLower() == "htminstruction")
            //                            {
            //                                idxOfHTMInstruction = iHIdex;
            //                                flagHTM = true;
            //                            }
            //                        }
            //                        //如果沒有New表示為全段尺寸表，將sampleStep、sampleSize清空
            //                        if (!flagSample) { sampleStep = ""; sampleSize = ""; }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            Console.WriteLine();
        }

        //static void parsePDF_UA(string pipid)
        static void parsePDF_UA()
        {
            SQLHelper sql = new SQLHelper();
            DataTable dt = new DataTable();
            string sSql = "";

            StringBuilder sbLog = new StringBuilder();

            List<Lu_LearnmgrItemDto> arrLu_LearnmgrItemDto = new List<Lu_LearnmgrItemDto>();


            using (System.Data.SqlClient.SqlCommand cm = new System.Data.SqlClient.SqlCommand(sSql, sql.getDbcn()))
            {

                //sSql = "select * from PDFTAG.dbo.P_inProcess where pipid='" + pipid + "' \n";

                //cm.CommandText = sSql;
                //cm.Parameters.Clear();
                //using (System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(cm))
                //{
                //    da.Fill(dt);
                //}
                

                //string titleType = dt.Rows[0]["titleType"].ToString();
                string sPDFPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/1362878 SS23 HOLD SAMPLING KIT 010722.pdf";
                
                string sSaveTxtPath = "C:/Users/jumbolin/Documents/Visual Studio 2015/Projects/Test/ConsoleApplication1/testFiles/1362878 SS23 HOLD SAMPLING KIT 010722.txt";

                //if (parse == "1" || !System.IO.File.Exists(sSaveTxtPath))
                    ConvertPDFToText3(sPDFPath, sSaveTxtPath);
                                                
                #region Parse Text for UA
                
                string sLastLine = "";
                string sLine = "";
                string sNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                bool isBom = false;
                bool isBomTypeData = false;
                bool isBomGarmentColor = false;
                string sBomType = "";
                List<int> arrBomGarmentColorStartIdx = new List<int>();
                bool isSizeTable = false;

                int iRow = 1;
                int iLineRow = 0;
                long lubcid = 0;
                string type = "";

                using (StreamReader data = new StreamReader(sSaveTxtPath))
                {
                    while (!data.EndOfStream)
                    {

                        sLine = data.ReadLine();
                        iLineRow++;

                        if (sLine.Contains("@Row:  %% (SET A:"))
                        {
                            isBom = true;
                            isSizeTable = false;
                        }

                        if (sLine.Contains("@Row: Code %%"))
                        {
                            isBom = false;
                            isSizeTable = true;
                        }

                        if (false)
                        {

                        }

                        if (isBom)
                        {
                            #region isBom

                            if (sLine.Contains("@Row:  %% (SET A:"))
                            {
                                isBomGarmentColor = true;

                                #region BomGarmentColor Header


                                sLine = data.ReadLine();
                                iLineRow++;

                                List<string> arrGarmentColor = new List<string>();
                                List<string> arrDescs = new List<string>();
                                List<string> arrCNumbers = new List<string>();

                                arrDescs = sLine.Trim().Split(new string[] { "%%" }, StringSplitOptions.None).Where(w => !string.IsNullOrWhiteSpace(w) && !w.Contains("@Row: CONTENT")).ToList();


                                sLine = data.ReadLine();
                                iLineRow++;

                                #endregion


                                if (sLine.StartsWith("@Row: Fabric")
                                    || sLine.StartsWith("@Row: Trim") || sLine.StartsWith("@Row: Thread") || sLine.StartsWith("@Row: Label") || sLine.StartsWith("@Row: Hangtag/Packaging"))
                                {

                                    isBomTypeData = true;

                                    #region Bom Type Data

                                    //Fabric      Loc       Usage          Qty

                                    List<string> arrBomHeaders = new List<string>();
                                    List<int> arrBomHeadersStartIdx = new List<int>();
                                    List<List<string>> arrLineRowDatas = new List<List<string>>();
                                    List<UA_BomDto> arrUA_BomDto = new List<UA_BomDto>();


                                    string[] arrHeaders = sLine.Trim().Split(new string[] { "%%" }, StringSplitOptions.None).Where(w => !string.IsNullOrWhiteSpace(w)).ToArray();

                                    sBomType = arrHeaders[0].Replace("@Row:", "").Trim();

                                    string sFirstCN = arrHeaders[4].Trim();

                                    bool isStartColor = false;
                                    foreach (var header in arrHeaders)
                                    {

                                        if (header.Contains(sFirstCN))
                                        {
                                            isStartColor = true;
                                        }
                                        if (isStartColor)
                                            arrCNumbers.Add(header);
                                    }

                                    int iBomHeadersCount = arrHeaders.Length;


                                    iRow = 1;
                                    while (iRow < 100)
                                    {
                                        sLastLine = sLine;
                                        sLine = data.ReadLine();
                                        iLineRow++;


                                        if (sLine.StartsWith("@Row: " + sFirstCN))
                                        {
                                            arrGarmentColor.AddRange(sLine.Replace("@Row:", "").Trim().Split(new string[] { "%%" }, StringSplitOptions.None).Where(w => !string.IsNullOrWhiteSpace(w)).ToArray());
                                            isBomTypeData = false;
                                            break;
                                        }


                                        if (sLine.StartsWith("@Row: Fabric") || sLine.StartsWith("@Row: Trim") || sLine.StartsWith("@Row: Thread") || sLine.StartsWith("@Row: Label") || sLine.StartsWith("@Row: Hangtag/Packaging"))
                                        {
                                            //新Type, 相同欄位
                                            sBomType = sLine.Trim().Split(new string[] { "%%" }, StringSplitOptions.None).Where(w => !string.IsNullOrWhiteSpace(w)).ToArray()[0].Replace("@Row:", "").Trim();
                                            continue;
                                        }

                                        string[] arrParts = sLine.Trim().Split(new string[] { "%%" }, StringSplitOptions.None);
                                        int iLength = arrParts.Length;

                                        arrUA_BomDto.Add(new UA_BomDto()
                                        {
                                            Type = sBomType,
                                            rowid = iRow,
                                            SupplierArticle = arrParts[0].Replace("@Row:", "").Trim(),
                                            Usage = arrParts[2].Replace("@Row:", "").Trim(),

                                            B1 = iLength >= 5 ? arrParts[4].Trim() : "",
                                            B2 = iLength >= 6 ? arrParts[5].Trim() : "",
                                            B3 = iLength >= 7 ? arrParts[6].Trim() : "",
                                            B4 = iLength >= 8 ? arrParts[7].Trim() : "",
                                            B5 = iLength >= 8 ? arrParts[8].Trim() : "",
                                            B6 = iLength >= 10 ? arrParts[9].Trim() : "",
                                            B7 = iLength >= 11 ? arrParts[10].Trim() : "",
                                            B8 = iLength >= 12 ? arrParts[11].Trim() : "",
                                            B9 = iLength >= 13 ? arrParts[12].Trim() : "",
                                            B10 = iLength >= 14 ? arrParts[13].Trim() : "",

                                        });

                                        iRow++;
                                    }

                                    #region Insert PDFTAG.dbo.UA_BOM

                                //    sSql = @"insert into PDFTAG.dbo.UA_BOMGarmentcolor  
                                //(pipid,luhid,HeaderCount,A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A1Desc,A2Desc,A3Desc,A4Desc,A5Desc,A6Desc,A7Desc,A8Desc,A9Desc,A10Desc,A1Number,A2Number,A3Number,A4Number,A5Number,A6Number,A7Number,A8Number,A9Number,A10Number) 
                                //values 
                                //(@pipid,@luhid,@HeaderCount,@A1,@A2,@A3,@A4,@A5,@A6,@A7,@A8,@A9,@A10,@A1Desc,@A2Desc,@A3Desc,@A4Desc,@A5Desc,@A6Desc,@A7Desc,@A8Desc,@A9Desc,@A10Desc,@A1Number,@A2Number,@A3Number,@A4Number,@A5Number,@A6Number,@A7Number,@A8Number,@A9Number,@A10Number); SELECT SCOPE_IDENTITY();";

                                //    int iGarmentColor = arrGarmentColor.Count;
                                //    int iGarmentColorDescs = arrDescs.Count;
                                //    int iGarmentColorNumbers = arrCNumbers.Count;

                                //    cm.CommandText = sSql;
                                //    cm.Parameters.Clear();
                                //    cm.Parameters.AddWithValue("@pipid", pipid);
                                //    cm.Parameters.AddWithValue("@luhid", 0);
                                //    cm.Parameters.AddWithValue("@HeaderCount", iGarmentColor);
                                //    cm.Parameters.AddWithValue("@A1", iGarmentColor >= 1 ? arrGarmentColor[0].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A2", iGarmentColor >= 2 ? arrGarmentColor[1].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A3", iGarmentColor >= 3 ? arrGarmentColor[2].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A4", iGarmentColor >= 4 ? arrGarmentColor[3].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A5", iGarmentColor >= 5 ? arrGarmentColor[4].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A6", iGarmentColor >= 6 ? arrGarmentColor[5].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A7", iGarmentColor >= 7 ? arrGarmentColor[6].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A8", iGarmentColor >= 8 ? arrGarmentColor[7].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A9", iGarmentColor >= 9 ? arrGarmentColor[8].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A10", iGarmentColor >= 10 ? arrGarmentColor[9].Trim() : "");

                                //    cm.Parameters.AddWithValue("@A1Desc", iGarmentColorDescs >= 1 ? arrDescs[0].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A2Desc", iGarmentColorDescs >= 2 ? arrDescs[1].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A3Desc", iGarmentColorDescs >= 3 ? arrDescs[2].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A4Desc", iGarmentColorDescs >= 4 ? arrDescs[3].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A5Desc", iGarmentColorDescs >= 5 ? arrDescs[4].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A6Desc", iGarmentColorDescs >= 6 ? arrDescs[5].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A7Desc", iGarmentColorDescs >= 7 ? arrDescs[6].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A8Desc", iGarmentColorDescs >= 8 ? arrDescs[7].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A9Desc", iGarmentColorDescs >= 9 ? arrDescs[8].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A10Desc", iGarmentColorDescs >= 10 ? arrDescs[9].Trim() : "");

                                //    cm.Parameters.AddWithValue("@A1Number", iGarmentColorNumbers >= 1 ? arrCNumbers[0].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A2Number", iGarmentColorNumbers >= 2 ? arrCNumbers[1].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A3Number", iGarmentColorNumbers >= 3 ? arrCNumbers[2].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A4Number", iGarmentColorNumbers >= 4 ? arrCNumbers[3].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A5Number", iGarmentColorNumbers >= 5 ? arrCNumbers[4].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A6Number", iGarmentColorNumbers >= 6 ? arrCNumbers[5].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A7Number", iGarmentColorNumbers >= 7 ? arrCNumbers[6].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A8Number", iGarmentColorNumbers >= 8 ? arrCNumbers[7].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A9Number", iGarmentColorNumbers >= 9 ? arrCNumbers[8].Trim() : "");
                                //    cm.Parameters.AddWithValue("@A10Number", iGarmentColorNumbers >= 10 ? arrCNumbers[9].Trim() : "");


                                //    lubcid = Convert.ToInt64(cm.ExecuteScalar().ToString());


                                    foreach (var item in arrUA_BomDto)
                                    {

                                        //sSql = @"insert into  PDFTAG.dbo.UA_BOM
                                        //                        (pipid,luhid,lubcid,type,rowid,Usage,SupplierArticle,Supplier,B1,B2,B3,B4,B5,B6,B7,B8,B9,B10,isEdit) 
                                        //                        values 
                                        //                        (@pipid,@luhid,@lubcid,@type,@rowid,@Usage,@SupplierArticle,@Supplier,@B1,@B2,@B3,@B4,@B5,@B6,@B7,@B8,@B9,@B10,@isEdit);";

                                        //cm.CommandText = sSql;
                                        //cm.Parameters.Clear();
                                        //cm.Parameters.AddWithValue("@pipid", pipid);
                                        //cm.Parameters.AddWithValue("@luhid", pipid);
                                        //cm.Parameters.AddWithValue("@lubcid", lubcid);
                                        //cm.Parameters.AddWithValue("@type", item.Type);
                                        //cm.Parameters.AddWithValue("@rowid", item.rowid);
                                        //cm.Parameters.AddWithValue("@SupplierArticle", item.SupplierArticle);
                                        //cm.Parameters.AddWithValue("@Usage", item.Usage);
                                        //cm.Parameters.AddWithValue("@Supplier", "");


                                        //cm.Parameters.AddWithValue("@B1", item.B1);
                                        //cm.Parameters.AddWithValue("@B2", item.B2);
                                        //cm.Parameters.AddWithValue("@B3", item.B3);
                                        //cm.Parameters.AddWithValue("@B4", item.B4);
                                        //cm.Parameters.AddWithValue("@B5", item.B5);
                                        //cm.Parameters.AddWithValue("@B6", item.B6);
                                        //cm.Parameters.AddWithValue("@B7", item.B7);
                                        //cm.Parameters.AddWithValue("@B8", item.B8);
                                        //cm.Parameters.AddWithValue("@B9", item.B9);
                                        //cm.Parameters.AddWithValue("@B10", item.B10);

                                        //cm.Parameters.AddWithValue("@isEdit", 0);
                                        //cm.ExecuteNonQuery();
                                    }
                                    #endregion

                                    #endregion

                                }

                            }

                            #endregion
                        }
                        else if (isSizeTable)
                        {
                            #region SizeTable

                            if (sLine.StartsWith("@Row: Code %%"))
                            {
                                #region SizeTable Data

                                // POM        Description              Add'l     Variation    QC       Tol(-)    Tol(+)

                                UA_SizeTableHeaderDto resUA_SizeTableHeaderDto = new UA_SizeTableHeaderDto();
                                List<string> arrSizeTableHeaders = new List<string>();
                                List<string> arrSizeHeaders = new List<string>();
                                List<int> arrSizeTableHeadersStartIdx = new List<int>();
                                List<List<string>> arrLineRowDatas = new List<List<string>>();
                                List<string> arrRowDatas = new List<string>();
                                int iTolAddIdx = 0;
                                long lusthid = 0;
                                bool isStartSizeColumn = false;

                                string[] arrHeaders = sLine.Trim().Split(new string[] { "%%" }, StringSplitOptions.None).Where(w => !string.IsNullOrWhiteSpace(w)).ToArray();


                                foreach (var header in arrHeaders)
                                {
                                    if (isStartSizeColumn)
                                    {
                                        arrSizeHeaders.Add(header);
                                    }


                                    if (header.Contains("Tol(+)"))
                                    {
                                        iTolAddIdx = sLine.IndexOf(header);
                                        isStartSizeColumn = true;
                                    }
                                }

                                int iSizeHeaderLength = arrSizeHeaders.Count;
                                resUA_SizeTableHeaderDto = new UA_SizeTableHeaderDto()
                                {
                                    HeaderCount = iSizeHeaderLength,
                                    H1 = iSizeHeaderLength >= 1 ? arrSizeHeaders[0].Trim() : "",
                                    H2 = iSizeHeaderLength >= 2 ? arrSizeHeaders[1].Trim() : "",
                                    H3 = iSizeHeaderLength >= 3 ? arrSizeHeaders[2].Trim() : "",
                                    H4 = iSizeHeaderLength >= 4 ? arrSizeHeaders[3].Trim() : "",
                                    H5 = iSizeHeaderLength >= 5 ? arrSizeHeaders[4].Trim() : "",
                                    H6 = iSizeHeaderLength >= 6 ? arrSizeHeaders[5].Trim() : "",
                                    H7 = iSizeHeaderLength >= 7 ? arrSizeHeaders[6].Trim() : "",
                                    H8 = iSizeHeaderLength >= 8 ? arrSizeHeaders[7].Trim() : "",
                                    H9 = iSizeHeaderLength >= 9 ? arrSizeHeaders[8].Trim() : "",
                                    H10 = iSizeHeaderLength >= 10 ? arrSizeHeaders[9].Trim() : "",
                                    H11 = iSizeHeaderLength >= 11 ? arrSizeHeaders[10].Trim() : "",
                                    H12 = iSizeHeaderLength >= 12 ? arrSizeHeaders[11].Trim() : "",
                                    H13 = iSizeHeaderLength >= 13 ? arrSizeHeaders[12].Trim() : "",
                                    H14 = iSizeHeaderLength >= 14 ? arrSizeHeaders[13].Trim() : "",
                                    H15 = iSizeHeaderLength >= 15 ? arrSizeHeaders[14].Trim() : "",

                                };


                                sLine = data.ReadLine();//Header 有2行
                                iLineRow++;

                                iRow = 1;


                                List<UA_SizeTableDto> arrUA_SizeTableDtos = new List<UA_SizeTableDto>();
                                while (iRow < 100)
                                {
                                    sLine = data.ReadLine();
                                    iLineRow++;

                                    sLastLine = sLine;

                                    if (string.IsNullOrEmpty(sLine) || sLine.Contains("@Row: Status:"))
                                    {
                                        break;
                                    }

                                    string[] arrParts = sLine.Trim().Split(new string[] { "%%" }, StringSplitOptions.None);
                                    int iLength = arrParts.Length;

                                    arrUA_SizeTableDtos.Add(new UA_SizeTableDto()
                                    {
                                        rowid = iRow,
                                        Code = arrParts[0].Trim().Replace("@Row:", ""),
                                        Description = arrParts[1].Trim(),
                                        TolA = arrParts[2].Trim(),
                                        TolB = arrParts[3].Trim(),

                                        A1 = iLength >= 5 ? arrParts[4].Trim() : "",
                                        A2 = iLength >= 6 ? arrParts[5].Trim() : "",
                                        A3 = iLength >= 7 ? arrParts[6].Trim() : "",
                                        A4 = iLength >= 8 ? arrParts[7].Trim() : "",
                                        A5 = iLength >= 9 ? arrParts[8].Trim() : "",
                                        A6 = iLength >= 10 ? arrParts[9].Trim() : "",
                                        A7 = iLength >= 11 ? arrParts[10].Trim() : "",
                                        A8 = iLength >= 12 ? arrParts[11].Trim() : "",
                                        A9 = iLength >= 13 ? arrParts[12].Trim() : "",
                                        A10 = iLength >= 14 ? arrParts[13].Trim() : "",
                                        A11 = iLength >= 15 ? arrParts[14].Trim() : "",
                                        A12 = iLength >= 16 ? arrParts[15].Trim() : "",
                                        A13 = iLength >= 17 ? arrParts[16].Trim() : "",
                                        A14 = iLength >= 18 ? arrParts[17].Trim() : "",
                                        A15 = iLength >= 19 ? arrParts[18].Trim() : "",

                                    });

                                    iRow++;
                                }

                                #region Insert PDFTAG.dbo.UA_SizeTable


//                                sSql = @"insert into PDFTAG.dbo.UA_SizeTable_Header 
//(pipid,HeaderCount,H1,H2,H3,H4,H5,H6,H7,H8,H9,H10,H11,H12,H13,H14,H15) 
//values 
//(@pipid,@HeaderCount,@H1,@H2,@H3,@H4,@H5,@H6,@H7,@H8,@H9,@H10,@H11,@H12,@H13,@H14,@H15); SELECT SCOPE_IDENTITY();";

//                                cm.CommandText = sSql;
//                                cm.Parameters.Clear();
//                                cm.Parameters.AddWithValue("@pipid", pipid);
//                                cm.Parameters.AddWithValue("@HeaderCount", resUA_SizeTableHeaderDto.HeaderCount);
//                                cm.Parameters.AddWithValue("@H1", resUA_SizeTableHeaderDto.H1);
//                                cm.Parameters.AddWithValue("@H2", resUA_SizeTableHeaderDto.H2);
//                                cm.Parameters.AddWithValue("@H3", resUA_SizeTableHeaderDto.H3);
//                                cm.Parameters.AddWithValue("@H4", resUA_SizeTableHeaderDto.H4);
//                                cm.Parameters.AddWithValue("@H5", resUA_SizeTableHeaderDto.H5);
//                                cm.Parameters.AddWithValue("@H6", resUA_SizeTableHeaderDto.H6);
//                                cm.Parameters.AddWithValue("@H7", resUA_SizeTableHeaderDto.H7);
//                                cm.Parameters.AddWithValue("@H8", resUA_SizeTableHeaderDto.H8);
//                                cm.Parameters.AddWithValue("@H9", resUA_SizeTableHeaderDto.H9);
//                                cm.Parameters.AddWithValue("@H10", resUA_SizeTableHeaderDto.H10);
//                                cm.Parameters.AddWithValue("@H11", resUA_SizeTableHeaderDto.H11);
//                                cm.Parameters.AddWithValue("@H12", resUA_SizeTableHeaderDto.H12);
//                                cm.Parameters.AddWithValue("@H13", resUA_SizeTableHeaderDto.H13);
//                                cm.Parameters.AddWithValue("@H14", resUA_SizeTableHeaderDto.H14);
//                                cm.Parameters.AddWithValue("@H15", resUA_SizeTableHeaderDto.H15);


//                                lusthid = Convert.ToInt64(cm.ExecuteScalar().ToString());

                                foreach (var item in arrUA_SizeTableDtos)
                                {
//                                    sSql = @"insert into PDFTAG.dbo.UA_SizeTable 
//(pipid,luhid,rowid,Code,Description,TolA,TolB,lusthid,A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,isEdit) 
//values 
//(@pipid,@luhid,@rowid,@Code,@Description,@TolA,@TolB,@lusthid,@A1,@A2,@A3,@A4,@A5,@A6,@A7,@A8,@A9,@A10,@A11,@A12,@A13,@A14,@A15,@isEdit)  SELECT SCOPE_IDENTITY();";

//                                    cm.CommandText = sSql;
//                                    cm.Parameters.Clear();
//                                    cm.Parameters.AddWithValue("@pipid", pipid);
//                                    cm.Parameters.AddWithValue("@luhid", pipid);
//                                    cm.Parameters.AddWithValue("@rowid", item.rowid);
//                                    cm.Parameters.AddWithValue("@Code", item.Code);
//                                    cm.Parameters.AddWithValue("@Description", item.Description);
//                                    cm.Parameters.AddWithValue("@TolA", item.TolA);
//                                    cm.Parameters.AddWithValue("@TolB", item.TolB);

//                                    cm.Parameters.AddWithValue("@lusthid", lusthid);
//                                    cm.Parameters.AddWithValue("@A1", item.A1);
//                                    cm.Parameters.AddWithValue("@A2", item.A2);
//                                    cm.Parameters.AddWithValue("@A3", item.A3);
//                                    cm.Parameters.AddWithValue("@A4", item.A4);
//                                    cm.Parameters.AddWithValue("@A5", item.A5);
//                                    cm.Parameters.AddWithValue("@A6", item.A6);
//                                    cm.Parameters.AddWithValue("@A7", item.A7);
//                                    cm.Parameters.AddWithValue("@A8", item.A8);
//                                    cm.Parameters.AddWithValue("@A9", item.A9);
//                                    cm.Parameters.AddWithValue("@A10", item.A10);
//                                    cm.Parameters.AddWithValue("@A11", item.A11);
//                                    cm.Parameters.AddWithValue("@A12", item.A12);
//                                    cm.Parameters.AddWithValue("@A13", item.A13);
//                                    cm.Parameters.AddWithValue("@A14", item.A14);
//                                    cm.Parameters.AddWithValue("@A15", item.A15);

//                                    cm.Parameters.AddWithValue("@isEdit", 0);
//                                    cm.ExecuteNonQuery();

                                }
                                #endregion

                                #endregion
                            }
                            #endregion
                        }
                        iRow++;
                        sLastLine = sLine;
                    }
                    //sSql = "select a.*,b.season,b.style,b.generateddate,c.gmname \n";
                    //sSql += "from PDFTAG.dbo.P_inProcess a              \n";
                    //sSql += " left join PDFTAG.dbo.UA_Header b on a.pipid=b.pipid and b.isshow=0             \n";
                    //sSql += " left join PDFTAG.dbo.GroupManage c on a.gmid=c.gmid and c.isshow=0             \n";
                    //sSql += " where 1=1 and a.isshow=0 and a.pipid='" + pipid + "' \n";

                    //cm.CommandText = sSql;
                    //cm.Parameters.Clear();
                    //dt = new DataTable();
                    //using (System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(cm))
                    //{
                    //    da.Fill(dt);
                    //}

                    //string ptitle = dt.Rows[0]["style"].ToString();
                    //string season2 = dt.Rows[0]["season"].ToString();
                    //string new_season = "";
                    //if (titleType == "1")
                    //{
                    //    ptitle += "-" + season2;

                    //}

                    //sSql = "update PDFTAG.dbo.P_inProcess    \n";
                    //sSql += " set mdate=@mdate \n";
                    //if (titleType == "1")
                    //{
                    //    sSql += " ,ptitle=@ptitle \n";
                    //}
                    //sSql += "where pipid=@pipid \n";
                    //cm.CommandText = sSql;
                    //cm.Parameters.Clear();
                    //cm.Parameters.AddWithValue("@pipid", pipid);
                    //cm.Parameters.AddWithValue("@mdate", sNow);
                    //cm.Parameters.AddWithValue("@ptitle", ptitle);
                    //cm.ExecuteNonQuery();
                }

                #endregion
                Console.WriteLine("執行完成!");
            }
        }

        static void ConvertPDFToText3(string sPdfPath, string sTxtPath)
        {
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

            Document pdfDocument = new Document(sPdfPath);

            StringBuilder sbFinal = new StringBuilder();
            StringBuilder sb = new StringBuilder();
            foreach (var page in pdfDocument.Pages)
            {
                Aspose.Pdf.Text.TableAbsorber absorber = new Aspose.Pdf.Text.TableAbsorber();
                absorber.Visit(page);
                foreach (AbsorbedTable table in absorber.TableList)
                {
                    Console.WriteLine("Table");
                    foreach (AbsorbedRow row in table.RowList)
                    {
                        sb = new StringBuilder();
                        sb.Append("@Row: ");

                        bool isHeader = false;

                        foreach (AbsorbedCell cell in row.CellList)
                        {
                            string cellText = "";
                            foreach (TextFragment fragment in cell.TextFragments)
                            {
                                //if (fragment.BaselinePosition.YIndent <= 11)
                                //{
                                //    //Remove "Modified Date:Aug 11, 2020.."
                                //    continue;
                                //}

                                foreach (TextSegment seg in fragment.Segments)
                                {
                                    //if(seg.Text== "Standard")
                                    //{

                                    //}

                                    cellText += seg.Text;
                                }
                                //sb.Append(seg.Text + " | ");
                                //Console.Write($"{sb.ToString()}|");                              
                            }

                            sb.Append(cellText + " %% ");
                        }

                        sbFinal.AppendLine(sb.ToString());
                    }

                    sb.AppendLine("----");
                }
            }

            System.IO.File.WriteAllText(sTxtPath, sbFinal.ToString());
        }

        public class Lu_SizeHeaderDto
        {
            public int Idx { get; set; }
            public string Name { get; set; }
        }

        static string ConvertToStrDouble(string sVal)
        {
            return sVal;
            if (string.IsNullOrWhiteSpace(sVal) || sVal.Contains("--"))
                return "";

            string res = "";
            double dVal = 0;
            string[] arrValue = sVal.Split(' ');

            if (arrValue.Length == 1)
            {
                if (arrValue[0].Contains("/"))
                {
                    string[] arr = arrValue[0].Split('/');
                    double dVal1 = 0;
                    double dVal2 = 0;
                    double dVal3 = 0;
                    double.TryParse(arr[0], out dVal1);
                    double.TryParse(arr[1], out dVal2);
                    if (dVal2 > 0)
                        dVal3 = dVal1 / dVal2;

                    double.TryParse(arrValue[0], out dVal);
                    dVal = dVal + dVal3;
                }
                else
                    double.TryParse(arrValue[0], out dVal);
            }
            else if (arrValue.Length == 2)
            {
                string s = arrValue[0];
                string[] arr = arrValue[1].Split('/');
                double dVal1 = 0;
                double dVal2 = 0;
                double dVal3 = 0;
                double.TryParse(arr[0], out dVal1);
                double.TryParse(arr[1], out dVal2);
                if (dVal2 > 0)
                    dVal3 = dVal1 / dVal2;

                double.TryParse(arrValue[0], out dVal);

                dVal = dVal + dVal3;
            }
            return dVal.ToString();
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

        public class Lu_LearnmgrItemDto
        {
            public long liid { get; set; }
            public string ColSource { get; set; }
            public string ColName { get; set; }
            public string FirstCharTermname_org { get; set; }
            public string Termname_org { get; set; }
            public string Termname { get; set; }

        }
        public class UA_BomDto
        {
            public long lubid { get; set; }
            public long luhid { get; set; }
            public string Type { get; set; }
            public int rowid { get; set; }
            public string SupplierArticle { get; set; }
            public string Usage { get; set; }
            public string Supplier { get; set; }
            public string lusthid { get; set; }
            public string B1 { get; set; }
            public string B2 { get; set; }
            public string B3 { get; set; }
            public string B4 { get; set; }
            public string B5 { get; set; }
            public string B6 { get; set; }
            public string B7 { get; set; }
            public string B8 { get; set; }
            public string B9 { get; set; }
            public string B10 { get; set; }
            public string B11 { get; set; }
            public string B12 { get; set; }
            public string B13 { get; set; }
            public string B14 { get; set; }
            public string B15 { get; set; }
            public long org_lubid { get; set; }
            public long lubcid { get; set; }
            public bool isEdit { get; set; }
        }
        public class UA_SizeTableHeaderDto
        {
            public long lusthid { get; set; }
            public long pipid { get; set; }
            public int HeaderCount { get; set; }
            public string H1 { get; set; }
            public string H2 { get; set; }
            public string H3 { get; set; }
            public string H4 { get; set; }
            public string H5 { get; set; }
            public string H6 { get; set; }
            public string H7 { get; set; }
            public string H8 { get; set; }
            public string H9 { get; set; }
            public string H10 { get; set; }
            public string H11 { get; set; }
            public string H12 { get; set; }
            public string H13 { get; set; }
            public string H14 { get; set; }
            public string H15 { get; set; }
        }
        public class UA_SizeTableDto
        {
            public long lustid { get; set; }
            public long luhid { get; set; }
            public int rowid { get; set; }
            public string Code { get; set; }
            public string Description { get; set; }
            public string TolA { get; set; }
            public string TolB { get; set; }
            public string lusthid { get; set; }
            public string A1 { get; set; }
            public string A2 { get; set; }
            public string A3 { get; set; }
            public string A4 { get; set; }
            public string A5 { get; set; }
            public string A6 { get; set; }
            public string A7 { get; set; }
            public string A8 { get; set; }
            public string A9 { get; set; }
            public string A10 { get; set; }
            public string A11 { get; set; }
            public string A12 { get; set; }
            public string A13 { get; set; }
            public string A14 { get; set; }
            public string A15 { get; set; }
            public long org_lustid { get; set; }
            public bool isEdit { get; set; }
        }
    }
}

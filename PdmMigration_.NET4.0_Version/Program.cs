using System;
using System.Configuration;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PdmMigration
{
    class Program
    {
        public static string catalogFile = @"";
        public static string inputFile = @"";
        public static string batchFile = @"";
        public static string serverName = "";
        public static string outputFile = @"";
        public static string misfitToys = @"";
        public static string jobTicketLocation = @"";
        public static string uncRawPrefix = @"";
        public static string uncPdfPrefix = @"";
        public static string adlibDTD = @"";
        public static DateTime recentDateTime = DateTime.MinValue;
        public static bool isWindows = false;
        public static bool isLuDateTime = false;
        public static bool isIeDateTime = false;
        public static bool isSLC = false;

        public static void LoadConfig()
        {
            catalogFile = ConfigurationManager.AppSettings["catalogFile"];
            inputFile = ConfigurationManager.AppSettings["inputFile"];
            batchFile = ConfigurationManager.AppSettings["batchFile"];
            serverName = ConfigurationManager.AppSettings["serverName"];
            outputFile = ConfigurationManager.AppSettings["outputFile"];
            misfitToys = ConfigurationManager.AppSettings["misfitToys"];
            jobTicketLocation = ConfigurationManager.AppSettings["jobTicketLocation"];
            uncRawPrefix = ConfigurationManager.AppSettings["uncRawPrefix"];
            uncPdfPrefix = ConfigurationManager.AppSettings["uncPdfPrefix"];
            adlibDTD = ConfigurationManager.AppSettings["adlibDTD"];
            recentDateTime = DateTime.Parse(ConfigurationManager.AppSettings["recentDateTime"]);
            isWindows = Convert.ToBoolean(ConfigurationManager.AppSettings["isWindows"]);
            isLuDateTime = Convert.ToBoolean(ConfigurationManager.AppSettings["isLuDateTime"]);
            isIeDateTime = Convert.ToBoolean(ConfigurationManager.AppSettings["isIeDateTime"]);
            isSLC = Convert.ToBoolean(ConfigurationManager.AppSettings["isSLC"]);

            Console.WriteLine(catalogFile);
            Console.WriteLine(inputFile);
            Console.WriteLine(batchFile);
            Console.WriteLine(serverName);
            Console.WriteLine(outputFile);
            Console.WriteLine(misfitToys);
            Console.WriteLine(jobTicketLocation);
            Console.WriteLine(uncRawPrefix);
            Console.WriteLine(uncPdfPrefix);
            Console.WriteLine(recentDateTime.ToString());
            Console.WriteLine(isWindows);
            Console.WriteLine(isLuDateTime);
            Console.WriteLine(isIeDateTime);
            Console.WriteLine(isSLC);
        }

        public static bool IsExt(string token)
        {
            switch (token.ToLower())
            {
                case "7z":
                case "cad":
                case "cg4":
                case "csh":
                case "csv":
                case "db":
                case "dis":
                case "dll_crea":
                case "dos":
                case "dot":
                case "dwg":
                case "dxf":
                case "edt":
                case "gif":
                case "gp4":
                case "gwk":
                case "hp":
                case "hpdf":
                case "hpg":
                case "hpp":
                case "htm":
                case "html":
                case "ini":
                case "jpg":
                case "js":
                case "mdb":
                case "mht":
                case "mil":
                case "msg":
                case "obd":
                case "oft":
                case "pcx":
                case "pdf":
                case "plt":
                case "png":
                case "ppt":
                case "pptx":
                case "pra":
                case "prt":
                case "reg":
                case "rss":
                case "rst":
                case "rtf":
                case "smf":
                case "ss":
                case "ss_old":
                case "tif":
                case "txt":
                case "url":
                case "vsd":
                case "wdf":
                case "wrl":
                case "xls":
                case "xlsx":
                case "xlt":
                case "xps":
                case "xs":
                case "xst":
                case "z":
                case "z_old":
                case "zip":
                case "flv":
                case "mpg":
                case "doc":
                case "docx":
                    return true;
                default:
                    return false;
            }
        }

        public static bool IsPdfAble(string fileName)
        {
            if (fileName.EndsWith(".Z"))
            {
                return false;
            }

            if (fileName.EndsWith("._"))
            {
                return false;
            }

            if (fileName.ToLower().EndsWith(".zip"))
            {
                return false;
            }

            if (fileName.ToLower().EndsWith(".doc"))
            {
                return false;
            }

            if (fileName.ToLower().EndsWith(".docx"))
            {
                return false;
            }

            if (fileName.ToLower().EndsWith(".mpg"))
            {
                return false;
            }

            if (fileName.ToLower().EndsWith(".flv"))
            {
                return false;
            }

            return true;
        }

        public static void JobTicketGenerator(Dictionary<string, List<PdmItem>> dictionary, List<string> batchLines)
        {
            int counter = 0;

            foreach (KeyValuePair<string, List<PdmItem>> kvp in dictionary)
            {
                counter++;
                Console.WriteLine("Key: " + kvp.Key);
                Console.WriteLine("Value: " + kvp.Value.Count);
                //if (kvp.Key != "-98081.AF")
                //{
                //    Console.WriteLine("     SKIPPING: " + kvp.Key);
                //    continue;
                //}

                //Console.WriteLine("NOT SKIPPING: " + kvp.Key);

                //if there is only one kvp, then we already have a pdf somewhere in theory
                if (kvp.Value.Count < 2)
                {
                    StringBuilder sourcePdfBuilder = new StringBuilder();

                    //find pdf and copy to correct folder; build batch file
                    if (kvp.Value[0].UncRaw.EndsWith(".pdf"))
                    {
                        sourcePdfBuilder.Append(kvp.Value[0].UncRaw.Replace("web\\", "web\\pdf\\"));
                    }
                    else
                    {
                        sourcePdfBuilder.Append(kvp.Value[0].UncRaw.Replace("web\\", "web\\pdf\\") + ".pdf");
                    }

                    if (File.Exists(sourcePdfBuilder.ToString()))
                    {
                        if(sourcePdfBuilder.ToString().Contains(" "))
                        {
                            batchLines.Add("Copy \"" + sourcePdfBuilder.ToString() + "\" " + uncPdfPrefix + "\\" + kvp.Key + ".pdf");
                        }
                        else
                        {
                            batchLines.Add("Copy " + sourcePdfBuilder.ToString() + " " + uncPdfPrefix + "\\" + kvp.Key + ".pdf");
                        }
                        continue;
                    }
                    else
                    {
                        //do nothing and build the job ticket
                        batchLines.Add("REM FILE DOES NOT EXIST: " + sourcePdfBuilder.ToString());
                    }
                }

                StringBuilder jobTicket = new StringBuilder();

                jobTicket.AppendLine("<?xml version=\"1.0\" encoding=\"ISO-8859-1\" ?>");
                jobTicket.AppendLine("<?AdlibExpress applanguage = \"USA\" appversion = \"4.11.0\" dtdversion = \"2.6\" ?>");
                jobTicket.AppendLine("<!DOCTYPE JOBS SYSTEM \"" + adlibDTD + "\">");
                jobTicket.AppendLine("<JOBS xmlns:JOBS=\"http://www.adlibsoftware.com\" xmlns:JOB=\"http://www.adlibsoftware.com\">");
                jobTicket.AppendLine("<JOB>");
                jobTicket.AppendLine("<JOB:DOCINPUTS>");

                DateTime mostRecentDate = DateTime.MinValue;
                foreach (var i in kvp.Value)
                {
                    //Find most recent date in list
                    if (i.FileDateTime.Date > mostRecentDate)
                    {
                        mostRecentDate = i.FileDateTime.Date;
                    }
                }

                var orderedItemShtNums = kvp.Value.OrderBy(x => x.ItemShtNum);

                foreach (var i in orderedItemShtNums)
                {
                    string filename = i.FileName;

                    if (filename.EndsWith(".Z") || filename.EndsWith("._"))
                    {
                        filename = filename.Remove(filename.Length - 2, 2);
                    }

                    if (filename.EndsWith(".pra"))
                    {
                        filename += ".plt";
                    }

                    if (i.PdfAble)
                    {
                        if (i.FileDateTime.Date == mostRecentDate)
                        {
                            if (Program.isSLC)
                            {
                                jobTicket.AppendLine("<JOB:DOCINPUT FILENAME=\"" + filename + "\" FOLDER=\"" + uncRawPrefix + i.FilePath.Remove(0, 10) + "\"/>");
                            }
                            else
                            {
                                jobTicket.AppendLine("<JOB:DOCINPUT FILENAME=\"" + filename + "\" FOLDER=\"" + uncRawPrefix + i.FilePath.Replace("/", "\\") + "\"/>");
                            }
                        }
                        else
                        {
                            jobTicket.AppendLine("<!-- SKIPPING(OLDER DATE): " + filename + ", " + i.FilePath.Replace("/", "\\") + " -->");
                        }
                    }
                    else
                    {
                        jobTicket.AppendLine("<!-- SKIPPING(NOT PDF-ABLE): " + filename + ", " + i.FilePath.Replace("/", "\\") + " -->");
                        Console.WriteLine("THIS IS NOT PDF-ABLE: " + filename + ", " + i.FilePath);
                    }
                }

                jobTicket.AppendLine("</JOB:DOCINPUTS>");
                jobTicket.AppendLine("<JOB:DOCOUTPUTS>");
                jobTicket.AppendLine("<JOB:DOCOUTPUT FILENAME=\"" + kvp.Key + ".pdf\" FOLDER=\"" + uncPdfPrefix + "\\\" DOCTYPE=\"PDF\" />");
                jobTicket.AppendLine("</JOB:DOCOUTPUTS>");
                jobTicket.AppendLine("<JOB:SETTINGS>");
                jobTicket.AppendLine("<JOB:PDFSETTINGS JPEGCOMPRESSIONLEVEL=\"5\" MONOIMAGECOMPRESSION=\"Default\" GRAYSCALE=\"No\" PAGECOMPRESSION=\"Yes\" DOWNSAMPLEIMAGES=\"No\" RESOLUTION=\"1200\" PDFVERSION=\"PDFVersion15\" PDFVERSIONINHERIT=\"No\" PAGES=\"All\" />");
                jobTicket.AppendLine("</JOB:SETTINGS>");
                jobTicket.AppendLine("</JOB>");
                jobTicket.AppendLine("</JOBS>");

                string jobFileName = jobTicketLocation + mostRecentDate.ToString("yyyy-MM-dd") + "_" + kvp.Key + ".xml";
                Console.WriteLine(jobFileName);
                File.WriteAllText(jobFileName, jobTicket.ToString());
            }
            File.WriteAllLines(batchFile, batchLines);
        }

        public static Hashtable LoadPdmCatalog()
        {
            Hashtable pdmHashTable = new Hashtable();

            //load Pdm Catalog File
            StreamReader sr = new StreamReader(catalogFile);
            string headerLine = sr.ReadLine();
            string catalogLine;

            while ((catalogLine = sr.ReadLine()) != null)
            {
                var pdmCatalogItem = new PdmItem();
                List<string> pdmCatalog = catalogLine.Split(',').ToList();

                pdmCatalogItem.Server = pdmCatalog[0];
                pdmCatalogItem.FileName = pdmCatalog[2];

                if (!pdmHashTable.ContainsKey(pdmCatalogItem.FileName))
                {
                    pdmHashTable.Add(pdmCatalogItem.FileName, null);
                }
            }
            return pdmHashTable;
        }

        static void Main(string[] args)
        {
            LoadConfig();

            Dictionary<string, List<PdmItem>> dictionary = new Dictionary<string, List<PdmItem>>();
            Hashtable pdmCatalogTable = LoadPdmCatalog();
            List<string> delimitedDataField = new List<string> { "FILE_SIZE,LAST_ACCESSED,ITEM,REV,SHEET,SERVER,UNC_RAW,UNC_PDF" };
            List<string> islandOfMisfitToys = new List<string>();
            List<string> batchLines = new List<string>();

            //parse extract file
            StreamReader file = new StreamReader(inputFile);
            string inputLine;

            while ((inputLine = file.ReadLine()) != null)
            {
                Console.WriteLine(inputLine);
                PdmItem pdmItem = new PdmItem();
                pdmItem.ParseInputLine(inputLine);

                if (pdmItem.FileDateTime < recentDateTime)
                {
                    continue;
                }

                if (!pdmCatalogTable.ContainsKey(pdmItem.FileName))
                {
                    pdmItem.IsMisfit = true;
                    islandOfMisfitToys.Add("Not in catalog: " + inputLine);
                    continue;
                }

                if (pdmItem.IsMisfit)
                {
                    islandOfMisfitToys.Add("Misfit: " + inputLine);
                }
                else
                {
                    delimitedDataField.Add(pdmItem.GetOutputLine());
                }

                //logic to handle no revs
                string uID;
                if (String.IsNullOrEmpty(pdmItem.ItemRev))
                {
                    uID = pdmItem.ItemName;
                }
                else
                {
                    uID = pdmItem.ItemName + "." + pdmItem.ItemRev;
                }

                if (!dictionary.Keys.Contains(uID))
                {
                    List<PdmItem> pdmItems = new List<PdmItem>();
                    pdmItems.Add(pdmItem);
                    dictionary.Add(uID, pdmItems);
                }
                else
                {
                    dictionary[uID].Add(pdmItem);
                }
            }

            //output all misfits to file
            File.WriteAllLines(misfitToys, islandOfMisfitToys);

            //Comment this next code until misfits are reviewed and corrected in source extract file
            //generate file for Graig
            File.WriteAllLines(outputFile, delimitedDataField);

            //generate XML job tickets
            JobTicketGenerator(dictionary, batchLines);
        }
    }
}

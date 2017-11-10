using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdmMigration
{
    public class PdmItem
    {
        public Int64 FileSize { get; set; }
        public string FileMonth { get; set; }
        public string FileDay { get; set; }
        public string FileYear { get; set; }
        public string FileTime { get; set; }
        public DateTime FileDateTime { get; set; }
        public string ItemName { get; set; }
        public string ItemRev { get; set; }
        public string ItemExt { get; set; }
        public string ItemSht { get; set; }
        public bool HasRev { get; set; }
        public bool HasSht { get; set; }
        public bool HasExt { get; set; }
        public string FileName { get; set; }     // ex: A12345.0.01.plt.Z
        public string FilePath { get; set; }     // ex: F:\archive\pdm\web\hpgl\A\
        public string FilePathName { get; set; } // ex: F:\archive\pdm\web\hpgl\A\A12345.0.01.plt.Z
        public string Server { get; set; }       // ex: pdm.moog.com
        public string UncRaw { get; set; }       // ex: \\eacmpnas01.moog.com\Vol5_Data\PDM\EA\archive\pdm\web\hpgl\A\A12345.0.01.plt
        public string UncPdf { get; set; }       // ex: \\eacmpnas01.moog.com\Vol5_Data\PDM\EA\tcpdf\A12345.0.pdf
        public bool IsMisfit { get; set; }
        public bool PdfAble { get; set; }
        public string PdfAbleFileName { get; set; }

        public void ParseInputLine(string line)
        {
            IsMisfit = false;
            PdfAble = false;
            HasRev = false;
            HasSht = false;
            HasExt = false;
            Server = Program.serverName;

            if (line.EndsWith(".ss"))
            {
                IsMisfit = true;
            }

            if (Program.isWindows)
            {
                //do parsing logic for windows
                List<string> windowsData = line.Split(' ').ToList();
                windowsData.RemoveAll(String.IsNullOrEmpty);

                FileSize = Convert.ToInt64(windowsData[0]);

                if (Program.isLuDateTime)
                {
                    FileDateTime = DateTime.ParseExact(windowsData[1] + ' ' + windowsData[2], "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    FilePathName = windowsData[3];
                }
                else if (Program.isIeDateTime)
                {
                    FileDateTime = DateTime.ParseExact(windowsData[1] + ' ' + windowsData[2], "dd/MM/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    FilePathName = windowsData[3];
                }
                else
                {
                    FileDateTime = Convert.ToDateTime(windowsData[1] + ' ' + windowsData[2] + ' ' + windowsData[3]);
                    FilePathName = windowsData[4];
                }

                int idx = FilePathName.LastIndexOf('\\');
                FilePath = FilePathName.Substring(2, idx + 1);
                FileName = FilePathName.Substring(idx + 1);

                PdfAbleFileName = FileName;
                if (FileName.EndsWith("._"))
                {
                    PdfAbleFileName = FileName.Remove(FileName.Length - 2, 2);
                }

                PdfAble = Program.IsPdfAble(PdfAbleFileName);

                string[] dataFileSplit = FilePathName.Split('.');

                int idx2 = dataFileSplit[0].LastIndexOf('\\');
                ItemName = dataFileSplit[0].Substring(idx2 + 1);

                if (dataFileSplit.Length == 2)
                {
                    HasExt = true;
                    ItemExt = dataFileSplit[1];

                    UncRaw = UncPaths.BuildUncRawPath(Program.uncRawPrefix, FilePathName);
                    UncPdf = UncPaths.BuildUncPdfPath(Program.uncPdfPrefix, ItemName, ItemRev);
                }
                else if (dataFileSplit.Length > 2)
                {
                    if (!Program.IsExt(dataFileSplit[1]))
                    {
                        HasRev = true;
                        ItemRev = dataFileSplit[1];
                    }
                    else
                    {
                        HasExt = true;
                        ItemExt = dataFileSplit[1];
                    }

                    if (!Program.IsExt(dataFileSplit[2]))
                    {
                        HasSht = true;
                        ItemSht = dataFileSplit[2];
                    }
                    else
                    {
                        HasExt = true;
                        ItemExt = dataFileSplit[2];
                    }

                    if (dataFileSplit.Length > 3 && Program.IsExt(dataFileSplit[3]))
                    {
                        HasExt = true;
                        ItemExt = dataFileSplit[3];
                    }

                    if (!HasRev && !HasSht && !HasExt)
                    {
                        IsMisfit = true;
                    }

                    UncRaw = UncPaths.BuildUncRawPath(Program.uncRawPrefix, FilePathName);
                    UncPdf = UncPaths.BuildUncPdfPath(Program.uncPdfPrefix, ItemName, ItemRev);
                }
                else
                {
                    IsMisfit = true;
                }
            }
            else
            {
                //parsing logic for linux
                //populate all data in PDMItem by parsing line with Linux rules
                List<string> linuxData = line.Split(' ').ToList();
                linuxData.RemoveAll(String.IsNullOrEmpty);
                linuxData.RemoveRange(0, 4);

                FileSize = Convert.ToInt64(linuxData[0]);
                FileMonth = linuxData[1];
                FileDay = linuxData[2];

                if (linuxData[3].Contains(":"))
                {
                    FileYear = "2017";
                    FileTime = linuxData[3];
                }
                else
                {
                    FileYear = linuxData[3];
                    FileTime = "00:00";
                }

                FileDateTime = Convert.ToDateTime(FileMonth + ' ' + FileDay + ' ' + FileYear + ' ' + FileTime);

                FilePathName = linuxData[4];

                int idx = FilePathName.LastIndexOf('/');
                FilePath = FilePathName.Substring(0, idx + 1);
                FileName = FilePathName.Substring(idx + 1);

                PdfAbleFileName = FileName;
                if (FileName.EndsWith(".Z"))
                {
                    PdfAbleFileName = FileName.Remove(FileName.Length - 2, 2);
                }

                PdfAble = Program.IsPdfAble(PdfAbleFileName);

                string[] linuxDataFileSplit = FilePathName.Split('.');

                int idx2 = linuxDataFileSplit[0].LastIndexOf('/');
                ItemName = linuxDataFileSplit[0].Substring(idx2 + 1);

                if (linuxDataFileSplit.Length == 2)
                {
                    HasExt = true;
                    ItemExt = linuxDataFileSplit[1];

                    UncRaw = UncPaths.BuildUncRawPath(Program.uncRawPrefix, FilePathName);
                    UncPdf = UncPaths.BuildUncPdfPath(Program.uncPdfPrefix, ItemName, ItemRev);
                }
                else if (linuxDataFileSplit.Length > 2)
                {
                    if (!Program.IsExt(linuxDataFileSplit[1]))
                    {
                        HasRev = true;
                        ItemRev = linuxDataFileSplit[1];
                    }
                    else
                    {
                        HasExt = true;
                        ItemExt = linuxDataFileSplit[1];
                    }

                    if (!Program.IsExt(linuxDataFileSplit[2]))
                    {
                        HasSht = true;
                        ItemSht = linuxDataFileSplit[2];
                    }
                    else
                    {
                        HasExt = true;
                        ItemExt = linuxDataFileSplit[2];
                    }

                    if (linuxDataFileSplit.Length > 3 && Program.IsExt(linuxDataFileSplit[3]))
                    {
                        HasExt = true;
                        ItemExt = linuxDataFileSplit[3];
                    }

                    if (!HasRev && !HasSht && !HasExt)
                    {
                        IsMisfit = true;
                    }

                    UncRaw = UncPaths.BuildUncRawPath(Program.uncRawPrefix, FilePathName);
                    UncPdf = UncPaths.BuildUncPdfPath(Program.uncPdfPrefix, ItemName, ItemRev);
                }
                else
                {
                    IsMisfit = true;
                }
            }
        }

        public string GetOutputLine()
        {
            //build and return full output line string using the data members above (for Graig's file)
            StringBuilder output = new StringBuilder(FileSize + "," + FileDateTime.ToString("MMM d yyyy HH:mm") + "," + ItemName);
            output.Append((HasRev) ? ("," + ItemRev) : (","));
            output.Append((HasSht) ? ("," + ItemSht) : (","));
            output.Append(("," + Server));
            output.Append("," + UncRaw + "," + UncPdf);

            return output.ToString();
        }
    }
}

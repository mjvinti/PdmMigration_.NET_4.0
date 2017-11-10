using System;
using System.Text;

namespace PdmMigration
{
    public class UncPaths
    {
        public static string BuildUncRawPath(string uncRawPrefix, string filePath)
        {
            StringBuilder uncRawPathName = new StringBuilder();

            if (filePath.EndsWith("._") || filePath.EndsWith(".Z"))
            {
                filePath = filePath.Remove(filePath.Length - 2);
            }

            uncRawPathName.Append(uncRawPrefix);

            if (filePath[1] == ':')
            {
                uncRawPathName.Append(filePath.Remove(0, 2));
            }
            else if (filePath.StartsWith("\\\\slctce02"))
            {
                uncRawPathName.Append(filePath.Remove(0, 10));
            }
            else
            {
                uncRawPathName.Append(filePath);
            }

            uncRawPathName.Replace('/', '\\');
            string uncRaw = uncRawPathName.ToString();

            return uncRaw;
        }

        public static string BuildUncPdfPath(string uncPdfPrefix, string itemName, string itemRev)
        {
            StringBuilder uncPdfPathName = new StringBuilder();

            uncPdfPathName.Append(uncPdfPrefix + "\\" + itemName);

            if (!String.IsNullOrEmpty(itemRev))
            {
                uncPdfPathName.Append("." + itemRev);
            }

            uncPdfPathName.Append(".pdf");
            uncPdfPathName.Replace('/', '\\');
            string uncPdf = uncPdfPathName.ToString();

            return uncPdf;
        }
    }
}

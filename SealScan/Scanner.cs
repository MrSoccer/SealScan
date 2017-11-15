using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SealScan
{
    class Scanner
    {
        public Scanner()
        {

        }

        internal static List<string> ScannedBarcodes(string fileName)
        {
            List<string> scannedBarcodes = new List<string>();
            string line;

            using (StreamReader sr = new StreamReader(fileName))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    char delimiterChars = ',';
                    string[] scanData = line.Split(delimiterChars);
                    scannedBarcodes.Add(scanData[3]);
                }
                sr.Close();
            }
            scannedBarcodes.Sort();
            List<string> sanitizedBarcodes = scannedBarcodes.Distinct().ToList();
            return sanitizedBarcodes;
        }

        internal static void ScannedBarcodes(DocClass docClass)
        {
            throw new NotImplementedException();
        }
    }
}

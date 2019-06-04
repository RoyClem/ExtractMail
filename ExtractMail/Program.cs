using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Interop.Outlook;
using Models;
using System.Linq;
using System.IO;

namespace ExtractMail
{
    class Program
    {
        static void Main(string[] args)
        {
            new Program().ParseMail("Crowbar");
        }

        public void ParseMail(string folder)
        {
            var outlookApplication = new Application();
            var outlookNamespace = outlookApplication.GetNamespace("MAPI");

            MAPIFolder inBox = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            MAPIFolder subFolders = inBox.Folders[folder];

            List<CrowbarModel> models = new List<CrowbarModel>();

            for (int i = 0; i < subFolders.Items.Count; i++)
            {
                MailItem itm = (MailItem)subFolders.Items[i + 1];
                var body = itm.Body;
                var cols = body.Split('\t');
                var ary = cols[1].Split(new char[] { '\r', '\n', ' ' });
                var numCrowbars = int.Parse(ary[2]);
                var dateCrowbar = DateTime.ParseExact(ary[7].Remove(ary[7].IndexOf("_PPLEXT.")), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");

                var model = new CrowbarModel();
                model.NumCrowbars = numCrowbars;
                model.ReportDate = DateTime.Parse(dateCrowbar);
                model.ReportDateStr = dateCrowbar;
                models.Add(model);

                for (int j = 9; j < cols.Length; j += 7)      // each line has 7 fields
                {
                    if (cols[j] == "")
                        break;

                    CrowbarRecord record = new CrowbarRecord();

                    record.Trace = cols[j].Remove(0, 2);     // remove \r\n
                    record.CID = cols[j + 1];
                    record.AmountStr = cols[j + 2];
                    record.Amount = decimal.Parse(cols[j + 2], NumberStyles.Currency);
                    record.Email = cols[j + 3];
                    record.Status = cols[j + 4];
                    record.LastUpdate = cols[j + 5];
                    record.Message = cols[j + 6];

                    model.CrowbarRecords.Add(record);
                }
            }

            WriteMail(models.OrderBy(m => m.ReportDate).ToList());
        }
        public void WriteMail(List<CrowbarModel> models)
        {
            string dir = @"C:\Projects\CrowbarPlot\CrowbarPlot\bin\";
            string serializationFile = Path.Combine(dir, "crowbars.bin");

            //serialize
            using (Stream stream = File.Open(serializationFile, FileMode.Create))
            {
                var bformatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();

                bformatter.Serialize(stream, models);
            }
        }

    }
}

using System.Collections.Generic;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;

namespace ImageConnect.Test.Func.Shared.Core.Helpers
{
    /// <summary>
    /// The csv helper.
    /// </summary>
    public static class CsvFileHelper
    {
        /// <summary>
        /// Writes data into csv.
        /// </summary>
        /// <typeparam name="T">Type of records.</typeparam>
        /// <param name="filePath">File path.</param>
        /// <param name="records">Values to write.</param>
        public static void WriteFile<T>(string filePath, List<T> records) where T : ClassMap<T>
        {
            using (var writer = new StreamWriter(filePath))
            {
                using (var csv = new CsvWriter(writer))
                {
                    csv.Configuration.RegisterClassMap<T>();
                    csv.WriteRecords(records);
                }
            }
        }
    }
}

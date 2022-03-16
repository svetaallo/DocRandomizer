using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace DocRandomizer
{
    class RandTable
    {
        private static Random random = new Random();
        private const string RandomizerInputDirectory = @"C:\DocRandomizer\частицы2022.docx";
        private const string RandomizerOutputDirectory = @"C:\DocRandomizer\";
        static RandTable()
        {
            if (!Directory.Exists(RandTable.RandomizerOutputDirectory))
            {
                Directory.CreateDirectory(RandTable.RandomizerOutputDirectory);
                var doc = DocX.Load(Program.DocumentDirectory + @"частицы2022.docx");
                doc.SaveAs(RandTable.RandomizerOutputDirectory + @"частицы2022.docx");
            }
        }
        public static void Do()
        {
            var document = DocX.Load(RandTable.RandomizerInputDirectory);
            var table = document.Tables[1];

            for (int i = 1; i < 6; i++)
            {
                table.Rows[i].Cells[3].Paragraphs[0].RemoveText(0);
                table.Rows[i].Cells[3].Paragraphs[0].Append(random.Next(90000, 210000).ToString());
                table.Rows[i].Cells[4].Paragraphs[0].RemoveText(0);
                table.Rows[i].Cells[4].Paragraphs[0].Append(random.Next(500, 1700).ToString());
            }
            document.SaveAs(RandTable.RandomizerOutputDirectory + @"частицы2022randomized.docx");
        }

    }
}

namespace Subtractive.Processor
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Text.RegularExpressions;

    /// <summary>
    /// エクセル処理プロセッサ
    /// </summary>
    public class ExcelProcessor : IProcessor
    {
        /// <summary>拡張子</summary>
        public string[] Extensions => new string[] { ".xlsx", ".xlsm" };

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public void Execute(string filePath)
        {
            this.QuantXlsx(filePath, "xl");
        }

        /// <summary>指定されたエクセルブックファイルに含まれる画像を圧縮します。</summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="ext">拡張子</param>
        internal void QuantXlsx(string filePath, string ext)
        {
            Regex regex = new Regex(string.Format(@"{0}/media/.*\.png", ext));

            Console.WriteLine("減色を開始します：{0}", filePath);

            string uncompressPath = Path.GetTempPath();

            ZipArchive archive = null;
            PngQuant pngquant = new PngQuant();
            string distinationFilePath = Path.GetDirectoryName(filePath);
            distinationFilePath = Path.Combine(distinationFilePath, "減色済" + Path.GetFileName(filePath));
            ZipArchive distinationArchive =
                new ZipArchive(new FileStream(distinationFilePath, FileMode.Create, FileAccess.Write), ZipArchiveMode.Create);
            try
            {
                archive = ZipFile.Open(filePath, ZipArchiveMode.Read);
                IEnumerable<ZipArchiveEntry> entries = archive.Entries;

                foreach (ZipArchiveEntry entry in entries)
                {
                    if (regex.IsMatch(entry.FullName.ToLower()) == false)
                    {
                        using (Stream stream = entry.Open())
                        {
                            ZipArchiveEntry e = distinationArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                            using (Stream s = e.Open())
                            {
                                byte[] bs = stream.ToByteArray((int)entry.Length);
                                s.Write(bs, 0, bs.Length);
                            }
                        }

                        continue;
                    }

                    // 解凍して一時ファイルを作成し、減色する
                    string entryFilePath = Path.GetTempFileName();
                    entryFilePath = pngquant.SubtractiveToTemporaryFile(entry);

                    // 減色したファイルをブックに再設定
                    distinationArchive.CreateEntryFromFile(entryFilePath, entry.FullName, CompressionLevel.Optimal);
                    Console.WriteLine("・{0}", entry.FullName);
                }
            }
            finally
            {
                distinationArchive.Dispose();

                if (archive != null)
                {
                    archive.Dispose();
                }

                pngquant.Dispose();
            }
        }
    }
}

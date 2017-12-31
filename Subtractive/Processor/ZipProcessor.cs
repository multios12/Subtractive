namespace Subtractive.Processor
{
    using System;
    using System.IO;
    using System.IO.Compression;
    using System.Text.RegularExpressions;

    /// <summary>
    /// ZIPファイル内のPNGファイルを減色して圧縮します
    /// </summary>
    public class ZipProcessor : IProcessor
    {
        /// <summary>ファイル名変更完了イベント</summary>
        public event EventHandler<QuantedEventArgs> FileNameChanged;

        /// <summary>拡張子</summary>
        public virtual string[] Extensions => new string[] { ".zip" };

        /// <summary>BMP/JPEGをPNGにコンバートする場合、True</summary>
        public virtual bool IsConvertToPng { get; set; } = true;

        /// <summary>出力したファイルのパス</summary>
        public virtual string OutputFilePath { get; set; }

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public virtual void Execute(string filePath)
        {
            this.QuantFromZipFile(filePath);
        }

        /// <summary>指定されたエクセルブックファイルに含まれる画像を圧縮します。</summary>
        /// <param name="sourceFilePath">ファイルパス</param>
        /// <param name="targetInnerPath">処理対象内部パス</param>
        internal void QuantFromZipFile(string sourceFilePath, string targetInnerPath = "")
        {
            string reg = this.IsConvertToPng ? @"{0}.*\.png|{0}.*\.jpeg|{0}.*\.jpg|{0}.*\.bmp" : @"{0}.*\.png";
            reg = string.Format(reg, targetInnerPath);
            Regex regex = new Regex(reg);

            Console.WriteLine("減色を開始します：{0}", sourceFilePath);

            string distinationFileName = string.Format(Properties.Settings.Default.FileNameTemplate, Path.GetFileName(sourceFilePath));
            string distinationFilePath = Path.GetDirectoryName(sourceFilePath);
            distinationFilePath = Path.Combine(distinationFilePath, distinationFileName);
            this.OutputFilePath = distinationFilePath;

            using (ZipArchive archive = ZipFile.Open(sourceFilePath, ZipArchiveMode.Read))
            using (ZipArchive distArchive = new ZipArchive(new FileStream(distinationFilePath, FileMode.Create, FileAccess.Write), ZipArchiveMode.Create))
            using (PngQuant pngquant = new PngQuant())
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (regex.IsMatch(entry.FullName.ToLower()) == false)
                    {
                        distArchive.CreateEntry(entry.FullName, entry);
                        continue;
                    }

                    // 解凍して一時ファイルを作成し、減色する
                    string entryFilePath = entry.SubtractiveToTemporaryFile(pngquant);

                    // 減色したファイルをブックに再設定
                    string distEntryName = PathUtils.ChangeExtension(entry.FullName, Path.GetExtension(entryFilePath));
                    distArchive.CreateEntryFromFile(entryFilePath, distEntryName, CompressionLevel.Optimal);
                    Console.WriteLine("・{0}", entry.FullName);

                    // 変換イベントを発生させる
                    this.FileNameChanged?.Invoke(this, new QuantedEventArgs(entry.FullName, distEntryName));
                }
            }
        }
    }
}

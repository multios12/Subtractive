namespace Subtractive.Processor
{
    using System.IO;
    using System.IO.Compression;
    using System.Xml;

    /// <summary>
    /// エクセルファイル内に含まれるPNG画像を減色します。
    /// </summary>
    public class ExcelProcessor : ZipProcessor
    {
        /// <summary>拡張子</summary>
        public override string[] Extensions => new string[] { ".xlsx", ".xlsm" };

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public override void Execute(string filePath)
        {
            this.QuantFromZipFile(filePath, "xl/media/");
        }

        /// <summary>
        /// TODO:ファイル内のリレーション情報を更新します。
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        private void updateRels(string filePath)
        {
            string rels = @"xl/drawings/_rels/drawing1.xml.rels";
            var document = new XmlDocument();
            using (var zipArchive = ZipFile.OpenRead(filePath))
            {
                ZipArchiveEntry e = zipArchive.GetEntry(rels);

                using (Stream stream = e.Open())
                using (StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8))
                {
                    document.LoadXml(reader.ReadToEnd());
                }
            }
        }
    }
}

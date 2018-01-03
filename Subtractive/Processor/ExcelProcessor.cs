namespace Subtractive.Processor
{
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// エクセルファイル内に含まれるPNG画像を減色します。
    /// </summary>
    public class ExcelProcessor : ZipProcessor
    {
        /// <summary>リレーション情報</summary>
        private XmlDocument document;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ExcelProcessor()
        {
            this.FileNameChanged += this.ExcelProcessor_FileNameChanged;
            this.IsConvertToPng = true;
        }

        /// <summary>拡張子</summary>
        public override string[] Extensions => new string[] { ".xlsx", ".xlsm" };

        /// <summary>リサイズ時の最大サイズ</summary>
        public new Size? ResizeSize { get; set; } = null;

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public override void Execute(string filePath)
        {
            if (this.IsConvertToPng)
            {
                using (var zipArchive = ZipFile.OpenRead(filePath))
                {
                    this.document = zipArchive.ReadXML(@"xl/drawings/_rels/drawing1.xml.rels", Encoding.UTF8);
                }
            }

            this.QuantFromZipFile(filePath, "xl/media/");
            if (this.IsConvertToPng)
            {
                using (var zipArchive = ZipFile.Open(this.OutputFilePath, ZipArchiveMode.Update))
                {
                    zipArchive.WriteXML(@"xl/drawings/_rels/drawing1.xml.rels", this.document);
                    zipArchive.WritePngContenttype();
                }
            }
        }

        /// <summary>
        /// ファイル名変更完了イベント
        /// </summary>
        /// <param name="sender">発生元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void ExcelProcessor_FileNameChanged(object sender, QuantedEventArgs e)
        {
            string target = string.Format("../media/{0}", Path.GetFileName(e.OldFileName));
            foreach (var element in this.document.GetElementsByTagName("Relationship").Cast<XmlElement>()
                .Where(v => v.Attributes["Target"].Value == target))
            {
                element.Attributes["Target"].Value = string.Format("../media/{0}", Path.GetFileName(e.NewFileName));
            }
        }
    }
}

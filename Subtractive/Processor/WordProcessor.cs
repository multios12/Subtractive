namespace Subtractive.Processor
{
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// ワード処理クラス
    /// </summary>
    public class WordProcessor : ZipProcessor
    {
        /// <summary>リレーション情報</summary>
        private XmlDocument document;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public WordProcessor()
        {
            this.FileNameChanged += this.WordProcessor_FileNameChanged;
            this.IsConvertToPng = true;
        }

        /// <summary>拡張子</summary>
        public override string[] Extensions => new string[] { ".docx", ".docm" };

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
                    var document = new XmlDocument();
                    document.LoadXml(zipArchive.ReadText(@"word/_rels/document.xml.rels", Encoding.UTF8));
                    this.document = document;
                }
            }

            this.QuantFromZipFile(filePath, "word/media/");
            if (this.IsConvertToPng)
            {
                using (var zipArchive = ZipFile.Open(this.OutputFilePath, ZipArchiveMode.Update))
                {
                    zipArchive.WriteXML(@"word/_rels/document.xml.rels", this.document);
                    zipArchive.WritePngContenttype();
                }
            }
        }

        /// <summary>
        /// ファイル名変更完了イベント
        /// </summary>
        /// <param name="sender">発生元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void WordProcessor_FileNameChanged(object sender, QuantedEventArgs e)
        {
            string target = string.Format("media/{0}", Path.GetFileName(e.OldFileName));
            foreach (var element in this.document.GetElementsByTagName("Relationship").Cast<XmlElement>()
                .Where(v => v.Attributes["Target"].Value == target))
            {
                element.Attributes["Target"].Value = string.Format("media/{0}", Path.GetFileName(e.NewFileName));
            }
        }
    }
}

namespace Subtractive.Processor
{
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

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public override void Execute(string filePath)
        {
            if (this.IsConvertToPng)
            {
                this.document = this.LoadRels(filePath);
            }

            this.QuantFromZipFile(filePath, "xl/media/");
            if (this.IsConvertToPng)
            {
                this.SaveRels(this.OutputFilePath, this.document);
            }
        }

        /// <summary>
        /// ファイル名変更完了イベント
        /// </summary>
        /// <param name="sender">発生元オブジェクト</param>
        /// <param name="e">イベントデータ</param>
        private void ExcelProcessor_FileNameChanged(object sender, FileNameChangedEventArgs e)
        {
            string target = string.Format("../media/{0}", Path.GetFileName(e.OldFileName));
            foreach (var element in this.document.GetElementsByTagName("Relationship").Cast<XmlElement>()
                .Where(v => v.Attributes["Target"].Value == target))
            {
                element.Attributes["Target"].Value = string.Format("../media/{0}", Path.GetFileName(e.NewFileName));
            }
        }

        /// <summary>
        /// 指定されたxlsx、xlsmファイル内の画像リレーション情報を取得します。
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        private XmlDocument LoadRels(string filePath)
        {
            using (var zipArchive = ZipFile.OpenRead(filePath))
            {
                var document = new XmlDocument();
                document.LoadXml(zipArchive.ReadText(@"xl/drawings/_rels/drawing1.xml.rels", Encoding.UTF8));
                return document;
            }
        }

        /// <summary>
        /// 指定されたファイルの画像リレーション情報を保存します。
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="document">リレーション情報</param>
        private void SaveRels(string filePath, XmlDocument document)
        {
            using (var zipArchive = ZipFile.Open(filePath, ZipArchiveMode.Update))
            {
                // リレーション情報の保存
                using (Stream stream = zipArchive.GetEntry(@"xl/drawings/_rels/drawing1.xml.rels").Open())
                using (XmlWriter writer = XmlWriter.Create(stream))
                {
                    document.WriteTo(writer);
                }

                // Content-Typeにpngを追加
                XmlDocument content = new XmlDocument();
                content.LoadXml(zipArchive.ReadText(@"[Content_Types].xml", Encoding.UTF8));
                var l = content.GetElementsByTagName("Default").Cast<XmlElement>().Where(e => e.Attributes["Extension"].Value == "png");
                if (l.Count() == 0)
                {
                    XmlElement element = content.CreateElement("Default", content.GetElementsByTagName("Types")[0].NamespaceURI);
                    element.SetAttribute("Extension", "png");
                    element.SetAttribute("ContentType", "image/png");
                    content.GetElementsByTagName("Types")[0].AppendChild(element);

                    zipArchive.WriteXML(@"[Content_Types].xml", content);
                }
            }
        }
    }
}

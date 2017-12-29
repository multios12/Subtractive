namespace Subtractive
{
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// ZipArchive拡張クラス
    /// </summary>
    public static class ZipArchiveExtensions
    {

        /// <summary>
        /// 指定されたエントリを指定されたエントリパスに保存します。
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        /// <param name="entryPath">エントリパス</param>
        /// <param name="entry">保存するエントリ</param>
        public static void CreateEntry(this ZipArchive zipArchive, string entryPath, ZipArchiveEntry entry)
        {
            using (Stream stream = entry.Open())
            using (Stream s = zipArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal).Open())
            {
                byte[] bs = stream.ToByteArray((int)entry.Length);
                s.Write(bs, 0, bs.Length);
            }
        }

        /// <summary>
        /// アーカイブファイルから、指定されたエントリを読み取り、テキストとして返します。
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        /// <param name="entryPath">読み取るエントリ</param>
        /// <param name="encoding">エンコーディング</param>
        /// <returns>テキスト</returns>
        public static string ReadText(this ZipArchive zipArchive, string entryPath, Encoding encoding)
        {
            using (Stream stream = zipArchive.GetEntry(entryPath).Open())
            using (StreamReader reader = new StreamReader(stream, encoding))
            {
                return reader.ReadToEnd();
            }
        }

        /// <summary>
        /// アーカイブファイルから、指定されたエントリを読み取り、イメージとして返します。
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        /// <param name="entryPath">読み取るエントリ</param>
        public static Image ReadImage(this ZipArchive zipArchive, string entryPath)
        {
            using (Stream stream = zipArchive.GetEntry(entryPath).Open())
            {
                return Image.FromStream(stream);
            }
        }

        /// <summary>
        /// アーカイブファイルから、指定されたエントリを読み取り、XMLとして返します。
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        /// <param name="entryPath">読み取るエントリ</param>
        /// <param name="encoding">エンコーディング</param>
        /// <returns>XML</returns>
        public static XmlDocument ReadXML(this ZipArchive zipArchive, string entryPath, Encoding encoding)
        {
            var document = new XmlDocument();
            document.LoadXml(zipArchive.ReadText(@"xl/drawings/_rels/drawing1.xml.rels", Encoding.UTF8));
            return document;
        }

        /// <summary>
        /// 指定されたエントリーパスに、XMLドキュメントを上書きします。
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        /// <param name="entryPath">エントリパス</param>
        /// <param name="document">XMLドキュメント</param>
        public static void WriteXML(this ZipArchive zipArchive, string entryPath, XmlDocument document)
        {
            using (Stream stream = zipArchive.GetEntry(entryPath).Open())
            using (XmlWriter writer = XmlWriter.Create(stream))
            {
                document.WriteTo(writer);
            }
        }

        /// <summary>
        /// OOXML形式のアーカイブファイルのContentTypeを更新します
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        public static void WritePngContenttype(this ZipArchive zipArchive)
        {
            if (zipArchive.GetEntry(@"[Content_Types].xml") == null)
            {
                throw new FileNotFoundException("ファイルにコンテントタイプが存在しません。", @"[Content_Types].xml");
            }

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

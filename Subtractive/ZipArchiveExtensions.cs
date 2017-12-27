namespace Subtractive
{
    using System.IO;
    using System.IO.Compression;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// ZipArchive拡張クラス
    /// </summary>
    public static class ZipArchiveExtensions
    {
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
        /// 指定されたエントリーパスに、XMLドキュメントを上書きします。
        /// </summary>
        /// <param name="zipArchive">アーカイブファイル</param>
        /// <param name="entryPath">エントリパス</param>
        /// <param name="document">XMLドキュメント</param>
        public static void WriteXML(this ZipArchive zipArchive, string entryPath, XmlDocument document)
        {
            using (Stream stream = zipArchive.GetEntry(@"[Content_Types].xml").Open())
            using (XmlWriter writer = XmlWriter.Create(stream))
            {
                document.WriteTo(writer);
            }
        }
    }
}

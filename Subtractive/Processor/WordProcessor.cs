namespace Subtractive.Processor
{
    /// <summary>
    /// パワーポイント処理クラス
    /// </summary>
    public class WordProcessor : ZipProcessor
    {
        /// <summary>拡張子</summary>
        public override string[] Extensions => new string[] { ".docx" };

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public override void Execute(string filePath)
        {
            this.QuantFromZipFile(filePath, "word/media/");
        }
    }
}

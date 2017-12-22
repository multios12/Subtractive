namespace Subtractive.Processor
{
    /// <summary>
    /// パワーポイント処理クラス
    /// </summary>
    public class PowerpointProcessor : ExcelProcessor
    {
        /// <summary>拡張子</summary>
        public new string[] Extensions => new string[] { ".pptx" };

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public new void Execute(string filePath)
        {
            this.QuantXlsx(filePath, "ppt");
        }
    }
}

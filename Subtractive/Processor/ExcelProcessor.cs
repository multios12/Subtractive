namespace Subtractive.Processor
{
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
    }
}

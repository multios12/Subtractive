namespace Subtractive.Processor
{
    using System.Drawing;

    /// <summary>
    /// パワーポイント処理クラス
    /// </summary>
    public class PowerpointProcessor : ZipProcessor
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public PowerpointProcessor()
        {
            this.IsConvertToPng = false;
        }

        /// <summary>拡張子</summary>
        public override string[] Extensions => new string[] { ".pptx" };

        /// <summary>リサイズ時の最大サイズ</summary>
        public new Size? ResizeSize { get; set; } = null;

        /// <summary>
        /// 処理を実行します
        /// </summary>
        /// <param name="filePath">処理するファイルのパス</param>
        public override void Execute(string filePath)
        {
            this.QuantFromZipFile(filePath, "ppt/media/");
        }
    }
}

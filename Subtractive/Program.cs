namespace Subtractive
{
    using System;
    using System.IO;
    using Subtractive.Processor;

    /// <summary>
    /// プログラム
    /// </summary>
    public class Program
    {
        /// <summary>プログラムのスタートアップポイント</summary>
        /// <param name="args">コマンドライン引数</param>
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine(Properties.Resources.CopyRight + "ファイルパスを指定してください");
                return;
            }

            string filePath = args[0];
            if (File.Exists(filePath) == false)
            {
                Console.WriteLine(Properties.Resources.CopyRight + "ファイルが見つかりません");
                return;
            }

            IProcessor processor = ProcessorFactory.CreateInstance(filePath);
            processor.Execute(filePath);
        }
    }
}

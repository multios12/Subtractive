namespace Subtractive
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using Subtractive.Processor;

    /// <summary>
    /// プログラム
    /// </summary>
    public class Program
    {
        /// <summary>
        /// プログラムのスタートアップポイント
        /// </summary>
        /// <param name="args">コマンドライン引数</param>
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine(Properties.Resources.CopyRight + Environment.NewLine
                    + "ファイルパスを指定してください");

                return;
            }

            string filePath = args[0];
            if (File.Exists(filePath) == false)
            {
                Console.WriteLine(Properties.Resources.CopyRight + Environment.NewLine
                    + "ファイルが見つかりません");
                return;
            }

            string extension = Path.GetExtension(filePath.ToLower());
            foreach (IProcessor t in GetInterfaces<IProcessor>().Select(c => Activator.CreateInstance(c) as IProcessor))
            {
                if (t.Extensions.Contains(extension))
                {
                    t.Execute(filePath);
                    return;
                }
            }

            Console.WriteLine(Properties.Resources.CopyRight + Environment.NewLine + "このファイルには対応していません。");
        }

        /// <summary>
        /// 現在実行中のコードを格納しているアセンブリ内の指定されたインターフェイスが実装されているすべての Type を返します
        /// </summary>
        private static Type[] GetInterfaces<T>() => Assembly.GetExecutingAssembly().GetTypes().Where(c => c.GetInterfaces().Any(t => t == typeof(T))).ToArray();
    }
}

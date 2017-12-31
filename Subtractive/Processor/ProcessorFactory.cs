namespace Subtractive.Processor
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    /// <summary>
    /// プロセッサ作成クラス
    /// </summary>
    public class ProcessorFactory
    {
        /// <summary>
        /// 指定されたファイルを操作できるプロセッサを作成してインスタンスを返します。
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>インスタンス、作成できない場合null</returns>
        public static IProcessor CreateInstance(string filePath)
        {
            string extension = Path.GetExtension(filePath.ToLower());
            foreach (IProcessor t in GetInterfaces<IProcessor>().Select(c => Activator.CreateInstance(c) as IProcessor))
            {
                if (t.Extensions.Contains(extension))
                {
                    return t;
                }
            }
            return new NullProcessor();
        }

        /// <summary>
        /// 現在実行中のコードを格納しているアセンブリ内の指定されたインターフェイスが実装されているすべての Type を返します
        /// </summary>
        private static Type[] GetInterfaces<T>() =>  Assembly.GetExecutingAssembly().GetTypes().Where(c => c.GetInterfaces().Any(t => t == typeof(T))).ToArray();
    }
}

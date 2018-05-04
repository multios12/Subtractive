using System;
using System.Diagnostics;
using System.IO;

/// <summary>
/// pngquantプロセスの作成及び、アクセスを行います。
/// </summary>
public class PngquantProcess : Process
{
    /// <summary>
    /// 指定された画像ファイルを減色する
    /// </summary>
    /// <param name="sourceStream">圧縮元PNG画像のストリーム</param>
    /// <returns>圧縮が完了したストリーム</returns>
    public Stream Execute(Stream sourceStream)
    {
        try
        {
            var info = new ProcessStartInfo("Pngquant.exe", " --force --verbose --ext=.png --speed=1 --ordered 256 {0}")
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardInput = true
            };

            this.StartInfo = info;
            this.Start();

            using (Stream s = this.StandardInput.BaseStream)
            {
                sourceStream.CopyTo(s);
            }

            using (var s = this.StandardOutput.BaseStream)
            {
                var distStream = new MemoryStream();
                s.CopyTo(distStream);
                distStream.Position = 0;
                return distStream;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("■Archive側のエラー");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("■resizer側のエラー");
            var v = this.StandardError.ReadToEnd();
            Console.WriteLine(v);
            Console.WriteLine("-----------------------------------------------");
            throw ex;
        }
    }
}

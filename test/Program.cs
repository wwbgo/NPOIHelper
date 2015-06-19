using System;
using System.Diagnostics;

namespace test {
    class Program {
        static void Main1(string[] args) {
            var path = String.Format("\"E:\\数据中心\\Source-分支-2015.04.13\\OnlyEdu.DCS\\lib\\Tools\\pdf2swf.exe\"" +
                                     " -t \"{0}\" -s flashversion=9 -s disablelinks -o \"{1}\"",
                                     "E:\\test\\3.pdf","E:\\test\\3.swf");
            var psi = new ProcessStartInfo(path);
            Console.WriteLine(path);
            //psi.CreateNoWindow = true;
            psi.UseShellExecute = false;
            psi.RedirectStandardError = true;
            //psi.RedirectStandardOutput = true;
            using (var pc = new Process()) {
                pc.StartInfo = psi;
                pc.Start();
                //var a = pc.StandardOutput;
                var b = pc.StandardError;
                //Console.WriteLine(a.ReadToEnd());
                Console.WriteLine(b.ReadToEnd());
            }
            Console.WriteLine("end");
            Console.Read();
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FontInstallerService
{
    class Program
    {
        [DllImport("gdi32.dll", EntryPoint = "AddFontResourceW", SetLastError = true)]
        public static extern int AddFontRessource([In][MarshalAs(UnmanagedType.LPWStr)] string lpFileName);

        static void Main(string[] args)
        {
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = @"C:\tmp\NewFonts";
            //watcher.NotifyFilter = NotifyFilters.LastWrite;
            //watcher.Filter = "*.*";
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.EnableRaisingEvents = true;

            Console.WriteLine("Watcher ready...");

            // wait - not to end
            new System.Threading.AutoResetEvent(false).WaitOne();
        }

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            Console.WriteLine($"File detected: {e.FullPath} ({e.ChangeType})");
            var fi = new FileInfo(e.FullPath);
            if(fi.Extension == ".ttf")
            {
                Console.WriteLine("File is a font! Installing font...");
                var res = installFont(fi);
                Console.WriteLine($"Installation result: Result: {res.Result}, ErrorCode: {res.Error}");
            }
        }

        private static InstallFontResult installFont(FileInfo fileInfo)
        {
            int result = -1;
            int error = 0;
            try
            {
                string fontPath = Path.Combine(System.Environment.GetEnvironmentVariable("windir"), "Fonts");
                Console.WriteLine($"Font Folder: {fontPath}");
                string destFile = Path.Combine(fontPath, fileInfo.Name);
                Console.WriteLine($"Target FullFileName: {destFile}");

                Console.WriteLine("Copy file to folder...");
                File.Copy(fileInfo.FullName, destFile, true);
                
                result = AddFontRessource(destFile);
                error = Marshal.GetLastWin32Error();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                error = -1;
            }
            return new InstallFontResult { Result = result, Error = error };
        }
    }

    public class InstallFontResult
    {
        public int Result { get; set; }
        public int Error { get; set; }
    }
}

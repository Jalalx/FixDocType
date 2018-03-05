using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FixDocType
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Error("Enter MS Word Document file path.");
                return;
            }

            var fileInfo = new FileInfo(args[0]);
            if (fileInfo.Extension.ToLower() != ".docx")
            {
                Error("Only MS Word document with .docx extension are supported.");
                return;
            }

            var content = File.ReadAllBytes(fileInfo.FullName);
            var serializer = new DocumentSerializer(content);
            serializer.Fix();

            var newFileName = fileInfo.FullName.Replace(fileInfo.Extension, "_fixed" + fileInfo.Extension);

            File.WriteAllBytes(newFileName, serializer.ToBytes());
            Console.WriteLine("MS Word file '{0}' fixed successfully.", newFileName);
            Console.ReadKey();
        }

        public static void Error(string error, params object[] args)
        {
            var previousColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(error, args);
            Console.ForegroundColor = previousColor;
        }
    }
}

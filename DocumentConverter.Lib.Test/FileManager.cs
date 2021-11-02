using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DocumentConverter.Lib.Test
{
    public static class FileManager
    {
        public static string _directory => Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static Stream LoadFile(string fileName)
        {
            var path = GetPath(fileName);
            return new FileStream(path, FileMode.Open);
        }
        public static string GetPath(string fileName)
        {
            var path = Path.Combine(_directory + $"\\Data\\", fileName);
            return path;
        }
        public static byte[] ReadFile(string fileName)
        {
            var path = Path.Combine(_directory + $"\\Data\\", fileName);
            return File.ReadAllBytes(path);
        }
        public static string WriteToFile(string filename, byte[] data,string folderName= "cloudmersive")
        {
            string path = FilePathToWriteTo(filename, folderName);
            using (Stream file = File.OpenWrite(path))
            {
                file.Write(data, 0, data.Length);
            }
            return path;
        }

        public static string FilePathToWriteTo(string filename, string folderName)
        {
            CreateCloudMersiveDirectory(folderName);
            string path = $@"{_directory}\{folderName}\{filename}";
            return path;
        }

        private static void CreateCloudMersiveDirectory(string folderName)
        {
            if (!Directory.Exists($@"{_directory}\{folderName}"))
            {
                Directory.CreateDirectory($@"{_directory}\{folderName}");
            }
        }

        public static void DownloadFile(string url)
        {
            if (!Directory.Exists($@"{_directory}\cloudmersive\images"))
            {
                Directory.CreateDirectory($@"{_directory}\cloudmersive\images");
            }

            using (WebClient client = new WebClient())
            {
                string fileName = $@"{_directory}\cloudmersive\images\{url.Split('/').ToList().Last()}";
                client.DownloadFile(url, fileName);
            }
        }

        public static string ReadFileAsString(string fileName)
        {
            string path = Path.Combine(_directory + $"\\Data\\", fileName);

            try
            {
                var sstreamReader = new StreamReader(path, Encoding.UTF8);

                using (var streamReader = new StreamReader(path, Encoding.UTF8))
                {
                    return streamReader.ReadToEnd();
                }
            }
            catch (Exception e)
            {

                throw e;
            }
        }
    }
}

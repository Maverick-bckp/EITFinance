using System;
using System.IO;

namespace EITFinance.Utilities
{
    public static class FileHelper
    {
        public static string[] GetFiles(string folderPath)
        {
            try
            {
                return Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories);
            }
            catch
            {
                return Array.Empty<string>();
            }
        }
        public static bool CopyFile(string sourcPath, string tagetPath)
        {
            try
            {
                string fileName = Path.GetFileName(sourcPath);
                string path = Path.GetDirectoryName(sourcPath);
                string relativePath = path.Replace(path, ""); if (!Directory.Exists(tagetPath + "/" + relativePath))
                {
                    Directory.CreateDirectory(tagetPath + "/" + relativePath);
                }
                File.Copy(sourcPath, tagetPath + "/" + relativePath + "/" + fileName, true); return true;
            }
            catch
            {
                return false;
            }
        }
        public static string GetSignature(string templateName)
        {
            try
            {
                using (StreamReader streamReader = new StreamReader(Path.Combine("wwwroot/Templates", templateName)))
                {
                    return streamReader.ReadToEnd();
                }
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}

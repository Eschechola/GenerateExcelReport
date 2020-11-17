using System.IO;

namespace GenerateExcel.Utilities
{
    public static class FilerSave
    {
        public static void Save(string path, byte[] bytes)
        {
            File.WriteAllBytes(path, bytes);
        }
    }
}

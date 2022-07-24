using System.Drawing;
using System.Text.RegularExpressions;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

class Program
{
    const int DefaultWidth = 800;
    const int Margin = 5;
    const int LineItemCount = 3;

    const int MaxItemCount = LineItemCount * 9 + 1;

    const int DefaultSavedImageWidth = 1280;
    const int DefaultSavedImageHeight = 720;

    static Brush s_backgroundBrush = Brushes.Azure;

    static void Main(string[] args)
    {
        try
        {
            Runs(args);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            Console.ReadLine();
        }
    }

    static void Runs(string[] args)
    {
        foreach (var item in args)
        {
            try
            {
                Console.WriteLine($"开始处理 {item}");
                Run(item);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理 {item} 失败 : {ex}");
            }
        }
    }

    static void Run(string path)
    {
        if (Directory.Exists(path))
        {
            Runs(Directory.EnumerateDirectories(path).ToArray());
            Runs(Directory.EnumerateFiles(path).ToArray());
        }
        else
        {
            var extension = Path.GetExtension(path);
            if (string.Equals(".ppt", extension, StringComparison.OrdinalIgnoreCase)
                || string.Equals(".pptx", extension, StringComparison.OrdinalIgnoreCase))
            {
                Generate(path, Path.ChangeExtension(path, ".jpg"));
            }
            else
            {
                Console.WriteLine($"{path} 不是 PowerPoint 文件");
            }
        }
    }

    static void Generate(string file, string output)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"PPTG-{Guid.NewGuid():n}");
        var width = DefaultWidth;

        try
        {
            Directory.CreateDirectory(tempDir);

            SaveAsPictures(file, tempDir);

            var pictures = GetOrderedPictures(tempDir);
            using var cover = Image.FromFile(pictures[0]);

            var percentage = (width - Margin * 2) / cover.Width * 1.0;

            var x = Margin;
            var y = Margin * 2;
            var lineCount = (int)(pictures.Length - 1) / LineItemCount + ((pictures.Length - 1) % LineItemCount > 0 ? 1 : 0);
            y += lineCount * Margin;

            var coverWidth = width - Margin * 2;
            var coverHeight = (int)(cover.Height * (coverWidth * 1.0 / cover.Width));
            y += coverHeight;

            var itemWidth = (int)Math.Round((width - Margin * (LineItemCount + 1)) / LineItemCount * 1.0);
            var itemHeight = (int)(DefaultSavedImageHeight * (itemWidth * 1.0 / DefaultSavedImageWidth));
            y += itemHeight * lineCount;

            using var targetImage = new Bitmap(width, y);
            using var g = Graphics.FromImage(targetImage);

            g.FillRectangle(s_backgroundBrush, 0, 0, width, y);

            x = Margin;
            y = Margin;

            g.DrawImage(cover, x, y, coverWidth, coverHeight);
            y += coverHeight + Margin;

            var lineItemSum = 0;
            foreach (var item in pictures.Skip(1))
            {
                using var img = Image.FromFile(item);
                g.DrawImage(img, x, y, itemWidth, itemHeight);
                x += Margin + itemWidth;
                if (++lineItemSum == LineItemCount)
                {
                    lineItemSum = 0;
                    x = Margin;
                    y += Margin + itemHeight;
                }
            }
            targetImage.Save(output);
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    static void SaveAsPictures(string file, string directory)
    {
        {
            var application = new ApplicationClass();

            var presentation = application.Presentations.Open(file, ReadOnly: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);

            presentation.SaveAs(directory, PpSaveAsFileType.ppSaveAsJPG);

            presentation.Close();

            application.Quit();
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    static string[] GetOrderedPictures(string directory, string fileFilter = "*.jpg")
    {
        return Directory.EnumerateFiles(directory, fileFilter).OrderBy(GetFileId).Take(MaxItemCount).ToArray();

        static int GetFileId(string filePath)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            return int.Parse(Regex.Match(fileName, "\\d+").Value);
        }
    }
}
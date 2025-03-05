using System;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System.Drawing;
using DocXToPdfConverter;

class Program
{
    static void Main(string[] args)
    {

        List<string> folderPaths = new List<string> { };

        if (args.Count() == 0)
        {
            return;
        }

        //string executableLocation = Path.GetDirectoryName("C:\\Users\\lenovo\\Desktop\\wordFiles\\");
        string executableLocation = Path.GetDirectoryName(args[0]);
        folderPaths.Add(executableLocation);

        // For loop on folder paths



        while (folderPaths.Count != 0)
        {
            string lastElement = folderPaths[folderPaths.Count - 1];
            folderPaths.RemoveAt(folderPaths.Count - 1);

            foreach (string file in Directory.GetFiles(lastElement))
            {
                string filePath = Path.GetFullPath(file);
                string extension = Path.GetExtension(filePath);
                Console.WriteLine($"{filePath} {extension}");
                switch (extension)
                {
                    //hexaware@01##

                    case ".docx":
                    //Docs
                        ConvertDocsToPdf(lastElement, filePath);
                        break;
                    case ".txt":
                        //Text
                        filePath = PathConverter(filePath);
                        TextToPdf(filePath);
                        break;
                    default:
                        Console.WriteLine("File Not Supported");
                        break;
                }
            }
            foreach (string dir in Directory.GetDirectories(lastElement))
            {
                Console.WriteLine(Path.GetFullPath(dir));
                folderPaths.Add(PathConverter(dir));

            }
            Console.WriteLine("Last Element" + lastElement);
        }


    }

    static string PathConverter(string path)
    {
        string modifiedPath = path.Replace(@"\", @"\\");
        return modifiedPath;
    }
    static void ConvertDocsToPdf(string folderLoc, string filePath)
    {
        string docxLocation = Path.Combine(folderLoc, Path.GetFileName(filePath));
        //string docxLocation = Path.Combine(executableLocation, "Offer Letter.docx");
        string locationOfLibreOfficeSoffice = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";
        var test = new ReportGenerator(locationOfLibreOfficeSoffice);
        test.Convert(docxLocation, Path.Combine(Path.GetDirectoryName(docxLocation), $"{Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath))}.pdf"));
        Console.WriteLine("File Converted To PDF");

    }
    static void TextToPdf(string filePath)
    {
        // text to pdf 
        PdfDocument doc = new PdfDocument();
        PdfPageBase page = doc.Pages.Add();
        PdfSolidBrush brushBlack = new PdfSolidBrush(new PdfRGBColor(System.Drawing.Color.Black));
        PdfTrueTypeFont paraFont = new PdfTrueTypeFont(new Font("Times New Roman", 12f, FontStyle.Regular), true);
        PdfStringFormat format1 = new PdfStringFormat();
        format1.Alignment = PdfTextAlignment.Center;
        string bodyText = File.ReadAllText(filePath);
        PdfTextWidget widget = new PdfTextWidget(bodyText, paraFont, brushBlack);
        Rectangle rect = new Rectangle(0, 100, (int)page.Canvas.ClientSize.Width, (int)page.Canvas.ClientSize.Height);
        PdfTextLayout layout = new PdfTextLayout();
        layout.Layout = PdfLayoutType.Paginate;
        widget.Draw(page, rect, layout);
        filePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath));
        doc.SaveToFile($"{filePath}.pdf");
    }
}
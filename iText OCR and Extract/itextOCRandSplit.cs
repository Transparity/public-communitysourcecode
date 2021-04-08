using System;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Xml;
using iText.IO.Source;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Utils;
using iText.Pdfocr;
using iText.Pdfocr.Tesseract4;
using Ghostscript.NET;
using Azure.Storage.Files.Shares;
using Azure.Storage.Files.Shares.Models;
using SP = Microsoft.SharePoint.Client;
using System.Security;
using System.Xml.Linq;

// Using iText7 - https://itextpdf.com/en 
// Using Ghostscript - https://www.ghostscript.com/
// using GhostScript.NET - https://archive.codeplex.com/?p=ghostscriptnet (Josip Habjan)

public static class Program
{

    private static readonly Tesseract4OcrEngineProperties tesseract4OcrEngineProperties = new Tesseract4OcrEngineProperties();
    private static List<FileInfo> LIST_IMAGES_OCR = new List<FileInfo>{};
    private static string SiteUrl = "";
    private static string DocumentLibrary = "";
    private static string ExtractFilePath = "";
    private static string UserName = "";
    private static string Password = "";
    private static string SiteRelativeUrl = "";
    private static string FileDownLoadPath = "";
    private static string ConnectionString = "";
    private static string ShareName = "";
    private static string TesseractFilePath = "";
    private static string ExtractPhrasefile = "";
    private static string NoMatchFilePath = "";
    private static string LogFilePath = "";
    private static int FileCount = 0;
    private static int ExtractFileCount = 0;
    private static string logfilename = "";
    private static string ExtractOnly = "N";

    static void Main(string[] args)
    {
        // get config settings
        SiteUrl = ConfigurationManager.AppSettings["SyntexSiteUrl"];
        DocumentLibrary = ConfigurationManager.AppSettings["SyntexDocumentLibraryTitle"];
        SiteRelativeUrl = ConfigurationManager.AppSettings["SyntexSiteRelativeUrl"];

        ExtractFilePath = ConfigurationManager.AppSettings["ExtractFilePath"];
        FileDownLoadPath = ConfigurationManager.AppSettings["DownLoadFilePath"];
        TesseractFilePath = ConfigurationManager.AppSettings["TesseractFilePath"];
        ExtractPhrasefile = ConfigurationManager.AppSettings["ExtractPhraseFile"];
        NoMatchFilePath = ConfigurationManager.AppSettings["NoMatchFilePath"];
        LogFilePath = ConfigurationManager.AppSettings["LogFilePath"];

        UserName = ConfigurationManager.AppSettings["UserName"];
        Password = ConfigurationManager.AppSettings["Password"];

        ConnectionString = ConfigurationManager.AppSettings["AzureFileConnectionString"];
        ShareName = ConfigurationManager.AppSettings["AzureFileShareName"];
        FileCount = Convert.ToInt32(ConfigurationManager.AppSettings["TestingFileCount"]);
        ExtractFileCount = Convert.ToInt32(ConfigurationManager.AppSettings["ExtractTestingFileCount"]);
        ExtractOnly = ConfigurationManager.AppSettings["ExtractOnly"];
        logfilename = String.Format("{0}{1}{2}{3}", LogFilePath, "OCR_log_", DateTime.Now.ToLongDateString(), ".txt");

        try
        {
            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log("Start OCR Extract run", w);
            }

            // download from Azure file share and OCR
            if (ExtractOnly == "N")
            {
                ProcessOCRFiles();
            }

            // read OCR files and extract portions of documents to new PDFs
            ProcessExtractFiles();

            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log("End OCR Extract run", w);
            }
        } catch (Exception ex) { Console.WriteLine(ex.Message); }
    } 
    private static void UploadFileToSharePoint(string siteUrl, string siteRelativeUrl, string docLibrary, string fileName, string login, string password, string pageCount)
    {
        try
        {
            var securePassword = new SecureString();
            foreach (char c in Password)
            { securePassword.AppendChar(c); }
            var spoCredentials = new SP.SharePointOnlineCredentials(login, securePassword);
            using (SP.ClientContext ctx = new SP.ClientContext(siteUrl))
            {
                ctx.Credentials = spoCredentials;
                SP.Web web = ctx.Web;
                SP.FileCreationInformation newFile = new SP.FileCreationInformation();
                byte[] FileContent = System.IO.File.ReadAllBytes(fileName);
                newFile.ContentStream = new MemoryStream(FileContent);
                newFile.Url = System.IO.Path.GetFileName(fileName);
                SP.List library = web.Lists.GetByTitle(docLibrary);

                SP.Folder Clientfolder = library.RootFolder; //.Folders.Add(ClientSubFolder);

                SP.File uploadFile = library.RootFolder.Files.Add(newFile);

                ctx.Load(library);
                ctx.Load(uploadFile);
                ctx.Load(uploadFile.ListItemAllFields);
                ctx.ExecuteQuery();


                // Set document properties
                var targetFileUrl = String.Format("{0}{1}", siteRelativeUrl, System.IO.Path.GetFileName(fileName));
                var uploadedFile = ctx.Web.GetFileByServerRelativeUrl(targetFileUrl);
                var listItem = uploadedFile.ListItemAllFields;
                listItem["PageCount"] = pageCount;
                listItem.Update();
                ctx.ExecuteQuery();


            }


        }
        catch (Exception exp)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(exp.Message + Environment.NewLine + exp.StackTrace);
        } 
        finally
        {
            Console.ReadLine();
        }
    }

    public static void ProcessExtractFiles()
    {
        using (StreamWriter w = File.AppendText(logfilename))
        {
            Log("Start ProcessExtractFiles", w);
        }
        try
        {
            string xmlconfigdoc = ExtractPhrasefile;
            var xml = XDocument.Load(xmlconfigdoc);
            IEnumerable<XElement> phrases = xml.Root.Elements();

            string[] files = Directory.GetFiles(String.Format("{0}{1}", ExtractFilePath, @"\ocr\"));
            string start2;
            string end2;
            string start1;
            string end1;
            int fileno = 0;
            bool match = false;
            foreach (string ocrfile in files)
            {
                fileno = 0;
                match = false;
                foreach (XElement phrase in phrases)
                {
                    fileno++;
                    start1 = (string)phrase.Attribute("start1");
                    end1 = (string)phrase.Attribute("end1");
                    start2 = (string)phrase.Attribute("start2");
                    end2 = (string)phrase.Attribute("end2");
                    Console.WriteLine("Start: " + start1);
                    Console.WriteLine("End: " + end1);
                    // extract pages to new PDF
                    bool phrasematch = ExtractDocument(start1, end1, start2, end2, ocrfile, String.Format("{0}{1}{2}{3}{4}{5}", ExtractFilePath, @"\extracted\",
                        System.IO.Path.GetFileNameWithoutExtension(ocrfile), "_", fileno.ToString(), ".pdf"));
                    if (phrasematch)
                    {
                        match = true;
                    }
                }
                if (match is false)
                {
                    File.Copy(ocrfile, String.Format("{0}{1}", NoMatchFilePath, System.IO.Path.GetFileName(ocrfile)), true);
                }
            }
        }
        catch (Exception ex)
        {
            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log(ex.Message, w);
            }
        }

        using (StreamWriter w = File.AppendText(logfilename))
        {
            Log("End ProcessExtractFiles", w);
        }
    }
    public static void ProcessOCRFiles()
    {
        using (StreamWriter w = File.AppendText(logfilename))
        {
            Log("Start ProcessOCRFiles", w);
        }
        try
        {
            // set counter for testing
            int filecounter = 0;
            // download from Azure file share
            // Connect to the share
            string connectionString = ConnectionString;
            string shareName = ShareName;
            ShareClient share = new ShareClient(connectionString, shareName);

            // Track the remaining directories to walk, starting from the root
            var remaining = new Queue<ShareDirectoryClient>();
            remaining.Enqueue(share.GetRootDirectoryClient());
            var directoryName = "";
            while (remaining.Count > 0)
            {
                // Get all of the next directory's files and subdirectories
                ShareDirectoryClient dir = remaining.Dequeue();

                foreach (ShareFileItem item in dir.GetFilesAndDirectories())
                {
                    // Print the name of the item
                    Console.WriteLine(item.Name);

                    // Keep walking down directories
                    if (item.IsDirectory)
                    {
                        directoryName = item.Name;
                        remaining.Enqueue(dir.GetSubdirectoryClient(item.Name));
                    }
                    else
                    {
                        // download file for processing
                        Console.WriteLine("Directory: " + dir.Name);

                        if (dir.Name == "Initial Scan")
                        {
                            filecounter++;
                            if (filecounter < FileCount)
                            {
                                var sourceFolder = dir.Path; // Path to the save the downloaded file
                                string localFilePath = FileDownLoadPath + item.Name;
                                try
                                {
                                    //// Get a reference to the file
                                    ShareFileClient file = dir.GetFileClient(item.Name);

                                    // Download the file
                                    ShareFileDownloadInfo download = file.Download();
                                    using (FileStream stream = File.OpenWrite(localFilePath))
                                    {
                                        download.Content.CopyTo(stream);
                                    }

                                    // clear list
                                    LIST_IMAGES_OCR.Clear();
                                    ClearFiles();

                                    // pdf extract as image b4 OCR
                                    if (System.IO.Path.GetExtension(localFilePath).ToLower() == ".pdf")
                                    {
                                        // export PDF as image for OCR
                                        PdfDocument pdfImageDoc = new PdfDocument(new PdfReader(localFilePath));

                                        for (int page = 1; page <= pdfImageDoc.GetNumberOfPages(); page++)
                                        {
                                            ExtractImage(localFilePath, String.Format("{0}{1}", ExtractFilePath, @"\images"), page);
                                        }

                                        // OCR images and output PDF
                                        OCRImages(TesseractFilePath, String.Format("{0}{1}{2}", ExtractFilePath, @"\ocr\", System.IO.Path.GetFileName(localFilePath)));
                                    }
                                    else if (System.IO.Path.GetExtension(localFilePath).ToLower() == ".tif")
                                    {
                                        LIST_IMAGES_OCR.Add(new FileInfo(localFilePath));
                                        OCRImages(TesseractFilePath, String.Format("{0}{1}{2}{3}", ExtractFilePath, @"\ocr\", System.IO.Path.GetFileNameWithoutExtension(localFilePath), ".pdf"));
                                    }
                                    else
                                    {
                                        System.IO.File.Copy(localFilePath, String.Format("{0}{1}{2}", ExtractFilePath, @"\skipped\", System.IO.Path.GetFileName(localFilePath)));
                                    }

                                    Console.WriteLine("OCR: " + localFilePath);

                                }
                                catch (Exception err)
                                {
                                    Console.WriteLine("Error: " + localFilePath + " _ " + err.Message);
                                }
                            }
                        }
                    }
                }
            }

        } catch (Exception ex)
        {
            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log(ex.Message, w);
            }
        }
        using (StreamWriter w = File.AppendText(logfilename))
        {
            Log("End ProcessOCRFiles", w);
        }
    }
    private static void ClearFiles()
    {

        try
        {
            string[] imageList = Directory.GetFiles(String.Format("{0}{1}", ExtractFilePath, @"\images"));


            // Copy picture files.
            foreach (string f in imageList)
            {
                // Remove path from the file name.
                File.Delete(f);
            }

        }
        catch { }
        
    } 
    public static bool ExtractDocument(string startText, string endText, string startText2, string endText2, string filePathIn, string filePathOut)
    {
        bool returnValue = false;
        try { 
            // split pdf
            PdfDocument pdfDoc = new PdfDocument(new PdfReader(filePathIn));
            int startpageextract = 0;
            int endpageextract = 0;
            List<string> pagelist = new List<string>();
            int extractcount = 0;
            
            int n = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= n; i++)
            {
                PdfPage pdfPage = pdfDoc.GetPage(i);

                string pdftext = PdfTextExtractor.GetTextFromPage(pdfPage);
                
                if (pdftext.Contains(startText) && pdftext.Contains(startText2)) // text extraction removes spaces!
                {
                    startpageextract = i;
                    endpageextract = 0;

                }
                //
                if (startpageextract > 0)
                {
                    pagelist.Add(i.ToString()); // add page to extract
                }
                if (pdftext.Contains(endText) && pdftext.Contains(endText2))  // text extraction removes spaces!
                {
                    endpageextract = i;
                }
                if (i == n && startpageextract > 0 && endpageextract == 0) // if no end text block found extract all since start text block 
                {
                    endpageextract = i;
                }
                if (startpageextract > 0 && endpageextract > 0)
                {
                    //extract pages
                    extractcount++;

                    string extractFilePath = filePathOut.Replace(System.IO.Path.GetFileNameWithoutExtension(filePathOut), String.Format("{0}{1}{2}", System.IO.Path.GetFileNameWithoutExtension(filePathOut), "_", extractcount.ToString()));
                    string pagerange = String.Join(", ", pagelist);
                    var splitpdf = new MyCustomPdfSplitter(pdfDoc, pageRange => new PdfWriter(extractFilePath));
                    var result = splitpdf.ExtractPageRange(new PageRange(pagerange));
                    result.Close();
                    Console.WriteLine("File:" + filePathIn);
                    Console.WriteLine("Page text:" + i.ToString() + " - " + pdftext);
                    using (StreamWriter w = File.AppendText(logfilename))
                    {
                        Log("File in: " + filePathIn + ", Extract: " + extractFilePath + ", range: " + pagerange, w);
                    }

                    returnValue = true;
                    // clear extraction variables
                    startpageextract = 0;
                    endpageextract = 0;
                    pagelist.Clear();
                }
            }
        }
        catch (Exception ex)
        {
            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log(ex.Message, w);
            }
        }
       
        return returnValue;
    }
    public static void ExtractImage(string inputPDFFile, string outputFilePath, int pageNumber)
    {
        try
        {
            string outImageName = System.IO.Path.GetFileNameWithoutExtension(inputPDFFile);
            outImageName = String.Format("{0}{1}{2}{3}", outImageName, "_image_", pageNumber.ToString(), ".png");


            GhostscriptPngDevice ghost = new GhostscriptPngDevice(GhostscriptPngDeviceType.Png256);
            ghost.GraphicsAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
            ghost.TextAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
            ghost.ResolutionXY = new GhostscriptImageDeviceResolution(290, 290);
            ghost.InputFiles.Add(inputPDFFile);
            ghost.Pdf.FirstPage = pageNumber;
            ghost.Pdf.LastPage = pageNumber;
            ghost.CustomSwitches.Add("-dDOINTERPOLATE");
            ghost.OutputPath = String.Format("{0}{1}{2}", outputFilePath, @"\", outImageName);
            ghost.Process();

            // add to list ready for OCR
            LIST_IMAGES_OCR.Add(new FileInfo(ghost.OutputPath));
        }  
        catch (Exception ex)
        {
            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log(ex.Message, w);
}
        }
       
    }
    // OCR and concatenate images to create PDF 
    private static void OCRImages(string tesseractPath, string fileOutputPath)
    {
        try
        {
            var tesseractReader = new Tesseract4LibOcrEngine(tesseract4OcrEngineProperties);
            tesseract4OcrEngineProperties.SetPathToTessData(new FileInfo(fileName: tesseractPath)); // files downloaded locally
            tesseract4OcrEngineProperties.SetLanguages(new ReadOnlyCollection<string>(new List<string> { "eng" }));
            tesseract4OcrEngineProperties.SetTextPositioning(TextPositioning.BY_WORDS_AND_LINES);

            var properties = new OcrPdfCreatorProperties();
            properties.SetPdfLang("en");

            var ocrPdfCreator = new OcrPdfCreator(tesseractReader, properties);
            using (var writer = new PdfWriter(filename: fileOutputPath, properties: new WriterProperties().AddXmpMetadata()))
            {

                ocrPdfCreator.CreatePdf(inputImages: new ReadOnlyCollection<FileInfo>(list: LIST_IMAGES_OCR), writer).Close();
            }
        }
        catch (Exception ex)
        {
            using (StreamWriter w = File.AppendText(logfilename))
            {
                Log(ex.Message, w);
            }
        }

    }


    public static void Log(string logMessage, TextWriter w)
{
    w.Write("\r\nLog Entry : ");
    w.WriteLine($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
    w.WriteLine("  :");
    w.WriteLine($"  :{logMessage}");
    w.WriteLine("-------------------------------");
}
}
// Thanks to https://stackoverflow.com/users/1729265/mkl for this!
class MyCustomPdfSplitter : PdfSplitter
{
    private Func<PageRange, PdfWriter> customWriter;
    public MyCustomPdfSplitter(PdfDocument pdfDocument, Func<PageRange, PdfWriter> customWriter) : base(pdfDocument)
    {
        this.customWriter = customWriter;
    }

    protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
    {
        return customWriter.Invoke(documentPageRange);
    }
}

using System.IO;
using System.IO.Compression;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ZipToPdfConverter
{
    public class FileConverter
    {
        private Word.Application WordApplication { get; set; }
        private PowerPoint.Application PowerPointApplication { get; set; }

        public FileConverter()
        {
            WordApplication = new Word.Application { Visible = false };
            PowerPointApplication = new PowerPoint.Application { Visible = MsoTriState.msoFalse, };
        }
        
        public void ConvertZipToPdf(string docPath)
        {
            string outputDir = Path.GetFileNameWithoutExtension(docPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);
            
            var tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);

            using (var archive = ZipArchive)
            {
                archive.ExtractToDirectory(tempDirectory);
            }


            Directory.Delete(tempDirectory, true);
            
        }
        
        private void ConvertWordToPdf(string docPath, string outPath)
        {
            WordApplication.Documents.Open(docPath, ReadOnly: false);
            
            WordApplication.ActiveDocument.ExportAsFixedFormat(outPath, Word.WdExportFormat.wdExportFormatPDF);
            
            WordApplication.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, false);
        }

        private void ConvertPowerPointToPdf(string docPath, string outPath)
        {
            PowerPointApplication.Presentations.Open(docPath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            
            PowerPointApplication.ActivePresentation.ExportAsFixedFormat(outPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

            PowerPointApplication.ActivePresentation.Close();
        }
    }
}
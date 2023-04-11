using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ZipToPdfConverter
{
    public class FileConverter
    {
        private Word.Application WordApplication { get; }
        private PowerPoint.Application PowerPointApplication { get; }
        
        private readonly List<string> _wordFileTypes;
        private readonly List<string> _powerPointFileTypes;

        public FileConverter()
        {
            WordApplication = new Word.Application { Visible = false };
            PowerPointApplication = new PowerPoint.Application();
            _wordFileTypes = new List<string> { ".docx", "doc" };
            _powerPointFileTypes = new List<string> { ".pptx", ".ppt" };
        }

        public void ConvertZipToPdf(string docPath)
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDirectory);

            ZipFile.ExtractToDirectory(docPath, tempDirectory);

            string destDir = Path.Combine(tempDirectory, Path.GetFileName(docPath.RemoveRightToChar("-")));

            Directory.CreateDirectory(destDir);
            

            foreach (string filePath in Directory.GetFiles(tempDirectory))
            {
                if (!_wordFileTypes.Concat(_powerPointFileTypes).Contains(Path.GetExtension(filePath)))
                {
                    File.Move(filePath, Path.Combine(destDir, Path.GetFileName(filePath)));
                    continue;
                }

                string newFilePath = filePath.Replace(Path.GetExtension(filePath), ".pdf");
                if (_wordFileTypes.Contains(Path.GetExtension(filePath)))
                    ConvertWordToPdf(filePath, newFilePath);
                else
                    ConvertPowerPointToPdf(filePath, newFilePath);

                File.Delete(filePath);
                File.Move(newFilePath, Path.Combine(destDir, Path.GetFileName(newFilePath)));
            }

            string outputDir = docPath.RemoveRightToChar("-") + ".zip";

            if (File.Exists(outputDir))
                File.Delete(outputDir);
            
            ZipFile.CreateFromDirectory(tempDirectory, outputDir);

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
            PowerPoint.Presentation presentation = PowerPointApplication.Presentations.Open(docPath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            
            presentation.ExportAsFixedFormat(outPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

            presentation.Close();
        }
    }
}
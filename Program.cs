using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;

namespace convert_docx_to_pdf_use_Microsoft.Office.Interop.Word
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /*
 * Convert Input.docx into Output.pdf
 * Please note: You must have the Microsoft Office 2007 Add-in: Microsoft Save as PDF or XPS installed
 * http://www.microsoft.com/downloads/details.aspx?FamilyId=4D951911-3E7E-4AE6-B059-A2E79ED87041&displaylang=en
 * Solution source http://cathalscorner.blogspot.com/2009/10/converting-docx-into-doc-pdf-html.html
 */
            Convert(@"D:\Tai_Lieu\convert docx to pdf use Microsoft.Office.Interop.Word\Data\GCN.docx",
                @"D:\Tai_Lieu\convert docx to pdf use Microsoft.Office.Interop.Word\Data\GCN.pdf", WdSaveFormat.wdFormatPDF);

            Console.WriteLine("Document... Converted!");
            Console.ReadKey();
        }
        public static void Convert(string input, string output, WdSaveFormat format)
        {
            // Create an instance of Word.exe
            /*           _Application oWord = new Word.Application
                       {

                           // Make this instance of word invisible (Can still see it in the taskmgr).
                           Visible = false
                       };*/
            Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = false;
            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;     // Does not cause any word dialog to show up
            //object readOnly = false;  // Causes a word object dialog to show at the end of the conversion
            object oInput = input;
            object oOutput = output;
            object oFormat = format;

            // Load a document into our instance of word.exe
            _Document oDoc = oWord.Documents.Open(
                ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Make this document the active document.
            oDoc.Activate();

            // Save this document using Word
            oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Always close Word.exe.
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }
    }
}

#This code in this script was adapted from the following sources: 
# https://community.spiceworks.com/topic/2302788-using-openxml-with-powershell
# https://stackoverflow.com/a/11526470/16864869

$SourceCode = @"
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public class CopyVba
{
    public static void FromToExcel(string inputFile, string outputFile)
    {
        using (SpreadsheetDocument ssDoc = SpreadsheetDocument.Open(inputFile, false))
        {
            //WorkbookPart wbPart = ssDoc.WorkbookPart;
            MemoryStream ms = new MemoryStream();
            CopyStream(ssDoc.WorkbookPart.VbaProjectPart.GetStream(), ms);

            using (SpreadsheetDocument ssDoc2 = SpreadsheetDocument.Open(outputFile, true))
            {
                Stream stream = ssDoc2.WorkbookPart.VbaProjectPart.GetStream();
                ms.WriteTo(stream);
            }
        }
    }

    public static void FromToWord(string inputFile, string outputFile)
    {
        using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(inputFile, false))
        {
            //MainDocumentPart wdPart = wdDoc.MainDocumentPart;
            MemoryStream ms = new MemoryStream();
            CopyStream(wdDoc.MainDocumentPart.VbaProjectPart.GetStream(), ms);

            using (WordprocessingDocument wdDoc2 = WordprocessingDocument.Open(outputFile, true))
            {
                Stream stream = wdDoc2.MainDocumentPart.VbaProjectPart.GetStream();
                ms.WriteTo(stream);
            }
        }
    }

    public static void FromToPowerPoint(string inputFile, string outputFile)
    {
        using (PresentationDocument prDoc = PresentationDocument.Open(inputFile, false))
        {
            //MainDocumentPart wdPart = wdDoc.MainDocumentPart;
            MemoryStream ms = new MemoryStream();
            CopyStream(prDoc.PresentationPart.VbaProjectPart.GetStream(), ms);

            using (PresentationDocument prDoc2 = PresentationDocument.Open(outputFile, true))
            {
                Stream stream = prDoc2.PresentationPart.VbaProjectPart.GetStream();
                ms.WriteTo(stream);
            }
        }
    }

    public static void CopyStream(Stream input, Stream output)
    {
        byte[] buffer = new byte[short.MaxValue + 1];
        while (true)
        {
            int read = input.Read(buffer, 0, buffer.Length);
            if (read <= 0)
                return;
            output.Write(buffer, 0, read);
        }
    }
}
"@

$Assemblies = (
  "windowsbase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35",
  "DocumentFormat.OpenXml, Version=2.5.5631.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35",
  "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "System.IO.Packaging, Version=0.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
)

Add-Type -ReferencedAssemblies $Assemblies -TypeDefinition $SourceCode -Language CSharp

# //Examples on how to use this script

# //Excel

# $SourcePath = "C:\temp\abc123.xlsm"

# $DestPath = "C:\temp\bcd234.xlsm"

# [CopyVba]::FromToExcel($SourcePath, $DestPath)

# //Word

# $SourcePath = "C:\temp\abc123.docm"

# $DestPath = "C:\temp\bcd234.docm"

# [CopyVba]::FromToWord($SourcePath, $DestPath)

# //PowerPoint

# $SourcePath = "C:\temp\abc123.pptm"

# $DestPath = "C:\temp\bcd234.pptm"

# [CopyVba]::FromToPowerPoint($SourcePath, $DestPath)
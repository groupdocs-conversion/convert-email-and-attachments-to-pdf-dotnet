// Topic: Convert email messages with attachments to PDF - the message alone, or the message
// and every attachment as separate files that never overwrite each other.
// Uses GroupDocs.Conversion for .NET 26.10: EmailLoadOptions and SaveContext.

using GroupDocs.Conversion;
using GroupDocs.Conversion.FileTypes;
using GroupDocs.Conversion.Options.Convert;
using GroupDocs.Conversion.Options.Load;

namespace Demo.EmailToPdf;

internal static class Program
{
    private const string InputFolder = "resources";
    private const string OutputFolder = "output";

    private static readonly string OutlookMessage =
        Path.Combine(InputFolder, "with-attachments.msg");
    private static readonly string SupportCase =
        Path.Combine(InputFolder, "support-case.eml");

    private static int Main()
    {
        ApplyLicense();

        foreach (string input in new[] { OutlookMessage, SupportCase })
        {
            if (!File.Exists(input))
            {
                Console.Error.WriteLine($"Missing input: {Path.GetFullPath(input)}");
                return 1;
            }
        }

        Directory.CreateDirectory(OutputFolder);

        Console.WriteLine("1. Outlook message as one PDF:");
        ConvertMessageToPdf(OutlookMessage, Path.Combine(OutputFolder, "message-only.pdf"));

        string itemsFolder = Path.Combine(OutputFolder, "message-and-attachments");
        Directory.CreateDirectory(itemsFolder);
        Console.WriteLine("2. Outlook message and each attachment as PDF:");
        int items = ConvertMessageAndAttachments(OutlookMessage, itemsFolder);
        Console.WriteLine($"   files: {items}");

        string caseFolder = Path.Combine(OutputFolder, "support-case");
        Directory.CreateDirectory(caseFolder);
        Console.WriteLine("3. Attachments that share a name or have no PDF route:");
        int kept = KeepAttachmentsWithoutPdfRoute(SupportCase, caseFolder);
        Console.WriteLine($"   kept as they are: {kept}");

        Console.WriteLine($"Output folder: {Path.GetFullPath(OutputFolder)}");
        return items == 3 && kept == 1 ? 0 : 2;
    }

    private static void ApplyLicense()
    {
        // Point this at your .lic file to remove evaluation limits.
        // Get a free temporary licence: https://purchase.groupdocs.com/temporary-license
        const string licensePath = "REPLACE_WITH_YOUR_LICENSE_PATH";
        if (File.Exists(licensePath))
        {
            new License().SetLicense(licensePath);
            Console.WriteLine("[license] applied");
        }
        else
        {
            Console.WriteLine("[license] no licence set - running in evaluation mode");
        }
    }

    /// <summary>
    /// Converts an email message to a single PDF that lists its attachments by name.
    /// </summary>
    /// <remarks>
    /// With ConvertOwned switched off only the message itself is converted: the header
    /// fields, the body and the list of attached files. Use it when the PDF is the record
    /// of the message and the attachments are archived separately.
    /// </remarks>
    private static void ConvertMessageToPdf(string messagePath, string outputPath)
    {
        var loadOptions = new EmailLoadOptions
        {
            ConvertOwner = true,
            ConvertOwned = false
        };

        using (var converter = new Converter(messagePath, (LoadContext _) => loadOptions))
        {
            converter.Convert(outputPath, new PdfConvertOptions());
        }

        Console.WriteLine($"   written: {Path.GetFileName(outputPath)}");
    }

    /// <summary>
    /// Converts an email message and each of its attachments to separate PDF files.
    /// </summary>
    /// <remarks>
    /// ConvertOwner and ConvertOwned convert the message and its attachments; Depth limits
    /// how far attachments inside attached messages are followed. SaveContext describes
    /// each output: ItemIndex 1 is the message, the attachments follow, and the index keeps
    /// every file name unique. Returns the number of files written.
    /// </remarks>
    private static int ConvertMessageAndAttachments(string messagePath, string outputFolder)
    {
        var loadOptions = new EmailLoadOptions
        {
            ConvertOwner = true,
            ConvertOwned = true,
            Depth = 2
        };
        int written = 0;

        using (var converter = new Converter(messagePath, (LoadContext _) => loadOptions))
        {
            converter.Convert((SaveContext item) =>
            {
                written++;
                string name = $"{item.ItemIndex:D2}-{item.SourceFileName}.pdf";
                Console.WriteLine($"   level {item.HierarchyLevel}: {name}");
                return File.Create(Path.Combine(outputFolder, name));
            }, new PdfConvertOptions());
        }

        return written;
    }

    /// <summary>
    /// Converts a message whose attachments share a file name or cannot become a PDF.
    /// </summary>
    /// <remarks>
    /// Two attachments named report.csv get different item indexes, so both PDFs are kept.
    /// An attachment with no route to PDF no longer stops the conversion: it is handed back
    /// unchanged with an unknown target format, so it is saved under its own name. Returns
    /// the number of attachments kept as they are.
    /// </remarks>
    private static int KeepAttachmentsWithoutPdfRoute(
        string messagePath, string outputFolder)
    {
        var loadOptions = new EmailLoadOptions { ConvertOwner = true, ConvertOwned = true };
        int keptAsTheyAre = 0;

        using (var converter = new Converter(messagePath, (LoadContext _) => loadOptions))
        {
            converter.Convert((SaveContext item) =>
            {
                bool passedThrough = item.TargetFormat == FileType.Unknown;
                string baseName = $"{item.ItemIndex:D2}-{item.SourceFileName}";
                string name = passedThrough
                    ? baseName
                    : $"{baseName}.{item.TargetFormat.Extension}";
                keptAsTheyAre += passedThrough ? 1 : 0;
                string action = passedThrough ? "kept     " : "converted";
                Console.WriteLine($"   {action}: {name}");
                return File.Create(Path.Combine(outputFolder, name));
            }, new PdfConvertOptions());
        }

        return keptAsTheyAre;
    }
}

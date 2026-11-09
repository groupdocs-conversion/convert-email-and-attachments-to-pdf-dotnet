# Convert Email Messages and Attachments to PDF in .NET

[![Product Page](https://img.shields.io/badge/Product%20Page-2865E0?style=for-the-badge&logo=appveyor&logoColor=white)](https://github.com/groupdocs-conversion/GroupDocs.Conversion-Docs)
[![Docs](https://img.shields.io/badge/Docs-2865E0?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.groupdocs.com/conversion/net/)
[![Blog](https://img.shields.io/badge/Blog-2865E0?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.groupdocs.com/categories/groupdocs.conversion-product-family/)
[![Free Support](https://img.shields.io/badge/Free%20Support-2865E0?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.groupdocs.com/c/conversion/)
[![Temporary License](https://img.shields.io/badge/Temporary%20License-2865E0?style=for-the-badge&logo=rocket&logoColor=white)](https://purchase.groupdocs.com/temp-license/100032)

## Overview

`convert-email-and-attachments-to-pdf-dotnet` is a .NET console project that
shows three ways to turn an Outlook message and its attachments into PDF
output with GroupDocs.Conversion for .NET 26.10. One mode produces a single
PDF for the message, one produces a PDF per attachment, and one keeps
attachments that have no route to PDF instead of failing the run. Each
mode is a few lines of code built on `EmailLoadOptions` and the `SaveContext`
callback. No manual export step is involved anywhere.

## Technology Stack

- **Platform**: .NET 8.0 (`net8.0`)
- **Product**: GroupDocs.Conversion for .NET 26.10.0
- **Language**: C#

## When should I convert each attachment separately instead of just the message?

Convert the message alone when the PDF only has to stand in for the email -
header and body, attachment names listed for reference. Convert the message
and every attachment separately when each attached file needs its own
review, filing, or index entry. The second mode costs one callback per item
instead of one call, a small trade for an addressable PDF per attachment.

## Use Case Scenarios

This repository addresses the following use cases:

- **Support case archiving**: Convert a support email together with its
  spreadsheet attachments into separate PDFs so each report can be filed and
  searched on its own.
- **Message-only record**: Produce a single PDF that captures the sender,
  recipient, subject and body of an Outlook message when the attachments are
  archived through a different process.
- **Attachment-by-attachment review**: Produce one PDF per attachment,
  including the message itself as item one, so a reviewer can open any single
  piece without the rest.
- **Mixed-format inboxes**: Keep an attachment with no conversion route to
  PDF, such as a proprietary capture file, alongside the converted PDFs
  instead of losing it or failing the whole run.

## Problem Statement

Turning an email into something a document workflow can use is more than
running one file through a converter. An Outlook `.msg` file can carry
several attachments of different formats. The header, the body and the
attachment list all need a predictable home. Names alone are not enough.
Two attachments can carry the exact same file name. The sample support case
in this repository attaches
two files both named `report.csv`, one with October figures and one with
September figures. Naming output files by attachment name alone would
silently overwrite one with the other.

Some attachments also have no route to PDF at all. Before GroupDocs.Conversion
for .NET 26.10, an attachment like that stopped the conversion with a
`ConversionNotSupportedException`. Nothing after it was converted.
`EmailLoadOptions` and the
`SaveContext` callback fix both problems at once: they describe every output
item precisely enough that names never collide and an unsupported attachment
never blocks the rest.

### Common Challenges

- ❌ Naming attachment outputs by file name alone, which breaks when two
  attachments share a name.
- ❌ Losing the whole conversion because one attachment has no route to the
  target format.
- ❌ Choosing between a single summary PDF and per-attachment PDFs without a
  documented way to switch between them.
- ❌ Reconstructing which attachment is which once everything has been
  flattened into generic output files.

## Solution Overview

GroupDocs.Conversion for .NET separates two concerns. One is what gets
converted. The other is how each output is named. `EmailLoadOptions.ConvertOwner`
and `ConvertOwned` decide
whether the message, the attachments, or both get processed; `Depth`
(default `2`) decides how many levels of nested attachments are followed.
`SaveContext` then describes every output item as the conversion runs:
`ItemIndex` counts up from 1 and stays unique even for identically named
attachments, `HierarchyLevel` tells a callback whether the current item is
the message (`0`) or an attachment (`1`), `SourceFileName` carries the
original name, and `TargetFormat` reports `FileType.Unknown` when an
attachment has no route to PDF so it can be saved under its own name
instead of blocking the run.

## Prerequisites

To use these examples, you'll need:

- **.NET SDK**: 8.0 or later (the sample project targets `net8.0`).
- **GroupDocs.Conversion for .NET**: version 26.10.0 or later via NuGet
  ([get a temporary license](https://purchase.groupdocs.com/temp-license/100032)
  to remove evaluation limits).
- **Sample input files**: an Outlook `.msg` message and an `.eml` message,
  both included under `resources/`.

## Getting Started

### Installation

```bash
dotnet add package GroupDocs.Conversion --version 26.10.0
```

### Configuration

1. Point `ApplyLicense`'s `licensePath` constant at a real `.lic` file, or
   leave the placeholder in place to run in evaluation mode.
2. Build the project with `dotnet build`.
3. Run it from the project folder so the `resources/` paths resolve.

## Repository Structure

```
convert-email-and-attachments-to-pdf-dotnet/
│
├── EmailToPdfDemo.csproj
├── Program.cs
└── resources/
    ├── support-case.eml
    └── with-attachments.msg
```

### File Descriptions

- `EmailToPdfDemo.csproj` - project file targeting `net8.0`, referencing
  `GroupDocs.Conversion` 26.10.0.
- `Program.cs` - the three conversion methods shown below, plus the console
  entry point that runs all three against the sample messages.
- `resources/support-case.eml` - a sample support e-mail with two
  attachments both named `report.csv` and one attachment of an unsupported
  type, used to demonstrate duplicate names and the no-route case.
- `resources/with-attachments.msg` - a sample Outlook message with an image
  and a Word document attached, used to demonstrate message-plus-attachments
  conversion.

## Code Implementation

### Implementation: Converting the Message Alone

With `ConvertOwned` switched off, only the message itself is converted into
one PDF: the header fields, the body, and the list of attached file names
for reference. The attachments appear by name only.

```csharp
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
```

#### Technical Details

`ConvertOwner = true` tells the loader to treat the message as the thing
being converted; `ConvertOwned = false` tells it to skip every attachment.
A single call to `Converter.Convert` with a plain output path is all that's
needed, exactly as it would be for any other single document. Running this
against the sample Outlook message produces one page whose header lists
`From`, `Sent`, `To`, `Subject` and `Attachments` (the attached file names),
followed by the body.

This is the smallest of the three methods. It also has one firm limit:
there is no way back to the attachment content from this PDF alone.

Use this when a case-management or audit system only needs the message as
a record, and the attachments already go through their own archiving path.

### Implementation: Converting the Message and Every Attachment

`ConvertOwner` and `ConvertOwned` convert the message and its attachments;
`Depth` limits how far attachments inside attached messages are followed.
`SaveContext` describes each output as the conversion produces it, instead
of handing back one fixed file.

```csharp
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
```

#### Technical Details

The callback runs once per output item. `ItemIndex` is 1 for the message
and counts up for each attachment, so the file name built from it never
repeats. `HierarchyLevel` tells the callback what kind of item it is
holding: 0 for the message, 1 for its attachments. Two attachments in,
three files out, because the message itself counts as item one.

Point this at the sample Outlook message and the output folder ends up
with:

```
01-with-attachments.msg.pdf   level 0 (the message)
02-attachment-image.png.pdf   level 1 (the image attachment)
03-attachment-word.doc.pdf    level 1 (the Word attachment)
```

Use this when a document management system needs the message and every
attachment indexed as separate, independently searchable PDF records.

### Implementation: Keeping Attachments That Have No PDF Route

Two attachments named `report.csv` get different `ItemIndex` values, so
both PDFs are kept. An attachment with no route to PDF no longer stops the
conversion: it is handed back unchanged with `TargetFormat` equal to
`FileType.Unknown`, so it is saved under its own name instead.

```csharp
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
```

#### Technical Details

Checking `item.TargetFormat == FileType.Unknown` before building the output
name is the whole trick. That one branch is what lets a run finish with
every input accounted for, converted or not.

| Output | Outcome |
|---|---|
| `01-support-case.eml.pdf` | converted (the message) |
| `02-report.csv.pdf` | converted (October figures) |
| `03-report.csv.pdf` | converted (September figures) |
| `04-gateway-capture.xyz` | kept, unchanged |

The kept file matches the attachment exactly - all 32 bytes accounted for,
nothing re-encoded or re-wrapped.

Use this for inboxes where attachment formats vary and a single unsupported
file should never take down the rest of the conversion.

## Best Practices

When implementing email-to-PDF conversion with attachments, consider these
practices:

- **Dispose the converter**: wrap it in a `using` block so file handles
  close even if a conversion throws partway through.
- **Name by `ItemIndex`, not `SourceFileName`**: the index is the only field
  guaranteed unique per output item.
- **Check `TargetFormat` before assuming PDF**: an item passed straight
  through still needs its own extension logic, not `.pdf`.
- **Keep `Depth` deliberate**: raise it only when an attached message can
  itself carry attachments worth converting.
- **Test with colliding names on purpose**: a sample with two identically
  named attachments catches naming bugs before production does.

## Additional Resources

For more in-depth information about converting email with attachments to
PDF, explore these resources:

* **Step-by-step use case guide in the documentation** - walks through the
  same message-plus-attachments scenario in the official docs: [Read the article →](https://docs.groupdocs.com/conversion/net/use-cases/convert-email-with-attachments-to-pdf/)

* **In-depth blog article about this project** - the full write-up behind
  this repository: [Read the article →](https://blog.groupdocs.com/conversion/convert-email-with-attachments-to-pdf-net/)

* **Converting each email attachment to a different output format** - the
  reference page for routing attachments to different target formats by
  type: [Read the article →](https://docs.groupdocs.com/conversion/net/convert-each-email-attachment-to-different-format/)

* **Load options for e-mail documents** - the full `EmailLoadOptions`
  reference, including every display and field-mapping option: [Read the article →](https://docs.groupdocs.com/conversion/net/load-email-document-with-options/)

* **Convert EML or MSG files to PDF in C#** - an earlier walkthrough of basic
  message-to-PDF conversion on the GroupDocs blog: [Read the article →](https://blog.groupdocs.com/conversion/convert-eml-or-msg-file-to-pdf-in-csharp/)

* **GroupDocs.Conversion for .NET 26.10 release highlights** - the release
  that stopped an unsupported attachment from failing the whole conversion:
  [Read the article →](https://blog.groupdocs.com/conversion/groupdocs-conversion-for-net-26-10/)

* **GroupDocs.Conversion for .NET documentation and API reference** - the
  full developer guide and class reference: [Read the article →](https://docs.groupdocs.com/conversion/net/)

## Keywords

`GroupDocs.Conversion`, `.NET`, `C#`, `email to PDF`, `MSG to PDF`, `EML to PDF`,
`email attachments`, `convert attachments to PDF`, `EmailLoadOptions`,
`SaveContext`, `email archiving`, `Outlook message conversion`,
`attachment conversion`, `document conversion`, `PDF generation`,
`.NET console application`, `ItemIndex`, `HierarchyLevel`, `TargetFormat`

## Support

For technical support, visit:
- [Free Support Forum](https://forum.groupdocs.com/c/conversion/)
- [Product Documentation](https://docs.groupdocs.com/conversion/net/)
- [Get Temporary License](https://purchase.groupdocs.com/temp-license/100032)

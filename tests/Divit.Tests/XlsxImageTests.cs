namespace Divit.Tests;

using Divit.Packaging;
using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

/// <summary>
/// Images: PNG, JPEG and GIF, their size read from the file, placed at a cell or filling a range.
/// </summary>
public class XlsxImageTests
{
    private const long Emu = 9525;

    private static byte[] OnePixelPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    /// <summary>
    /// The header of a 300 × 200 GIF.
    /// </summary>
    private static byte[] Gif() => new byte[] { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 0x2C, 0x01, 0xC8, 0x00, 0, 0, 0 };

    /// <summary>
    /// The start of a 300 × 200 JPEG: an APP0 block, a Huffman table (DHT, which shares the SOF number range), fill
    /// bytes, then the frame header.
    /// </summary>
    private static byte[] Jpeg(byte frame = 0xC0)
    {
        var app0 = new byte[] { 0xFF, 0xE0, 0x00, 0x10, (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0, 1, 1, 0, 0, 1, 0, 1, 0, 0 };
        var dht = new byte[] { 0xFF, 0xC4, 0x00, 0x04, 0, 0 };
        var sof = new byte[] { 0xFF, 0xFF, 0xFF, frame, 0x00, 0x11, 0x08, 0x00, 0xC8, 0x01, 0x2C, 0x03, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1 };
        return new byte[] { 0xFF, 0xD8 }.Concat(app0).Concat(dht).Concat(sof).ToArray();
    }

    [Fact]
    public void SizesAreReadFromTheFile()
    {
        var png = ImageInfo.Read(OnePixelPng())!.Value;
        Assert.Equal(("png", "image/png", 1, 1), (png.Extension, png.ContentType, png.Width, png.Height));
        var gif = ImageInfo.Read(Gif())!.Value;
        Assert.Equal(("gif", 300, 200), (gif.Extension, gif.Width, gif.Height));
        var jpeg = ImageInfo.Read(Jpeg())!.Value;
        Assert.Equal(("jpeg", "image/jpeg", 300, 200), (jpeg.Extension, jpeg.ContentType, jpeg.Width, jpeg.Height));
        var progressive = ImageInfo.Read(Jpeg(0xC2))!.Value;
        Assert.Equal((300, 200), (progressive.Width, progressive.Height));
        Assert.Null(ImageInfo.Read(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 }));
        Assert.Null(ImageInfo.Read(new byte[] { 0x42, 0x4D, 0, 0 }));
    }

    [Fact]
    public void ImagesOfEveryFormatAreWritten()
    {
        var book = new XlsxWorkbook { Created = XlsxTestKit.FixedTime };
        var sheet = book.AddSheet("Images");
        sheet.AddImage(1, 1, Jpeg(), description: "A photo");
        sheet.AddImage(1, 8, Gif(), width: 150);
        sheet.AddImage(12, 1, OnePixelPng(), 40, 40);
        sheet.AddImageInCells(12, 4, 20, 9, Jpeg());
        var path = Path.Combine(XlsxTestKit.NewFolder(), "images-check.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));

        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            Assert.Equal(new[] { "image/jpeg", "image/gif", "image/png", "image/jpeg" },
                drawing.ImageParts.Select(p => p.ContentType).ToArray());
            Assert.Equal(new[] { "/xl/media/image1.jpeg", "/xl/media/image2.gif", "/xl/media/image3.png", "/xl/media/image4.jpeg" },
                drawing.ImageParts.Select(p => p.Uri.ToString()).ToArray());

            var anchors = drawing.WorksheetDrawing.ChildElements.ToList();
            var natural = (Xdr.OneCellAnchor)anchors[0];
            Assert.Equal(300 * Emu, natural.Extent!.Cx!.Value);
            Assert.Equal(200 * Emu, natural.Extent.Cy!.Value);
            Assert.Equal("A photo", natural.Descendants<Xdr.NonVisualDrawingProperties>().Single().Description!.Value);

            // Only the width given: the height follows the proportions.
            var scaled = (Xdr.OneCellAnchor)anchors[1];
            Assert.Equal(150 * Emu, scaled.Extent!.Cx!.Value);
            Assert.Equal(100 * Emu, scaled.Extent.Cy!.Value);
            Assert.Null(scaled.Descendants<Xdr.NonVisualDrawingProperties>().Single().Description);

            Assert.Equal(40 * Emu, ((Xdr.OneCellAnchor)anchors[2]).Extent!.Cx!.Value);

            // The range image runs from the corner of its first cell to the corner after its last cell.
            var cells = (Xdr.TwoCellAnchor)anchors[3];
            Assert.Equal(Xdr.EditAsValues.TwoCell, cells.EditAs!.Value);
            Assert.Equal(("3", "11"), (cells.FromMarker!.ColumnId!.Text, cells.FromMarker.RowId!.Text));
            Assert.Equal(("9", "20"), (cells.ToMarker!.ColumnId!.Text, cells.ToMarker.RowId!.Text));
        }
    }

    [Fact]
    public void PngOnlyWorkbooksDeclareNoOtherFormats()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("One").AddImage(1, 1, OnePixelPng());
        var path = Path.Combine(XlsxTestKit.NewFolder(), "png-check.xlsx");
        book.Save(path);
        using (var zip = new System.IO.Compression.ZipArchive(File.OpenRead(path), System.IO.Compression.ZipArchiveMode.Read))
        using (var reader = new StreamReader(zip.GetEntry("[Content_Types].xml")!.Open()))
        {
            var types = reader.ReadToEnd();
            Assert.Contains("Extension=\"png\"", types);
            Assert.DoesNotContain("jpeg", types);
            Assert.DoesNotContain("gif", types);
        }
    }

    [Fact]
    public void BadImagesAreRejected()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentException>(() => sheet.AddImage(1, 1, new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }));
        Assert.Throws<ArgumentException>(() => sheet.AddImageInCells(1, 1, 2, 2, new byte[] { 0x42, 0x4D, 0, 0 }));

        // A JPEG whose frame header cannot be found needs its size from the caller.
        var headless = new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 };
        Assert.Throws<ArgumentException>(() => sheet.AddImage(1, 1, headless));
        sheet.AddImage(1, 1, headless, 100, 50);
        Assert.Throws<ArgumentException>(() => sheet.AddImageInCells(3, 3, 1, 1, Jpeg()));
    }

#if NETFRAMEWORK
    /// <summary>
    /// A gallery file with real pictures: a JPEG photo-like gradient at its own size and one filling a block of cells,
    /// a GIF and a PNG; resize the rows and columns in Excel to see which ones follow.
    /// </summary>
    [Fact]
    public void GalleryOfImages()
    {
        byte[] Draw(System.Drawing.Imaging.ImageFormat format, System.Drawing.Color from, System.Drawing.Color to, string label)
        {
            using (var bitmap = new System.Drawing.Bitmap(320, 200))
            using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
            using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(new System.Drawing.Rectangle(0, 0, 320, 200), from, to, 35f))
            using (var font = new System.Drawing.Font("Segoe UI", 20, System.Drawing.FontStyle.Bold))
            using (var stream = new MemoryStream())
            {
                graphics.FillRectangle(brush, 0, 0, 320, 200);
                graphics.FillEllipse(System.Drawing.Brushes.White, 200, 30, 90, 90);
                graphics.DrawString(label, font, System.Drawing.Brushes.White, 16, 140);
                bitmap.Save(stream, format);
                return stream.ToArray();
            }
        }

        var book = new XlsxWorkbook { Title = "Images", Created = XlsxTestKit.FixedTime };
        var sheet = book.AddSheet("Images");
        sheet.Set(1, 1, "JPEG at its own size: moves with B3, keeps its size.");
        sheet.AddImage(3, 2, Draw(System.Drawing.Imaging.ImageFormat.Jpeg, System.Drawing.Color.SteelBlue, System.Drawing.Color.DarkOrange, "JPEG"), description: "A gradient drawn as a JPEG");
        sheet.Set(1, 8, "GIF, 160 pixels wide: the height follows.");
        sheet.AddImage(3, 8, Draw(System.Drawing.Imaging.ImageFormat.Gif, System.Drawing.Color.SeaGreen, System.Drawing.Color.Gold, "GIF"), width: 160);
        sheet.Set(15, 1, "JPEG filling B17:F28: stretches when rows or columns change.");
        sheet.AddImageInCells(17, 2, 28, 6, Draw(System.Drawing.Imaging.ImageFormat.Jpeg, System.Drawing.Color.Purple, System.Drawing.Color.HotPink, "In cells"));
        sheet.Set(15, 8, "PNG");
        sheet.AddImage(17, 8, Draw(System.Drawing.Imaging.ImageFormat.Png, System.Drawing.Color.Black, System.Drawing.Color.Gray, "PNG"), width: 200);
        var path = Path.Combine(XlsxTestKit.NewFolder(), "images.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        XlsxTestKit.Keep(path, "images.xlsx");
    }
#endif
}

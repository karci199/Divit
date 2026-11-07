namespace Divit.Tests;

using Divit.Packaging;
using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Xunit;

/// <summary>
/// Sheet and workbook locks: the SHA-512 hash Excel checks, permissions, editable cells and reproducible salts.
/// </summary>
public class XlsxProtectionTests
{
    private static string Save(XlsxWorkbook book, string name, bool keep = false)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    /// <summary>
    /// A vector from an independent implementation: ClosedXML protected a sheet with this password, and wrote this
    /// salt and hash. Matching it means Excel will accept the password.
    /// </summary>
    [Fact]
    public void HashMatchesAnIndependentImplementation()
    {
        var salt = Convert.FromBase64String("0p07k8nOwkXha/xSswhJEJUpaLEvE02kVJd+3B1A/tE=");
        Assert.Equal("kFNLpf30kEKkUpFGb3277CMHOCG4JQO+HNAHHgmlvXl2a41+L+JVuXGEK4+Z4/GWGCijQGicZ6EV1Ma66z08/g==",
            PasswordHash.Sha512("Divit-ş3cret", salt));
    }

    [Fact]
    public void HashWithoutRoundsIsPlainSha512OfSaltAndPassword()
    {
        var salt = new byte[] { 1, 2, 3 };
        byte[] expected;
        using (var sha = System.Security.Cryptography.SHA512.Create())
        {
            expected = sha.ComputeHash(salt.Concat(System.Text.Encoding.Unicode.GetBytes("ab")).ToArray());
        }

        Assert.Equal(Convert.ToBase64String(expected), PasswordHash.Sha512("ab", salt, 0));
    }

    [Fact]
    public void SaltsAreRandomUnlessSeeded()
    {
        Assert.Equal(16, PasswordHash.Salt(null).Length);
        Assert.NotEqual(PasswordHash.Salt(null), PasswordHash.Salt(null));
        Assert.Equal(PasswordHash.Salt("seed"), PasswordHash.Salt("seed"));
        Assert.NotEqual(PasswordHash.Salt("seed"), PasswordHash.Salt("other"));
    }

    [Fact]
    public void HashingIsFastEnough()
    {
        var watch = Stopwatch.StartNew();
        PasswordHash.Sha512("password", PasswordHash.Salt("speed"));
        Assert.True(watch.ElapsedMilliseconds < 2000, "100 000 rounds took " + watch.ElapsedMilliseconds + " ms");
    }

    /// <summary>
    /// An order form: labels locked, input cells editable, the total's formula hidden, lookup data in a very hidden
    /// sheet behind a structure lock.
    /// </summary>
    private static XlsxWorkbook OrderForm()
    {
        var book = new XlsxWorkbook { Title = "Protection", Created = XlsxTestKit.FixedTime };
        var form = book.AddSheet("Order form");
        var prices = book.AddSheet("Prices");
        prices.Set(1, 1, "Pen");
        prices.Set(1, 2, 1.5);
        prices.Set(2, 1, "Paper");
        prices.Set(2, 2, 4.25);
        prices.Visibility = XlsxSheetVisibility.VeryHidden;
        book.DefineName("Products", prices, 1, 1, 2, 1);
        book.DefineName("PriceList", prices, 1, 1, 2, 2);

        form.Set(1, 1, "Fill in the yellow cells; everything else is locked (password: divit).");
        form.Merge(1, 1, 1, 4);
        form.Set(3, 1, "Product");
        form.Set(4, 1, "Quantity");
        form.Set(5, 1, "Total");
        form.Format(3, 1, 5, 1, (ref XlsxStyle s) => s.Bold = true);
        form.Format(3, 2, 4, 2, (ref XlsxStyle s) => { s.Unlocked = true; s.Fill = "#FFF2CC"; });
        form.Validate(3, 2, 3, 2, XlsxValidation.ListFrom("Products"));
        form.Validate(4, 2, 4, 2, XlsxValidation.WholeNumber(XlsxComparison.Between, 1, 1000));
        form.SetFormula(5, 2, "IFERROR(VLOOKUP(B3,PriceList,2,FALSE)*B4,0)", 0.0);
        form.Format(5, 2, (ref XlsxStyle s) => { s.HideFormula = true; s.NumberFormat = "#,##0.00"; });
        form.ColumnWidth(1, 12);
        form.ColumnWidth(2, 16);
        form.Protect("divit", XlsxAllow.FormatColumns | XlsxAllow.Sort);
        book.Protect("structure");
        return book;
    }

    [Fact]
    public void SheetLockUsesSha512AndKeepsInputCellsEditable()
    {
        var path = Save(OrderForm(), "protection.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var part = document.WorkbookPart!;
            var sheet = ((WorksheetPart)part.GetPartById(part.Workbook.Sheets!.Elements<Sheet>().First().Id!)).Worksheet;
            var protection = sheet.GetFirstChild<SheetProtection>()!;
            Assert.Equal("SHA-512", protection.AlgorithmName!.Value);
            Assert.Equal(100000u, protection.SpinCount!.Value);
            Assert.Null(protection.Password);
            var salt = Convert.FromBase64String(protection.SaltValue!.Value!);
            Assert.Equal(16, salt.Length);
            Assert.Equal(PasswordHash.Sha512("divit", salt), protection.HashValue!.Value);
            Assert.NotEqual(PasswordHash.Sha512("Divit", salt), protection.HashValue!.Value);

            // "0" allows: only the two permissions given are written.
            Assert.False(protection.FormatColumns!.Value);
            Assert.False(protection.Sort!.Value);
            Assert.Null(protection.FormatCells);
            Assert.Null(protection.DeleteRows);
            Assert.True(protection.Sheet!.Value);

            var formats = part.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ToList();
            var input = formats[(int)sheet.Descendants<Cell>().Single(c => c.CellReference == "B3").StyleIndex!.Value];
            Assert.True(input.ApplyProtection!.Value);
            Assert.False(input.Protection!.Locked!.Value);
            var total = formats[(int)sheet.Descendants<Cell>().Single(c => c.CellReference == "B5").StyleIndex!.Value];
            Assert.True(total.Protection!.Hidden!.Value);
            Assert.Null(total.Protection.Locked);
            var label = formats[(int)sheet.Descendants<Cell>().Single(c => c.CellReference == "A3").StyleIndex!.Value];
            Assert.Null(label.Protection);

            var lockStructure = part.Workbook.WorkbookProtection!;
            Assert.True(lockStructure.LockStructure!.Value);
            Assert.Equal("SHA-512", lockStructure.WorkbookAlgorithmName!.Value);
            var bookSalt = Convert.FromBase64String(lockStructure.WorkbookSaltValue!.Value!);
            Assert.Equal(PasswordHash.Sha512("structure", bookSalt), lockStructure.WorkbookHashValue!.Value);
            Assert.NotEqual(protection.SaltValue!.Value, lockStructure.WorkbookSaltValue!.Value);
        }
    }

    [Fact]
    public void LocksWithoutPasswordWriteNoHash()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("One").Protect();
        book.Protect();
        var path = Save(book, "lock-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var protection = document.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProtection>()!;
            Assert.Null(protection.HashValue);
            Assert.Null(protection.AlgorithmName);
            Assert.True(protection.Sheet!.Value);
            var workbook = document.WorkbookPart.Workbook.WorkbookProtection!;
            Assert.Null(workbook.WorkbookHashValue);
            Assert.True(workbook.LockStructure!.Value);
        }
    }

    [Fact]
    public void EveryPermissionHasItsAttribute()
    {
        var book = new XlsxWorkbook();
        var all = Enum.GetValues(typeof(XlsxAllow)).Cast<XlsxAllow>().Aggregate((a, b) => a | b);
        book.AddSheet("One").Protect(null, all);
        var path = Save(book, "allow-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var p = document.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProtection>()!;
            var flags = new[]
            {
                p.FormatCells, p.FormatColumns, p.FormatRows, p.InsertColumns, p.InsertRows, p.InsertHyperlinks,
                p.DeleteColumns, p.DeleteRows, p.Sort, p.AutoFilter, p.PivotTables,
            };
            Assert.All(flags, f => Assert.False(f!.Value));
        }
    }

    [Fact]
    public void LockedFilesAreReproducibleWithACreationTime()
    {
        var folder = XlsxTestKit.NewFolder();
        var a = Path.Combine(folder, "a.xlsx");
        var b = Path.Combine(folder, "b.xlsx");
        OrderForm().Save(a);
        OrderForm().Save(b);
        Assert.Equal(File.ReadAllBytes(a), File.ReadAllBytes(b));
    }

    [Fact]
    public void WithoutACreationTimeEverySaveGetsAFreshSalt()
    {
        string SaltOf(string path)
        {
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                return document.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProtection>()!.SaltValue!.Value!;
            }
        }

        var book = new XlsxWorkbook();
        book.AddSheet("One").Protect("x");
        var folder = XlsxTestKit.NewFolder();
        book.Save(Path.Combine(folder, "a.xlsx"));
        book.Save(Path.Combine(folder, "b.xlsx"));
        Assert.NotEqual(SaltOf(Path.Combine(folder, "a.xlsx")), SaltOf(Path.Combine(folder, "b.xlsx")));
    }

    [Fact]
    public void SheetsWithTheSamePasswordGetDifferentSalts()
    {
        var book = new XlsxWorkbook { Created = XlsxTestKit.FixedTime };
        book.AddSheet("One").Protect("same");
        book.AddSheet("Two").Protect("same");
        var path = Save(book, "salt-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var hashes = document.WorkbookPart!.WorksheetParts.Select(p => p.Worksheet.GetFirstChild<SheetProtection>()!.HashValue!.Value).ToList();
            Assert.NotEqual(hashes[0], hashes[1]);
        }
    }
}

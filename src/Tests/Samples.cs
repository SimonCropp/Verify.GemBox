[TestFixture]
public class Samples
{
    [Test]
    public Task VerifyPdf() =>
        VerifyFile("sample.pdf");

    [Test]
    public Task VerifyPdfStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.pdf"));
        return Verify(stream, "pdf");
    }

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

    [Test]
    public Task VerifyWord() =>
        VerifyFile("sample.docx");

    [Test]
    public Task VerifyWordStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
        return Verify(stream, "docx");
    }
}
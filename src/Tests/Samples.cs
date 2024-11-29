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
}
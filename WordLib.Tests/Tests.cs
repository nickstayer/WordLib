using WordLib;

namespace WordLib.Tests;

public class Tests
{
    WordApp? _word;
    string? _dataDir;
    [SetUp]
    public void Setup()
    {
        _word = new WordApp(false);
        _dataDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\data\"));
    }

    [Test]
    public void TestOpenFile()
    {
        var testFile = _dataDir + "test.docx";
        if (!File.Exists(testFile))
        {
            throw new FileNotFoundException(testFile);
        }
        Assert.DoesNotThrow(() => _word?.OpenDoc(testFile));
    }

    [TearDown]
    public void TearDown() 
    {
        _word?.Quit();
    }
}
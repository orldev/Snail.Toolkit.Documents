using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>The sizes a zip declares are claims, and a bomb declares what it likes; the check counts what actually unpacks.</summary>
public sealed class ZipChaosTests
{
    /// <summary>
    /// The directory says a thousand bytes, the deflate stream holds sixty-four megabytes of zeros. Believed, the claim let the SDK buffer
    /// 47 MB before the zeros failed as XML. Counted before the SDK sees the file, the entry is refused within its first chunk — by the
    /// count passing a hundred times the file, or by .NET's zip reader finding more bytes than the directory declared, whichever comes
    /// first; either is a cheap refusal, which is what the limit promises. The reader is called on this thread so that only its allocations
    /// are measured, whatever the other tests are doing.
    /// </summary>
    [Fact]
    public async Task Read_ZipWhoseDirectoryUnderstatesItsSizes_IsRefusedCheaply()
    {
        byte[] lie = Hostility.LyingZip();
        long before = GC.GetAllocatedBytesForCurrentThread();

        var refused = await Record.ExceptionAsync(() => new WordReader().ReadAsync(new MemoryStream(lie)));

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(refused is DocumentTooLargeException or InvalidDataException, refused?.GetType().Name ?? "read as a document");
        Assert.True(allocated < 8 * 1024 * 1024, $"a zip of {lie.Length / 1024} KB claiming 1000 bytes was refused after allocating {allocated / 1_048_576} MB");
    }
}

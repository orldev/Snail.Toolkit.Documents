using System.Diagnostics;
using Microsoft.Extensions.Time.Testing;

namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>The half-minute deadline holds whatever the format's reader is doing: the reader runs on the pool, and the caller is answered at the deadline.</summary>
public sealed class DeadlineChaosTests
{
    /// <summary>
    /// Every reader but HTML runs synchronously inside <c>Task.FromResult</c>, so a reader busy in one step past the deadline is not
    /// stopped by its token. Run on the caller's thread, it finished and the answer was a document; run on the pool under the deadline,
    /// the answer is <see cref="ReadingFailure.TooSlow"/> at the deadline, and the step is left to end at its next check of the token.
    /// </summary>
    [Fact]
    public async Task Read_SynchronousReaderBusyInOneStepPastTheDeadline_IsRefusedAsTooSlow()
    {
        var clock = new FakeTimeProvider();
        var reader = new DocumentReader([new Stubborn(TimeSpan.FromSeconds(2))], clock);
        var deadlinePasses = Task.Run(async () =>
        {
            await Task.Delay(300);
            clock.Advance(TimeSpan.FromSeconds(ReadingLimits.MostReadingSeconds + 1));
        });

        var reading = await reader.ReadAsync(new MemoryStream(), "stubborn.txt");
        await deadlinePasses;

        Assert.Equal(ReadingFailure.TooSlow, reading.Failure);
    }

    /// <summary>
    /// <c>ReadAsync</c> returns its task at once: run on the caller's thread, a reader held that thread for the whole parse, and the caller's
    /// cancellation, the deadline and a second request on the same thread all waited behind it.
    /// </summary>
    [Fact]
    public async Task Read_SynchronousReader_ReturnsATaskOnlyAfterTheWholeParse()
    {
        var reader = new DocumentReader([new Stubborn(TimeSpan.FromSeconds(1))]);
        var clock = Stopwatch.StartNew();

        var reading = reader.ReadAsync(new MemoryStream(), "stubborn.txt");
        var untilTheTaskExists = clock.Elapsed;
        await reading;

        Assert.InRange(untilTheTaskExists, TimeSpan.Zero, TimeSpan.FromMilliseconds(200));
    }

    /// <summary>A reader that is busy for a while in one step and never looks at its token, as a parser of a hostile file is.</summary>
    private sealed class Stubborn(TimeSpan busy) : IFormatReader
    {
        public DocumentKind Kind => DocumentKind.Text;

        public IReadOnlyList<string> Extensions { get; } = [".txt"];

        public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default)
        {
            var clock = Stopwatch.StartNew();

            while (clock.Elapsed < busy)
                Thread.Sleep(10);

            return Task.FromResult<IReadOnlyList<Block>>([new Paragraph("done late")]);
        }
    }
}

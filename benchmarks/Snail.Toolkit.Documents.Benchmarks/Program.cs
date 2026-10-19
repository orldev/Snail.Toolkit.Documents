using BenchmarkDotNet.Running;
using Snail.Toolkit.Documents.Benchmarks;

BenchmarkSwitcher.FromAssembly(typeof(ReadingBenchmarks).Assembly).Run(args);

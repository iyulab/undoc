using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace Undoc.Tests;

/// <summary>
/// A test that can be skipped at runtime based on conditions.
/// </summary>
[XunitTestCaseDiscoverer("Undoc.Tests.SkippableFactDiscoverer", "Undoc.Tests")]
public class SkippableFactAttribute : FactAttribute
{
}

/// <summary>
/// Exception to indicate a test should be skipped.
/// </summary>
public class SkipException : Exception
{
    public SkipException(string reason) : base(reason) { }
}

/// <summary>
/// Helper class for skipping tests.
/// </summary>
public static class Skip
{
    /// <summary>
    /// Skip the test if the condition is true.
    /// </summary>
    public static void If(bool condition, string reason = "Test skipped")
    {
        if (condition)
            throw new SkipException(reason);
    }

    /// <summary>
    /// Skip the test if the condition is false.
    /// </summary>
    public static void IfNot(bool condition, string reason = "Test skipped")
    {
        if (!condition)
            throw new SkipException(reason);
    }
}

/// <summary>
/// Discoverer for skippable fact tests.
/// </summary>
public class SkippableFactDiscoverer : IXunitTestCaseDiscoverer
{
    private readonly IMessageSink _diagnosticMessageSink;

    public SkippableFactDiscoverer(IMessageSink diagnosticMessageSink)
    {
        _diagnosticMessageSink = diagnosticMessageSink;
    }

    public IEnumerable<IXunitTestCase> Discover(
        ITestFrameworkDiscoveryOptions discoveryOptions,
        ITestMethod testMethod,
        IAttributeInfo factAttribute)
    {
        yield return new SkippableTestCase(
            _diagnosticMessageSink,
            discoveryOptions.MethodDisplayOrDefault(),
            discoveryOptions.MethodDisplayOptionsOrDefault(),
            testMethod);
    }
}

/// <summary>
/// Test case that supports runtime skipping.
/// </summary>
public class SkippableTestCase : XunitTestCase
{
    [Obsolete("Called by the de-serializer; should only be called by deriving classes for de-serialization purposes")]
    public SkippableTestCase() { }

    public SkippableTestCase(
        IMessageSink diagnosticMessageSink,
        TestMethodDisplay defaultMethodDisplay,
        TestMethodDisplayOptions defaultMethodDisplayOptions,
        ITestMethod testMethod,
        object[]? testMethodArguments = null)
        : base(diagnosticMessageSink, defaultMethodDisplay, defaultMethodDisplayOptions, testMethod, testMethodArguments)
    {
    }

    public override async Task<RunSummary> RunAsync(
        IMessageSink diagnosticMessageSink,
        IMessageBus messageBus,
        object[] constructorArguments,
        ExceptionAggregator aggregator,
        CancellationTokenSource cancellationTokenSource)
    {
        var runner = new SkippableTestCaseRunner(
            this, DisplayName, SkipReason, constructorArguments,
            TestMethodArguments, messageBus, aggregator, cancellationTokenSource);
        return await runner.RunAsync();
    }
}

/// <summary>
/// Test case runner that handles skip exceptions.
/// </summary>
public class SkippableTestCaseRunner : XunitTestCaseRunner
{
    public SkippableTestCaseRunner(
        IXunitTestCase testCase,
        string displayName,
        string? skipReason,
        object[] constructorArguments,
        object[]? testMethodArguments,
        IMessageBus messageBus,
        ExceptionAggregator aggregator,
        CancellationTokenSource cancellationTokenSource)
        : base(testCase, displayName, skipReason, constructorArguments,
               testMethodArguments, messageBus, aggregator, cancellationTokenSource)
    {
    }

    protected override async Task<RunSummary> RunTestAsync()
    {
        try
        {
            return await base.RunTestAsync();
        }
        catch (SkipException e)
        {
            var test = new XunitTest(TestCase, DisplayName);
            if (!MessageBus.QueueMessage(new TestSkipped(test, e.Message)))
                CancellationTokenSource.Cancel();
            return new RunSummary { Skipped = 1, Total = 1 };
        }
    }
}

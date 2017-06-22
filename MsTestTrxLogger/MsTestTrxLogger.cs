using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;

namespace MsTestTrxLogger
{
    [ExtensionUri("logger://MsTestTrxLogger/v1")]
    [FriendlyName("MsTestTrxLogger")]
    public class MsTestTrxLogger : ITestLogger
    {
        /// <summary>
        /// Entry point for the logger
        /// </summary>
        public void Initialize(TestLoggerEvents events, string testRunDirectory)
        {
            Console.WriteLine("Initializing MsTestTrxLogger...");
            Console.WriteLine($"Running in: {testRunDirectory}");

            var testRunStarted = DateTime.Now;
            var testResults = new List<TestResult>();

            events.TestResult += (sender, eventArgs) =>
            {
                if (!IsTestIgnored(eventArgs.Result))
                {
                    testResults.Add(eventArgs.Result);
                }
            };

            events.TestRunMessage += (sender, args) =>
            {
                Console.WriteLine($"[{args.Level}] {args.Message}");
            };

            events.TestRunComplete += (sender, args) =>
            {
                var writer = new MsTestTrxWriter(testResults, args, testRunStarted);
                writer.DumpOutput(testRunDirectory);
            };
        }

        /// <summary>
        /// Returns whether the test was ignored or not.
        /// </summary>
        /// <remarks>
        /// The object model doesn't indicate whether a test was ignored with an IgnoreAttribute, or was skipped for other reasons.
        /// It seems to be a reliable way to recognize if a test was actually ignored if we check whether the number of its messages is 0.
        /// </remarks>
        private static bool IsTestIgnored(TestResult test)
        {
            return test.Outcome == TestOutcome.Skipped && test.Messages.Count == 0;
        }
    }
}
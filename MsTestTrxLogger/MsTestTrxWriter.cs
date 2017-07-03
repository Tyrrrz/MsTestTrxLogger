using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MsTestTrxLogger
{
    public class MsTestTrxWriter
    {
        private const string AdapterTypeName =
                "Microsoft.VisualStudio.TestTools.TestTypes.Unit.UnitTestAdapter, Microsoft.VisualStudio.QualityTools.Tips.UnitTest.Adapter, Version=12.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";

        private readonly IList<TestResult> _testResults;
        private readonly TestRunCompleteEventArgs _completeEventArgs;
        private readonly DateTime _testRunStarted;

        private readonly Dictionary<TestResult, Guid> _executionIds;
        private readonly Dictionary<string, Assembly> _assemblies;

        public MsTestTrxWriter(IList<TestResult> testResults, TestRunCompleteEventArgs completeEventArgs, DateTime testRunStarted)
        {
            _testResults = testResults;
            _completeEventArgs = completeEventArgs;
            _testRunStarted = testRunStarted;
            _executionIds = new Dictionary<TestResult, Guid>();
            _assemblies = new Dictionary<string, Assembly>();
        }

        /// <summary>
        /// Clean off empty namespaces
        /// </summary>
        private void CleanXmlNamespaces(XDocument doc)
        {
            if (doc?.Root == null)
                return;

            foreach (var node in doc.Root.Descendants())
            {
                // Sanity check
                if (node.Parent == null)
                    continue;

                // Clean namespace attribute
                if (node.Name.NamespaceName == "")
                {
                    node.Attributes("xmlns").Remove();
                    node.Name = node.Parent.Name.Namespace + node.Name.LocalName;
                }
            }
        }

        /// <summary>
        /// Calculates a hash of the string and copies the first 128 bits of the hash to a new Guid.
        /// </summary>
        private Guid GuidFromString(string data)
        {
            using (var hasher = new SHA1CryptoServiceProvider())
            {
                var raw = Encoding.Unicode.GetBytes(data);
                var hash = hasher.ComputeHash(raw);

                return new Guid(hash.Take(16).ToArray());
            }
        }

        /// <summary>
        /// Returns the output tag containing both the normal test output, and the error message (if there was any).
        /// </summary>
        private XElement CreateOutputXml(TestResult result)
        {
            var element = new XElement("Output",
                new XElement("StdOut",
                    string.Join(Environment.NewLine, result.Messages.Select(m => m.Text).ToArray())));

            if (!string.IsNullOrEmpty(result.ErrorMessage) || !string.IsNullOrEmpty(result.ErrorStackTrace))
            {
                element.Add(new XElement("ErrorInfo",
                    new XElement("Message", result.ErrorMessage),
                    new XElement("StackTrace", result.ErrorStackTrace)));
            }

            return element;
        }

        /// <summary>
        /// Returns the execution id for the given test.
        /// </summary>
        /// <remarks>
        /// The execution ids can be generated randomly, but we have to use the same id for a given test in multiple places in the XML.
        /// Hence the ids are stored in a dictionary for every test.
        /// </remarks>
        private Guid GetExecutionId(TestResult result)
        {
            if (_executionIds.TryGetValue(result, out var id))
                return id;
            return _executionIds[result] = Guid.NewGuid();
        }

        /// <summary>
        /// Loads the assembly from <paramref name="path" /> and stores its reference so that we don't load an assembly multiple times.
        /// </summary>
        private Assembly GetAssembly(string path)
        {
            if (_assemblies.TryGetValue(path, out var assembly))
                return assembly;
            return _assemblies[path] = Assembly.LoadFrom(path);
        }

        /// <summary>
        /// Returns the full description text of the unit test.
        /// </summary>
        /// <remarks>
        /// The description text (specified with <see cref="DescriptionAttribute" />) is not present
        /// in the object model provided in <see cref="Microsoft.VisualStudio.TestPlatform.ObjectModel" />, so we have to look it up in the unit test assembly.
        /// </remarks>
        private string GetDescription(TestResult test)
        {
            var assembly = GetAssembly(test.TestCase.Source);

            var className = test.TestCase.FullyQualifiedName.Substring(0, test.TestCase.FullyQualifiedName.LastIndexOf('.'));
            var methodName = test.TestCase.FullyQualifiedName.Substring(test.TestCase.FullyQualifiedName.LastIndexOf('.') + 1);

            var type = assembly.GetType(className);
            var method = type.GetMethod(methodName);
            var attributes = method.GetCustomAttributes<DescriptionAttribute>().ToList();

            return attributes.FirstOrDefault()?.Description ?? test.TestCase.DisplayName;
        }

        /// <summary>
        /// Returns the list of TestPropertyAttributes specified for <paramref name="test" />.
        /// </summary>
        /// <remarks>
        /// The information in the TestPropertyAttributes (<see cref="Microsoft.VisualStudio.TestTools.UnitTesting.TestPropertyAttribute" />) is not present
        /// in the object model provided in Microsoft.VisualStudio.TestPlatform.ObjectModel, so we have to look it up in the unit test assembly.
        /// </remarks>
        private IEnumerable<TestPropertyAttribute> GetPropertyAttributes(TestResult test)
        {
            var assembly = GetAssembly(test.TestCase.Source);

            var className = test.TestCase.FullyQualifiedName.Substring(0, test.TestCase.FullyQualifiedName.LastIndexOf('.'));
            var methodName = test.TestCase.FullyQualifiedName.Substring(test.TestCase.FullyQualifiedName.LastIndexOf('.') + 1);

            var type = assembly.GetType(className);
            var method = type.GetMethod(methodName);

            return method.GetCustomAttributes<TestPropertyAttribute>();
        }

        /// <summary>
        /// Returns the list of TestCategoryAttributes specified for <paramref name="test" />.
        /// </summary>
        /// <remarks>
        /// The information in the TestCategoryAttributes (<see cref="TestCategoryAttribute" />) is not present
        /// in the object model provided in <see cref="Microsoft.VisualStudio.TestPlatform.ObjectModel" />, so we have to look it up in the unit test assembly.
        /// </remarks>
        private IEnumerable<TestCategoryAttribute> GetTestCategory(TestResult test)
        {
            var assembly = GetAssembly(test.TestCase.Source);

            var className = test.TestCase.FullyQualifiedName.Substring(0, test.TestCase.FullyQualifiedName.LastIndexOf('.'));
            var methodName = test.TestCase.FullyQualifiedName.Substring(test.TestCase.FullyQualifiedName.LastIndexOf('.') + 1);

            var type = assembly.GetType(className);
            var method = type.GetMethod(methodName);

            return method.GetCustomAttributes<TestCategoryAttribute>();
        }

        /// <summary>
        /// Returns the fully qualified assembly name of the unit test class.
        /// </summary>
        /// <remarks>
        /// This information is not present in the object model provided in <see cref="Microsoft.VisualStudio.TestPlatform.ObjectModel" />, so we have to look it up in the unit test assembly.
        /// </remarks>
        private string GetClassFullName(TestResult test)
        {
            var assembly = GetAssembly(test.TestCase.Source);

            var className = test.TestCase.FullyQualifiedName.Substring(0, test.TestCase.FullyQualifiedName.LastIndexOf('.'));
            var type = assembly.GetType(className);

            return type.AssemblyQualifiedName;
        }

        /// <summary>
        /// Dumps the test results to the given directory
        /// </summary>
        public void DumpOutput(string outputDirectoryPath)
        {
            Console.WriteLine("Started generating TRX output...");

            var testRunId = Guid.NewGuid();

            // Generate XML structure
            var ns = XNamespace.Get("http://microsoft.com/schemas/VisualStudio/TeamTest/2010");
            var doc = new XDocument(new XDeclaration("1.0", "UTF-8", null),
                new XElement(ns + "TestRun",
                    new XAttribute("id", testRunId.ToString()),
                    new XAttribute("name", $"{Environment.UserName}@{Environment.MachineName} {DateTime.UtcNow}"),
                    new XAttribute("runUser", $@"{Environment.UserDomainName}\{Environment.UserName}"),
                    new XElement("Results",
                        _testResults.Select(result => new XElement("UnitTestResult",
                            new XAttribute("computerName", Environment.MachineName),
                            new XAttribute("duration", result.Duration.ToString()),
                            new XAttribute("endTime", result.EndTime.ToString("o")),
                            new XAttribute("executionId", GetExecutionId(result)),
                            new XAttribute("outcome", result.Outcome == TestOutcome.Skipped ? "Inconclusive" : result.Outcome.ToString()),
                            new XAttribute("relativeResultsDirectory", GetExecutionId(result)),
                            new XAttribute("startTime", result.StartTime.ToString("o")),
                            new XAttribute("testId", GuidFromString(result.TestCase.FullyQualifiedName)),
                            new XAttribute("testListId", testRunId.ToString()),
                            new XAttribute("testName", result.TestCase.DisplayName),
                            new XAttribute("testType", testRunId.ToString()),
                            CreateOutputXml(result)))),
                    new XElement("ResultSummary",
                        new XAttribute("outcome", _completeEventArgs.IsAborted ? "Aborted" : _completeEventArgs.IsCanceled ? "Canceled" : "Completed"),
                        new XElement("Counters",
                            new XAttribute("aborted", 0),
                            new XAttribute("completed", 0),
                            new XAttribute("disconnected", 0),
                            new XAttribute("error", 0),
                            new XAttribute("executed", _testResults.Count(r => r.Outcome != TestOutcome.Skipped)),
                            new XAttribute("failed", _testResults.Count(r => r.Outcome == TestOutcome.Failed)),
                            new XAttribute("inconclusive", _testResults.Count(r => r.Outcome == TestOutcome.Skipped || r.Outcome == TestOutcome.NotFound || r.Outcome == TestOutcome.None)),
                            new XAttribute("inProgress", 0),
                            new XAttribute("notExecuted", _testResults.Count(r => r.Outcome == TestOutcome.Skipped)),
                            new XAttribute("notRunnable", 0),
                            new XAttribute("passed", _testResults.Count(r => r.Outcome == TestOutcome.Passed)),
                            new XAttribute("passedButRunAborted", 0),
                            new XAttribute("pending", 0),
                            new XAttribute("timeout", 0),
                            new XAttribute("total", _testResults.Count),
                            new XAttribute("warning", 0))),
                      new XElement("TestDefinitions",
                        _testResults.Select(result => new XElement("UnitTest",
                            new XAttribute("id", GuidFromString(result.TestCase.FullyQualifiedName)),
                            new XAttribute("name", result.TestCase.DisplayName),
                            new XAttribute("storage", result.TestCase.Source),
                            new XElement("Description", GetDescription(result)),
                            new XElement("Execution", new XAttribute("id", GetExecutionId(result))),
                            new XElement("Properties",
                                GetPropertyAttributes(result).Select(p => new XElement("Property",
                                    new XElement("Key", p.Name),
                                    new XElement("Value", p.Value)))),
                            new XElement("TestCategory",
                                GetTestCategory(result).Select(c => new XElement("TestCategoryItem", c.TestCategories.First()))),
                            new XElement("TestMethod",
                                new XAttribute("adapterTypeName", AdapterTypeName),
                                new XAttribute("className", GetClassFullName(result)),
                                new XAttribute("codeBase", result.TestCase.Source),
                                new XAttribute("name", result.TestCase.DisplayName))))),
                      new XElement("TestEntries",
                        _testResults.Select(result => new XElement("TestEntry",
                            new XAttribute("executionId", GetExecutionId(result)),
                            new XAttribute("testId", GuidFromString(result.TestCase.DisplayName)),
                            new XAttribute("testListId", testRunId.ToString())))),
                      new XElement("TestLists",
                        new XElement("TestList", new XAttribute("id", testRunId.ToString()), new XAttribute("name", "Test list"))),
                      new XElement("Times",
                        new XAttribute("creation", _testRunStarted.ToString("o")),
                        new XAttribute("finish", DateTime.Now.ToString("o")),
                        new XAttribute("queueing", _testRunStarted.ToString("o")),
                        new XAttribute("start", _testRunStarted.ToString("o")))
                    ));

            CleanXmlNamespaces(doc);

            Console.WriteLine("Finished generating TRX output...");

            // Save to file
            string fileName =
                $"{Environment.UserName}_{Environment.MachineName} {DateTime.Now:yyyy-MM-dd HH_mm_ss}.trx";
            string filePath = Path.Combine(outputDirectoryPath, fileName);
            doc.Save(filePath);

            Console.WriteLine($"Saved to: {filePath}");
        }
    }
}
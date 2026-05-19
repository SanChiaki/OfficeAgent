using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateVersionComparerTests
    {
        [Theory]
        [InlineData("1.0.176", "1.0.175", true)]
        [InlineData("v1.0.176", "1.0.175", true)]
        [InlineData("1.1.0", "1.0.999", true)]
        [InlineData("2.0.0", "1.9.999", true)]
        [InlineData("1.0.175", "1.0.175", false)]
        [InlineData("1.0.174", "1.0.175", false)]
        [InlineData("not-a-version", "1.0.175", false)]
        [InlineData("1.0.176", "not-a-version", false)]
        [InlineData("", "1.0.175", false)]
        [InlineData(null, "1.0.175", false)]
        public void IsNewerThanCurrentComparesSupportedVersions(string latestVersion, string currentVersion, bool expected)
        {
            Assert.Equal(expected, UpdateVersionComparer.IsNewerThanCurrent(latestVersion, currentVersion));
        }
    }
}

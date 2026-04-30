using OfficeAgent.ExcelAddIn;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ProjectSearchMatcherTests
    {
        [Fact]
        public void IsMatchReturnsTrueWhenQueryAppearsInsideProjectLabel()
        {
            Assert.True(ProjectSearchMatcher.IsMatch("delivery-tracker-交付跟踪项目", "tracker"));
            Assert.True(ProjectSearchMatcher.IsMatch("delivery-tracker-交付跟踪项目", "TRACK"));
        }

        [Fact]
        public void IsMatchReturnsTrueWhenAllQueryTermsAppearInAnyOrder()
        {
            Assert.True(ProjectSearchMatcher.IsMatch("customer-onboarding-客户上线项目", "onboard customer"));
        }

        [Fact]
        public void IsMatchReturnsTrueForNonContiguousCharacterQuery()
        {
            Assert.True(ProjectSearchMatcher.IsMatch("delivery-tracker-交付跟踪项目", "dt"));
        }

        [Fact]
        public void IsMatchReturnsFalseWhenQueryDoesNotMatchProjectLabel()
        {
            Assert.False(ProjectSearchMatcher.IsMatch("delivery-tracker-交付跟踪项目", "performance"));
        }
    }
}

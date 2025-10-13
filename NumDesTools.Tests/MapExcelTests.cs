using System.Collections.Generic;
using Xunit;
using NumDesTools;

namespace NumDesTools.Tests
{
    public class MapExcelTests
    {
        [Fact]
        public void TablePathFix_Localization_MapsCorrectly()
        {
            var workbookPath = "C:\\Projects\\Test";
            var input = "Localizations.xlsx";
            var output = MapExcel.TablePathFix(input, workbookPath);
            Assert.Equal("C:\\Projects\\Test\\Localizations\\Localizations.xlsx", output);
        }

        [Fact]
        public void TablePathFix_ClondikeComposite_MapsCorrectly()
        {
            var workbookPath = "C:\\Projects\\Test";
            var input = "something���ʴ���##sub#sheet"; // contains ���ʴ���## and $
            // Simulate case with $ in string
            input = "abc���ʴ���##sub$sheet";
            var output = MapExcel.TablePathFix(input, workbookPath);
            Assert.Contains("���ʴ���", output);
        }
    }
}

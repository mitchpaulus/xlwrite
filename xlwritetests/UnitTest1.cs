using NUnit.Framework;
using System.Text.RegularExpressions;
using xlwrite;

namespace Tests
{
    public class Tests
    {
        [Test]
        public void CellReferenceParser()
        {
            string a1Reference = "B23";
            bool success  = XlWriteUtilities.TryParseCellReference(a1Reference, out Cell cell);

            Assert.AreEqual(null, cell.SheetName);
            Assert.AreEqual(-1, cell.SheetNum);
            Assert.AreEqual(2, cell.Column);
            Assert.AreEqual(23, cell.Row);

            string r1c1Reference = "r12c65";
            success = XlWriteUtilities.TryParseCellReference(r1c1Reference, out cell);
            Assert.AreEqual(65, cell.Column);
            Assert.AreEqual(12, cell.Row);
        }

        [Test]
        public void NamedWorksheetTest()
        {
            string namedWorksheetReference = "'Sheet 1'!B23";
            bool success  = XlWriteUtilities.TryParseCellReference(namedWorksheetReference, out Cell cell);
            Assert.AreEqual("Sheet 1", cell.SheetName);
            Assert.AreEqual(-1, cell.SheetNum);
            Assert.AreEqual(2, cell.Column);
            Assert.AreEqual(23, cell.Row);
        }

        [Test]
        public void NumberedSheetTest()
        {
            string numberSheetReference = "1!B23";
            bool success  = XlWriteUtilities.TryParseCellReference(numberSheetReference, out Cell cell);
            Assert.AreEqual(null, cell.SheetName);
            Assert.AreEqual(1, cell.SheetNum);
            Assert.AreEqual(2, cell.Column);
            Assert.AreEqual(23, cell.Row);
        }

        [Test]
        public void RegexTests()
        {
            string test = "]";

            Regex regex = new Regex(@"[\]]");

            Assert.IsTrue(regex.Match(test).Success);


        }
    }
}
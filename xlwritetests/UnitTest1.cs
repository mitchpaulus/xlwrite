using NUnit.Framework;
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

            Assert.AreEqual(2, cell.Column);
            Assert.AreEqual(23, cell.Row);

            string r1c1Reference = "r12c65";
            success = XlWriteUtilities.TryParseCellReference(r1c1Reference, out cell);
            Assert.AreEqual(65, cell.Column);
            Assert.AreEqual(12, cell.Row);
        }
    }
}
using Moq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcel
{
    abstract class IOExcel
    {
        protected Elements elements;
        protected ILetterConvertor convertor;
        private Alphabet alphabet;

        protected Point[] range;

        public IOExcel(ref Elements elements)
        {
            this.elements = elements;
            alphabet = new EN();
            convertor = new ExcelNumberConverter(alphabet);

            range = initPoints(new Point[2]);
        }
        private Point[] initPoints(Point[] points)
        {
            for (var i = 0; i < points.Length; i++)
            {
                points[i] = new Point(ref convertor);
            }

            return points;
        }

        protected abstract int getColumnsCount();
        protected abstract int getCellsCount();
    }

    class Reader : IOExcel
    {
        public Reader(ref Elements elements) : base(ref elements)
        { }

        public object[,] Read()
        {
            range[0].set(1, 1);
            return Read(range[0]);
        }
        public object[,] Read(Point start)
        {
            range[0] = start;
            range[1].set(getColumnsCount(), getCellsCount());
            return Read(range[0], range[1]);
        }
        public object[,] Read(Point start, Point end)
        {
            var table = elements.sheet.Range[start.get(), end.get()].Value2;
            return table;
        }

        protected override int getCellsCount()
        {
            return elements.sheet.UsedRange.Columns.Count;
        }
        protected override int getColumnsCount()
        {
            return elements.sheet.UsedRange.Cells.Count;
        }
    }

    class Writer : IOExcel
    {
        object[,] table;

        public Writer(ref Elements elements) : base(ref elements)
        { }

        public void Write(object[,] table)
        {
            range[0].set(1, 1);
            Write(table, range[0]);
        }
        public void Write(object[,] table, Point start)
        {
            this.table = table;
            checkOnChar();
            range[0] = start;
            range[1].set(getColumnsCount() + dataColumnsLength(), getCellsCount() + dataCellsLength());
            elements.sheet.Range[range[0].get(), range[1].get()].Value2 = table;
        }

        private void checkOnChar()
        {
            for (int i = 0; i < table.GetLength(0); i++)
            {
                for (int j = 0; j < table.GetLength(1); j++)
                {
                    if (table[i, j] != null && table[i, j].GetType() == typeof(char))
                    {
                        throw new Exception("Excel преобразует char бувы('E') в цифры. При восстановлении данные не сойдутся.");
                    }
                }
            }
        }

        protected override int getCellsCount()
        {
            return table.GetLength(0);
        }
        private int dataCellsLength()
        {
            return range[0].y - 1;
        }
        protected override int getColumnsCount()
        {
            return table.GetLength(1);
        }
        private int dataColumnsLength()
        {
            return range[0].x - 1;
        }
    }

    [TestFixture]
    class IOOperations
    {
        Writer writer;
        Reader reader;
        Elements elements;

        ILetterConvertor convertor;
        Alphabet alphabet;

        object[,] data = new object[,] { { "Name", "Second Name", "Account id" },
                { "Qml", "DNSmasq", 902 },
                { null, null, null },
                {"D",218,01 },
                {123,123,123 },
                { null, true,912639}
            };
        string docName = "test.xlsx";


        [SetUp]
        public void Start()
        {
            elements = new Elements();

            writer = new Writer(ref elements);
            reader = new Reader(ref elements);

            alphabet = new EN();
            convertor = new ExcelNumberConverter(alphabet);

            elements.Start();
        }

        [Test, TestCaseSource(nameof(checkOnCharData))]
        public void checkOnChar(object[,] obj, bool result)
        {
            // Excel преобразует char в int
            elements.createWorkbook();
            elements.createSheet(1);
            elements.setWorksheet(1);

            if (result)
            {
                Assert.DoesNotThrow(() => { writer.Write(obj); });
            }
            else
            {
                Assert.Throws<Exception>(() => { writer.Write(obj); });
            }
        }
        private static IEnumerable<TestCaseData> checkOnCharData()
        {
            yield return new TestCaseData(new object[,] {
                    { 1, true, "E"},
                    { 12,99,00},
                    { "Word", "Hello",null },
                    { "R", "W", "X"} ,
                    { "&&&&", "{}", "123"},
                    { "true", "false", false},
                    { 1.2, 0.06 ,902.5 } }, true);
            yield return new TestCaseData(new object[,] { { 'R', '*', '&' } }, false);
            yield return new TestCaseData(new object[,] {
                    { 1, true, "E"},
                    { 12,'7',00},
                    { "Word", "Hello",null },
                    { "R", "W", "X"} , }, false);
        }

        [TestCase(1, 1, 3, 7)]
        [TestCase(100, 1, 1, 700)]
        [TestCase(100, 100, 3000, 1)]
        [TestCase(100, 99, 5000, 1000)]
        public void writeAndReadInDifferentPlaces(int sheets, int activeSheet, int x, int y)
        {
            elements.createWorkbook();
            elements.createSheet(sheets);
            elements.setWorksheet(activeSheet);

            var point = new Point(ref convertor);
            var point1 = new Point(ref convertor);
            point.set(x, y);
            point1.set(x + data.GetLength(1) - 1, y + data.GetLength(0) - 1);

            writer.Write(data, point);
            elements.Save(docName);

            elements.Open(docName);
            elements.setWorksheet(activeSheet);
            var result = reader.Read(point, point1);

            Assert.AreEqual(result, data);
        }

        [TearDown]
        public void Stop()
        {
            elements.Stop();
        }
    }
}

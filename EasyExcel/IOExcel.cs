using Moq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcel
{
    public abstract class IOExcel
    {
        protected Elements elements;
        protected ILetterConvertor convertor;
        private Alphabet alphabet;

        protected Point[] range;

        public IOExcel(ref Elements elements)
        {
            this.elements = elements;
            alphabet = new EN();
            convertor = new NumberConverter(alphabet);

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

        public abstract int getColumnsCount();
        public abstract int getRowsCount();

        public Point createPoint()
        {
            var lang = new EN();
            var converter = new NumberConverter(lang);
            return new Point(ref convertor);
        }
    }

    public class Reader : IOExcel
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
            range[1].set(getColumnsCount(), getRowsCount());
            return Read(range[0], range[1]);
        }
        public object[,] Read(Point start, Point end)
        {
            var table = elements.sheet.Range[start.get(), end.get()].Value2;
            if (table!=null) {
                return convertToNormal(table);
            }
            return null;
        }
        private object[,] convertToNormal(object[,] excel)
        {
            // Преобразует массив так, чтобы он начинался с нуля
            int x = excel.GetLength(0);
            int y = excel.GetLength(1);
            object[,] mas = new object[x, y];

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    mas[i, j] = excel[i + 1, j + 1];
                }
            }
            return mas;
        }

        public override int getRowsCount()
        {
            return elements.sheet.UsedRange.Rows.Count;
        }
        public override int getColumnsCount()
        {
            return elements.sheet.UsedRange.Columns.Count;
        }
    }

    public class Writer : IOExcel
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
            if (table!=null) {
                checkOnChar();
                range[0] = start;
                range[1].set(getColumnsCount() + dataColumnsLength(), getRowsCount() + dataCellsLength());
                elements.sheet.Range[range[0].get(), range[1].get()].Value2 = table;
            }
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

        public override int getRowsCount()
        {
            return table.GetLength(0);
        }
        private int dataCellsLength()
        {
            return range[0].y - 1;
        }
        public override int getColumnsCount()
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

        string path = Environment.CurrentDirectory + "/";

        object[,] data = new object[,] { { "Name", "Second Name", "Account id" },
                { "Qml", "DNSmasq", 902 },
                { null, null, null },
                {"D",218,01 },
                {123,123,123 },
                { null, true,912639}
            };

        // data have not null objects
        object[,] dataFull = new object[,] { { "Name", "Second Name", "Account id" },
                { "Qml", "DNSmasq", 902 },
                {"D",218,01 },
                {123,123,123 },
                { "test", true,912639}
            };

        string docName = "test.xlsx";


        [SetUp]
        public void Start()
        {
            elements = new Elements();

            writer = new Writer(ref elements);
            reader = new Reader(ref elements);

            alphabet = new EN();
            convertor = new NumberConverter(alphabet);

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
            elements.Save(path+docName);

            elements.Open(path+docName);
            elements.setWorksheet(activeSheet);
            var result = reader.Read(point, point1);

            Assert.AreEqual(result, data);
        }

        [Test]
        public void editDocument()
        {
            elements.Open(path+docName);
            elements.setWorksheet(1);

            Point point = reader.createPoint();
            point.set(5, 5);

            Assert.DoesNotThrow(()=> { 
                var result = reader.Read();
                writer.Write(result, point);

                elements.Save(path+docName);
            });

        }

        [Test]
        public void EmptyRead()
        {
            elements.createWorkbook();
            elements.Save(path+docName);

            elements.Stop();
            elements.Start();

            elements.Open(path+docName);
            elements.setWorksheet(1);

            Assert.DoesNotThrow(() => { 
                var result = reader.Read();
                Assert.AreEqual(result, null);
            });
        }

        [Test]
        public void Read()
        {
            createDefaultDocument();

            elements.Stop();
            elements.Start();

            elements.Open(path + docName);
            elements.setWorksheet(1);
            var result = reader.Read();

            Assert.AreEqual(result, dataFull);
        }
        [Test]
        public void SecondRead()
        {
            createDefaultDocument();

            elements.Stop();
            elements.Start();

            elements.Open(path + docName);
            elements.setWorksheet(1);

            var point = reader.createPoint();
            point.set(1, 1);
            var result = reader.Read(point);

            Assert.AreEqual(result, dataFull);
        }
        private void createDefaultDocument()
        {
            elements.createWorkbook();
            elements.setWorksheet(1);
            writer.Write(dataFull);
            elements.Save(path + docName);
        }

        [TearDown]
        public void Stop()
        {
            elements.Stop();
        }
    }
}

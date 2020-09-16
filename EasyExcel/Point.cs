using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Moq;
using NUnit.Framework;


namespace EasyExcel
{
    public class Point
    {
        ILetterConvertor converter;

        public int x { get; set; }
        public int y { get; set; }

        public Point(ref ILetterConvertor converter)
        {
            this.converter = converter;
        }
        public string get()
        {
            return converter.Convert(x - 1) + y;
        }
        public void set(int x, int y)
        {
            this.x = x;
            this.y = y;
        }

    }

    [TestFixture]
    class PointTest
    {
        Alphabet en;
        Alphabet ru;
        ILetterConvertor convertor;
        Point point;

        [OneTimeSetUp]
        public void Start()
        {
            en = new EN();
            ru = new RU();
            convertor = new NumberConverter(en);
            point = new Point(ref convertor);
        }

        [TestCase(1, 1, "A1")]
        [TestCase(1, 8, "A8")]
        [TestCase(11, 1, "K1")]
        [TestCase(18, 9, "R9")]
        [TestCase(44, 9, "AR9")]
        [TestCase(57, 90, "BE90")]
        public void positivePoint(int x, int y, string result)
        {
            point.set(x, y);
            Assert.AreEqual(point.get(), result);
        }

        [TestCase(0, 0)]
        [TestCase(-18, 9)]
        [TestCase(null, null)]
        public void negativePoint(int x, int y)
        {
            point.set(x, y);
            var ex = Assert.Catch<Exception>(() => point.get());
            Assert.AreEqual("Значение не может быть меньше нуля.", ex.Message);
        }

        [Test]
        public void notIntegratedTest()
        {
            var converter = new Mock<ILetterConvertor>();
            converter.Setup(x => x.Convert(17)).Returns("R");
            var converterObj = converter.Object;

            var point = new Point(ref converterObj);
            point.set(18,9);

            Assert.AreEqual(point.get(), "R9");
        }

        [OneTimeTearDown]
        public void Stop()
        {
            convertor = null;
        }

    }

}

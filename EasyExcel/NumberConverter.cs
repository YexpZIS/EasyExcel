using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Moq;
using NUnit.Framework;

namespace EasyExcel
{
    public interface ILetterConvertor
    {
        string Convert(int number);
        void setAlphabet(Alphabet alphabet);
    }

    public class NumberConverter : ILetterConvertor
    {
        private char[] letters;

        public NumberConverter(Alphabet alphabet)
        {
            setAlphabet(alphabet);
        }

        public void setAlphabet(Alphabet alphabet)
        {
            int[] range = alphabet.getRange();
            checkRange(range);
            fillLetters(range);

        }
        private void fillLetters(int[] range)
        {
            int from = range[0];
            int to = range[1];

            letters = new char[to - from + 1];

            for (int i = from; i <= to; i++)
            {
                letters[i - from] = (char)i;
            }
        }
        private void checkRange(int[] range)
        {
            // 1. На содержание 2 значений в массиве
            if (range.Length < 2)
            {
                throw new Exception("Эффективная длинна массива = 2 ");
            }

            // 2. На положительность данных значений
            for (var i = 0; i < 2; i++)
            {
                if (range[i] < 0)
                {
                    throw new Exception("В минусовом диапазоне нет букв.");
                }

            }

            // 3. Первое число всегда должно быть меньше второго
            if (range[0] >= range[1])
            {
                throw new Exception("Данный диапазон не поддерживается(example: new int[65, 90]).");
            }

        }

        public string Convert(int number)
        {
            if (number < 0)
            {
                throw new Exception("Значение не может быть меньше нуля.");
            }

            return ConvertToLetters(number);
        }

        private string ConvertToLetters(int number)
        {
            string result = "";

            int alphabetLength = letters.Length;
            int c = 0;
            do
            {
                c = number % alphabetLength;
                result = letters[c] + result;

                number = number / alphabetLength;
                number--;
            }
            while (number > alphabetLength);

            if (number >= 0)
            {
                result = letters[number] + result;
            }

            return result;
        }


    }



    [TestFixture]
    class NumberConverterTest
    {
        ILetterConvertor convertor;
        Alphabet en;
        Alphabet ru;

        [OneTimeSetUp]
        public void Start()
        {
            en = new EN();
            ru = new RU();
            convertor = new NumberConverter(en);
        }

        [TestCase(0, "A")]
        [TestCase(10, "K")]
        [TestCase(25, "Z")]
        [TestCase(26, "AA")]
        [TestCase(54, "BC")]
        [TestCase(682, "ZG")]
        [TestCase(16383, "XFD")]
        public void ENpositiveNumbers(int data, string result)
        {
            convertor.setAlphabet(en);
            string a = convertor.Convert(data);
            Assert.AreEqual(a, result);
        }

        [TestCase(0, "А")]
        [TestCase(10, "К")]
        [TestCase(31, "Я")]
        [TestCase(32, "АА")]
        [TestCase(64, "БА")]
        [TestCase(501, "ОХ")]
        [TestCase(16383, "ОЮЯ")]
        public void RUpositiveNumbers(int data, string result)
        {
            convertor.setAlphabet(ru);
            string a = convertor.Convert(data);
            Assert.AreEqual(a, result);
        }

        [TestCase(-1)]
        [TestCase(-10000)]
        [TestCase(-893498)]
        public void negativeNumbers(int data)
        {
            var ex = Assert.Catch<Exception>(() => convertor.Convert(data));
            Assert.IsNotEmpty(ex.ToString());
        }


        [TestCase(0, false)]
        [TestCase(1, false)]
        [TestCase(2, true)]
        [TestCase(99, true)]
        public void changeMassiveSize(int size, bool result)
        {
            var alphabet = new Mock<Alphabet>();
            int[] mas = fillMassive(new int[size]);
            alphabet.Setup(a => a.getRange()).Returns(mas);

            if (result)
            {

                Assert.DoesNotThrow(() => setAlphabet(alphabet));
            }
            else
            {
                Assert.Throws<Exception>(() => setAlphabet(alphabet));
            }
        }
        private int[] fillMassive(int[] mas)
        {
            for (var i = 0; i < mas.Length; i++)
            {
                mas[i] = i;
            }
            return mas;
        }


        [TestCase(0, 0, false)]
        [TestCase(1, 1, false)]
        [TestCase(27, 100, true)]
        [TestCase(0, 3000, true)]
        [TestCase(-300, 40, false)]
        [TestCase(null, null, false)]
        public void positiveNumbersInMassive(int a, int b, bool result)
        {
            var alphabet = new Mock<Alphabet>();
            alphabet.Setup(x => x.getRange()).Returns(new int[] { a, b });

            if (result)
            {
                Assert.DoesNotThrow(() => setAlphabet(alphabet));
            }
            else
            {
                Assert.Throws<Exception>(() => setAlphabet(alphabet));
            }

        }

        public void setAlphabet(Mock<Alphabet> mock)
        {
            convertor.setAlphabet(mock.Object);
        }

        [Test]
        public void notIntegratedTest()
        {
            var alphabet = new Mock<Alphabet>();
            //от 65 до 90 находится английский алфавит
            alphabet.Setup(x => x.getRange()).Returns(new int[] { 65, 90 });

            var converter = new NumberConverter(alphabet.Object);

            Assert.AreEqual(converter.Convert(54),"BC");
        }

        [OneTimeTearDown]
        public void Stop()
        {
            convertor = null;
        }
    }
}

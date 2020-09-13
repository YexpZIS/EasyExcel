using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SqlServer.Server;
using NUnit.Framework;

namespace EasyExcel
{
    public interface Alphabet
    {
        int[] getRange();
    }

    public class RU : Alphabet
    {
        public int[] getRange()
        {
            // Отсутствует буква Ё
            return new int[2] { 1040, 1071 };
        }
    }

    public class EN : Alphabet
    {
        public int[] getRange()
        {
            // От 65 до 90 включительно идут английские буквы в верхнем регистре
            return new int[2] { 65, 90 };
        }
    }

    [TestFixture]
    class AlphabetTest
    {
        Alphabet en;
        Alphabet ru;

        [OneTimeSetUp]
        public void Start()
        {
            en = new EN();
            ru = new RU();
        }

        [TestCase(0, "А")]
        [TestCase(1, "Я")]
        public void correctRangeRU(int index, string letter)
        {
            var RU = ru.getRange();
            rangeChecker(RU, index, letter);
        }
        [TestCase(0, "A")]
        [TestCase(1, "Z")]
        public void correctRangeEN(int index, string letter)
        {
            var EN = en.getRange();
            rangeChecker(EN, index, letter);
        }
        private void rangeChecker(int[] range,int index ,string letter)
        {
            string character = charToString(range[index]);
            Assert.AreEqual(character, letter);
        }
        private string charToString(int number)
        {
            return ((char)number).ToString();
        }

        [TestCase(2, "В")]
        [TestCase(17, "С")]
        public void lettersInIntervalRU(int index, string letter)
        {
            var RU = ru.getRange();
            intervalChecker(RU, index, letter);
        }

        [TestCase(2, "C")]
        [TestCase(14, "O")]
        public void lettersInIntervalEN(int index, string letter)
        {
            var EN = en.getRange();
            intervalChecker(EN, index, letter);
        }
        private void intervalChecker(int[] range, int index, string letter)
        {
            index += range[0];
            string character = charToString(index);
            Assert.AreEqual(character, letter);
        }
    }
}

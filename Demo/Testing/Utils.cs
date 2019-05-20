/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Testing
{
    /// <summary>
    /// Utils class for demos and tests
    /// </summary>
    public class Utils
    {

        private static Random rnd;

        /// <summary>
        /// Gets a (pseudo) random string of ASCII characters
        /// </summary>
        /// <param name="length">Length of the string</param>
        /// <returns>Randomly generated string</returns>
        public static string PseudoRandomString(int length)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < length; i++)
            {
                sb.Append((char)PseudoRandomInteger(32, 126));
            }
            return sb.ToString();
        }

        /// <summary>
        /// Gets a (pseudo) random string of ASCII characters within a minimum and maximum range
        /// </summary>
        /// <param name="minLength">Minimum length</param>
        /// <param name="maxLength">Maximum length</param>
        /// <returns>Randomly generated string</returns>
        public static string PseudoRandomString(int minLength, int maxLength)
        {
            int len = PseudoRandomInteger(minLength, maxLength);
            return PseudoRandomString(len);
        }

        /// <summary>
        /// Gets a (pseudo) random long within a minimum and maximum value
        /// </summary>
        /// <param name="minLength">Minimum value</param>
        /// <param name="maxLength">Maximum value</param>
        /// <returns>Randomly generated long</returns>
        public static long PseudoRandomLong(long min, long max)
        {
            if (rnd == null) { rnd = new Random(DateTime.Now.Millisecond); }
            byte[] buffer = new byte[8];
            rnd.NextBytes(buffer);
            long longRnd = BitConverter.ToInt64(buffer, 0);
            return (Math.Abs(longRnd % (max - min)) + min);
        }

        /// <summary>
        /// Gets a (pseudo) random integer within a minimum and maximum value
        /// </summary>
        /// <param name="minLength">Minimum value</param>
        /// <param name="maxLength">Maximum value</param>
        /// <returns>Randomly generated integer</returns>
        public static int PseudoRandomInteger(int min, int max)
        {
            if (rnd == null) { rnd = new Random(DateTime.Now.Millisecond); }
            return Utils.rnd.Next(min, max) + 1;
        }

        /// <summary>
        /// Gets a (pseudo) random double within a minimum and maximum value
        /// </summary>
        /// <param name="minLength">Minimum value</param>
        /// <param name="maxLength">Maximum value</param>
        /// <returns>Randomly generated double</returns>
        public static double PseudoRandomDouble(double min, double max)
        {
            if (rnd == null) { rnd = new Random(DateTime.Now.Millisecond); }
            return (Utils.rnd.NextDouble() * (max - min)) + min;
        }

        /// <summary>
        /// Gets a (pseudo) random bool 
        /// </summary>
        /// <returns>Randomly generated bool</returns>
        public static bool PseudoRandomBool()
        {
            int i = Utils.PseudoRandomInteger(0, 1);
            if (i == 0) { return false; }
            else { return true; }
        }

        /// <summary>
        /// Gets a (pseudo) random DateTime object within a minimum and maximum value
        /// </summary>
        /// <param name="minLength">Minimum date</param>
        /// <param name="maxLength">Maximum date</param>
        /// <returns>Randomly generated DateTime</returns>
        /// <remarks>Excel, respectively its OADate function does not support a dates before 30th December 1899. Such dates will cause an exception in NanoXLSX</remarks>
        public static DateTime PseduoRandomDate(DateTime min, DateTime max)
        {
            long ticks = PseudoRandomLong(min.Ticks, max.Ticks);
            return new DateTime(ticks);
        }


    }
}

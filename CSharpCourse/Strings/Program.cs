using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Strings
{
    class Program
    {
        static void Main(string[] args)
        {
            string city = "Ankara";
            Console.WriteLine(city[0]);
            foreach(var letter in city)
            {
                Console.WriteLine(letter);
            }

            string city2 = "İstanbul";
            string result = city + city2;
            Console.WriteLine(result);
            Console.ReadLine();
        }
    }
}

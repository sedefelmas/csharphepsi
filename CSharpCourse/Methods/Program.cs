using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Methods
{
    class Program
    {
        static void Main(string[] args)
        {
            int number1 = 20;
            int number2 = 100;
            Console.WriteLine(number1);
            Console.WriteLine(Onemli(number1, number2));
            Console.WriteLine(Onemli2(ref number1, number2));
            Console.WriteLine(number1);
            Console.WriteLine(Add(4, 5));
            Console.WriteLine(Add2(1, 2, 3, 4, 5));
            Yazı();
            Console.ReadLine();
            //ref = out ama refte sayılara değer atanmış olmak zorunda, outta böyle bir zorunluluk yok
        }
        static int Add(int number1, int number2)
        {
            var result = number1 + number2;
            return result;
        }
        static void Yazı()
        {
            Console.WriteLine("Yazı yazan method");
        }
        static int Onemli(int number1, int number2)
        {
            number1 = 30;
            return number1 + number2;
        }
        static int Onemli2(ref int number1, int number2)
        {
            number1 = 30;
            return number1 + number2;
        }
        static int Add2(params int[] numbers)
        {
            return numbers.Sum();
        }
    }
}

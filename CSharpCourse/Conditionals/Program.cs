using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Conditionals
{
    class Program
    {
        static void Main(string[] args)
        {
            var number = 10;
            //if (number == 10)
            //{
            //    Console.WriteLine("Number is 10");
            //}
            //    else
            //   {
            //    Console.WriteLine("Number is not 10");
            //}
            //Console.WriteLine(number == 10 ? "Number is 10" : "Number is not 10");
            bool isGirl = false;
            string[] animals = { "dog", "cat", "monkey", "rabbit" };
            switch (isGirl)
            {
                case true:
                    Console.WriteLine("That's a girl");
                    break;
                case false:
                    Console.WriteLine("That's not a girl");
                    break;
            }
            for (int index = 0; index <= animals.Length; index++)
            {
                switch (index)
                {
                    case 0:
                        Console.WriteLine("dog");
                        break;
                    case 1:
                        Console.WriteLine("cat");
                        break;
                    case 2:
                        Console.WriteLine("monkey");
                        break;
                    case 3:
                        Console.WriteLine("rabbit");
                        break;
                    default:
                        Console.WriteLine("The name of the last animal is not given!");
                        break;
                }
            }
            Console.ReadLine();
        }
    }
}

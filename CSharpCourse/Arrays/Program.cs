using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Arrays
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] students = new string[3];
            //string [] students = new [] {"Sedef", "Beyza", "Ayşe"};
            students[0] = "Sedef";
            students[1] = "Beyza";
            students[2] = "Ayşe";
            foreach(var student in students)
            {
                Console.WriteLine(student);
            }
            string[,] regions = new string[5, 3]
            {
                {"Ankara", "Konya", "Kırıkkale"},
                {"Rize", "Trabzon", "Samsun"},
                {"Manisa", "İzmir", "Aydın"},
                {"İstanbul", "Edirne", "Kocaeli"},
                {"Antalya", "Adana", "Mersin" }
            };
            for(int i = 0; i <= regions.GetUpperBound(0); i++)
            {
                for(int j = 0; j <= regions.GetUpperBound(1); j++)
                {
                    Console.WriteLine(regions[i, j]);
                }
            }
            Console.ReadLine();
        }
    }
}

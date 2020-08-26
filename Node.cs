using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tekstualna_steganografija
{
   public class Node
    {
        public int freq { get; set; }
        public char data { get; set; }
        public Node child0, child1;


        public Node()
        { }

        public Node(char d, int f = -1)
        {
            data = d;
            freq = f;
            child0 = null;
            child1 = null;

        }

        public Node(Node c0, Node c1)
        {

            freq = c0.freq + c1.freq;
            child0 = c0;
            child1 = c1;
        }

        public List<bool> Traverse(char symbol, List<bool> code)
        {

            if (child1 == null && child0 == null)
            {
                if (symbol.Equals(this.data))
                {
                    return code;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                List<bool> left = null;
                List<bool> right = null;

                if (child0 != null)
                {
                    List<bool> leftPath = new List<bool>();
                    leftPath.AddRange(code);
                    leftPath.Add(false);

                    left = child0.Traverse(symbol, leftPath);
                }

                if (child1 != null)
                {
                    List<bool> rightPath = new List<bool>();
                    rightPath.AddRange(code);
                    rightPath.Add(true);
                    right = child1.Traverse(symbol, rightPath);
                }

                if (left != null)
                {
                    return left;
                }
                else
                {
                    return right;
                }
            }
        }

        public void Traverse1(string code = "")
        {

            if (child0 != null)
            {
                child0.Traverse1(code + '0');
                child1.Traverse1(code + '1');
            }
            else
            {
                if (data.Equals('\0'))
                    Console.WriteLine("Karakter: null" + " Frekvenca: " + freq + " Kod: " + code);
                else
                    Console.WriteLine("Karakter: " + data + " Frekvenca: " + freq + " Kod: " + code);
                List<int> frequency = new List<int>();


            }


        }
    }
}

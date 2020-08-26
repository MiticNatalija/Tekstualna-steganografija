using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;


namespace Tekstualna_steganografija
{
    public class HuffmanTree
    {
        private List<Node> nodes = new List<Node>();
        public Node Root { get; set; }
        public Dictionary<char, int> Frequencies = new Dictionary<char, int>();
        Node currNode { get; set; }

        public void build(string source)
        {
            for (int i = 0; i < source.Length; i++)
            {
                if (!Frequencies.ContainsKey(source[i]))
                {
                    Frequencies.Add(source[i], 0);
                }

                Frequencies[source[i]]++;
            }
            buildFromFrequences();


        }
        public void buildFromFrequences()
        {
            if (Frequencies == null)
                return;
            foreach (KeyValuePair<char, int> symbol in Frequencies)
            {
                nodes.Add(new Node() { data = symbol.Key, freq = symbol.Value });
            }

            while (nodes.Count > 1)
            {
                List<Node> orderedNodes = nodes.OrderBy(node => node.freq).ToList<Node>();

                if (orderedNodes.Count >= 2)
                {

                    List<Node> taken = orderedNodes.Take(2).ToList<Node>();

                    Node parent = new Node()
                    {
                        data = '*',
                        freq = taken[0].freq + taken[1].freq,
                        child0 = taken[0],
                        child1 = taken[1]
                    };

                    nodes.Remove(taken[0]);
                    nodes.Remove(taken[1]);
                    nodes.Add(parent);
                }

                this.Root = nodes.FirstOrDefault();

            }
        }
        public List<bool> encode(string source)
        {
            List<bool> encodedSource = new List<bool>();

            for (int i = 0; i < source.Length; i++)
            {
                List<bool> encodedSymbol = this.Root.Traverse(source[i], new List<bool>());
                encodedSource.AddRange(encodedSymbol);
            }

           // BitArray bits = new BitArray(encodedSource.ToArray());

            return encodedSource;
        }

        public string decode(List<bool> bits)
        {
            Node current = this.Root;
            string decoded = "";

            foreach (bool bit in bits)
            {
                if (bit)
                {
                    if (current.child1 != null)
                    {
                        current = current.child1;
                    }
                }
                else
                {
                    if (current.child0 != null)
                    {
                        current = current.child0;
                    }
                }

                if (isLeaf(current))
                {
                    decoded += current.data;
                    if (current.data.ToString() == "\0")
                        break;
                    current = this.Root;

                }
            }

            return decoded;
        }
        public string decodeChar(bool bit)
        {
            if (currNode == null)
                currNode = this.Root;
            string toRet = string.Empty;
            if (bit)
            {
                if (currNode.child1 != null)
                {
                    currNode = currNode.child1;
                }
            }
            else
            {
                if (currNode.child0 != null)
                {
                    currNode = currNode.child0;
                }
            }
            if (isLeaf(currNode))
            {
                toRet = currNode.data.ToString();
                currNode = this.Root;
            }
            return toRet;
        }

        public bool isLeaf(Node node)
        {
            return (node.child0 == null && node.child1 == null);
        }

    }

}

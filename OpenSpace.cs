using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Windows;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Globalization;
using System.Collections;
using System.Security.Cryptography;

namespace Tekstualna_steganografija
{
    public class OpenSpace
    {


        private static List<bool> txtToBinary(string txt)
        {
            List<bool> msg = new List<bool>();
            foreach (char c in txt)
            {
                int i;
                for (i = 0; i < 8; i++)
                {
                    bool bit = ((c & (128 >> i)) != 0) ? true : false;
                    msg.Add(bit);
                }
            }
            return msg;
        }
        private static string binaryToString(List<bool> binary)
        {
            string toRet = "";
            int i = 0;
            while (i + 8 <= binary.Count)
            {

                List<bool> oneByte = binary.GetRange(i, 8);
                int value = 0;

                for (int j = oneByte.Count - 1; j >= 0; j--)
                {
                    if (oneByte[j])
                        value += Convert.ToInt16(Math.Pow(2, oneByte.Count - j - 1));
                }

                toRet += (char)value;
                i += 8;
            }

            return toRet;
        }
        public static bool stego(string txt, string fileName, string newFileName)
        {

            int noWords = WordDocument.noWords(fileName);
            if (noWords - 1 < txt.Length * 8)
            {
                return false;
            }


            List<bool> msg = txtToBinary(txt);
            string coverString = WordDocument.returnDocument(fileName);
            List<char> cover = coverString.ToList();

            string stegoText = "";
            int j = 0;
            int k = 0;


            do
            {

                char pref = cover[k++];
                if (pref == ' ')
                {
                    if (msg[j])
                    {

                        stegoText += pref + " ";
                    }
                    else
                    {
                        stegoText += pref;

                    }
                    j++;
                }
                else
                    stegoText += pref;

            } while (k < cover.Count - 1 && j < msg.Count);

            if (j < msg.Count)
            {
                if (msg[j])
                {

                    stegoText += "  ";
                }
                else
                {
                    stegoText += " ";

                }
            }
            if (k < cover.Count - 1)
            {
                stegoText += coverString.Substring(k);

            }


            WordDocument.createWord(stegoText, newFileName);

            return true;
        }
        public static string stegoRet(string fileName)
        {
            string toRet = "";
            List<bool> binary = new List<bool>();
            string stego = WordDocument.returnDocument(fileName);
            List<char> chars;
            chars = stego.ToCharArray().ToList();
            int end = 0;
            int bajt = 0;
            int i;
            for (i = 0; i < chars.Count; i++)
            {

                if (chars[i] == ' ')
                {
                    if (chars[i + 1] == ' ')
                    {
                        binary.Add(true);
                        i++;
                        end = 0;
                    }
                    else
                    {
                        binary.Add(false);
                        end++;
                        if (end > 7 && bajt == 7)
                            break;
                    }
                    bajt = (bajt == 7) ? 0 : bajt + 1;
                }

            }


            i = 0;
            while (i < binary.Count)
            {

                List<bool> oneByte = binary.GetRange(i, 8);
                int value = 0;

                for (int j = oneByte.Count - 1; j >= 0; j--)
                {
                    if (oneByte[j])
                        value += Convert.ToInt16(Math.Pow(2, oneByte.Count - j - 1));
                }

                toRet += (char)value;
                i += 8;
            }


            return toRet;
        }
        public static Dictionary<char, int> stegoHuffman(string txt, string fileName, string newFileName)
        {


            List<bool> msg = new List<bool>();
            HuffmanTree huffTree = new HuffmanTree();
            txt += '\0';
            huffTree.build(txt);
            huffTree.Root.Traverse1();

            List<bool> encoded = huffTree.encode(txt);
            foreach (bool b in encoded)
            {
                msg.Add(b);
            }

            int noWords = WordDocument.noWords(fileName);
            if (noWords - 1 < msg.Count)
            {
                return null;
            }
            string coverString = WordDocument.returnDocument(fileName);
            List<char> cover = coverString.ToCharArray().ToList();

            string stegoText = "";
            int j = 0;
            int k = 0;


            do
            {

                char pref = cover[k++];
                if (pref == ' ')
                {
                    if (msg[j])
                    {

                        stegoText += pref + " ";
                    }
                    else
                    {
                        stegoText += pref;

                    }
                    j++;
                }
                else
                    stegoText += pref;

            } while (k < cover.Count - 1 && j < msg.Count);

            if (j < msg.Count)
            {
                if (msg[j])
                {

                    stegoText += "  ";
                }
                else
                {
                    stegoText += " ";

                }
            }
            if (k < cover.Count - 1)
            {
                stegoText += coverString.Substring(k);

            }


            WordDocument.createWord(stegoText, newFileName);
            return huffTree.Frequencies;
        }

        public static string stegoHuffmanRet(string fileName, Dictionary<char, int> frequencies)
        {
            string toRet = "";
            List<bool> binary = new List<bool>();
            List<char> chars;
            chars = WordDocument.returnChars(fileName);
            HuffmanTree huffTree = new HuffmanTree();
            huffTree.Frequencies = frequencies;
            huffTree.buildFromFrequences();
            for (int i = 0; i < chars.Count; i++)
            {
                if (chars[i] == ' ')
                {
                    string temp;
                    if (chars[i + 1] == ' ')
                    {
                        binary.Add(true);
                        i++;
                        temp = huffTree.decodeChar(true);

                    }
                    else
                    {
                        binary.Add(false);
                        temp = huffTree.decodeChar(false);

                    }
                    if (temp != string.Empty && temp != "\0")
                    {
                        toRet += temp;
                    }
                    if (temp == "\0")
                        break;

                }
            }

            return toRet;

        }
        public static bool stegoJustification(string txt, string fileName, string newFileName)
        {
            List<int> linesNum;
            bool broj = WordDocument.justificationMax(fileName, txt.Length * 8, out linesNum);
            if (!broj)
                return false;
            List<bool> msg = txtToBinary(txt);
            List<string> lines = WordDocument.returnLines(fileName, msg, linesNum);

            WordDocument.createWord(string.Join(" ", lines), newFileName);


            return true;
        }
        public static string stegoJustificationRet(string fileName)
        {
            string toRet = WordDocument.returnCodedLines(fileName);


            return toRet;
        }

        public static bool stegoZWC(string txt, string fileName, string newFileName)
        {
            int noWords = WordDocument.noWords(fileName);
            if ((noWords - 1) * 2 < txt.Length * 8)
            {
                return false;
            }
            List<bool> msg = txtToBinary(txt);
            string coverString = WordDocument.returnDocument(fileName);
            List<char> cover = coverString.ToCharArray().ToList();

            string stegoText = "";
            int j = 0;
            int k = 0;

            do
            {

                char pref = cover[k++];
                if (pref == ' ')
                {
                    if (!msg[j])
                    {
                        if (!msg[j + 1])
                            stegoText += pref;
                        else
                            stegoText += pref + "\u200B";
                    }
                    else
                    {
                        if (!msg[j + 1])
                            stegoText += "\u200B" + pref;
                        else
                            stegoText += "\u200B" + pref + "\u200B";

                    }
                    j += 2;

                }
                else
                    stegoText += pref;

            } while (k < cover.Count && j < msg.Count);


            if (k < cover.Count)
            {
                stegoText += coverString.Substring(k);

            }

            WordDocument.createWord(stegoText, newFileName);

            return true;
        }
        public static string stegoZWCret(string fileName)
        {
            List<char> lista = WordDocument.returnChars(fileName);
            List<bool> binary = new List<bool>();
            int end = 0;
            int bajt = 0;
            for (int i = 0; i < lista.Count; i++)
            {
                if (lista[i] == ' ')
                {
                    if ((int)lista[i - 1] == 8203)
                    {
                        binary.Add(true);
                        end = 0;
                    }
                    else
                    {
                        binary.Add(false);
                        end++;
                    }

                    if ((int)lista[i + 1] == 8203)
                    {

                        binary.Add(true);
                        end = 0;
                    }
                    else
                    {
                        binary.Add(false);
                        end++;
                    }

                }

                if (end > 7 && bajt == 7)
                    break;
                bajt = (bajt == 7) ? 0 : bajt + 1;
            }
            string toRet = binaryToString(binary);
            return toRet;

        }
        public static bool stegoChar(string txt, string fileName, string newFileName, int[] stegoKey, List<bool> compMsg)
        {
            long noChars = WordDocument.noChars(fileName);
            if ((noChars - 1) * 4 < txt.Length * 8)
            {
                return false;
            }

            List<bool> msg;
            if (compMsg != null)
            {
                msg = compMsg;
            }
            else
                msg = txtToBinary(txt);


            string coverString = WordDocument.returnDocument(fileName);
            List<char> cover = coverString.ToCharArray().ToList();

            string stegoText = "";
            int j = 0;
            int k = 0;

            do
            {

                char pref = cover[k++];
                stegoText += pref;
                if (msg[j++])
                {
                    stegoText += Convert.ToChar(stegoKey[0]);
                }
                if (msg[j++])
                {
                    stegoText += Convert.ToChar(stegoKey[1]);
                }
                if (msg[j++])
                {
                    stegoText += Convert.ToChar(stegoKey[2]);
                }
                if (msg[j++])
                {
                    stegoText += Convert.ToChar(stegoKey[3]);
                }


            } while (k < cover.Count - 1 && j < msg.Count);


            if (k < cover.Count - 1)
            {
                stegoText += coverString.Substring(k);

            }

            WordDocument.createWord(stegoText, newFileName);

            return true;
        }
        public static string stegoCharRet(string fileName, int[] stegoKey, string[] compressKey, List<string> table)
        {
            string toRet = "";
            List<char> lista = WordDocument.returnChars(fileName);
            List<bool> binary = new List<bool>();
            int end = 0;
            int bajt = 0;
            int i = 0;
            List<int> key = stegoKey.ToList();
            while (i + 4 < lista.Count)
            {
                int group = 0;
                for (int j = 0; j < 4; j++)
                {
                    int value = (int)lista[i + 1];
                    int n = key.IndexOf(value);
                    if (n < 0)
                        break;
                    for (int k = group; k < n; k++)
                    {
                        binary.Add(false);
                        group++;
                        end++;
                        bajt++;

                    }
                    if (group < 4)
                    {
                        binary.Add(true);
                        group++;
                        i++;
                        end = 0;
                        bajt++;
                    }

                }
                while (group < 4)
                {
                    binary.Add(false);
                    group++;
                    end++;
                    bajt++;
                }
                i++;
                if (table != null)
                {
                    if (binary.Count >= table.Count * 2)
                        break;
                }
                else
                {
                    if (end > 7 && bajt == 8)
                        break;
                    if (bajt == 8)
                        bajt = 0;
                }
            }
            if (table != null)
            {
                binary = groupDecompress(binary, table, compressKey);
            }
            toRet = binaryToString(binary);

            return toRet;

        }
        public static List<bool> groupCompress(string txt, string[] stegoKey, out List<string> table)
        {
            List<bool> binary = txtToBinary(txt);
            int i = 0;
            List<bool> toRet = new List<bool>();
            table = new List<string>();
            while (i + 3 < binary.Count)
            {
                toRet.Add(binary[i + 2]);
                toRet.Add(binary[i + 3]);
                if (binary[i])
                {
                    if (binary[i + 1])
                    {
                        table.Add(stegoKey[3]);
                    }
                    else
                    {
                        table.Add(stegoKey[2]);
                    }
                }
                else
                {
                    if (binary[i + 1])
                    {
                        table.Add(stegoKey[1]);
                    }
                    else
                    {
                        table.Add(stegoKey[0]);
                    }
                }
                i += 4;
            }


            return toRet;

        }
        public static List<bool> groupDecompress(List<bool> binary, List<string> table, string[] stegoKey)
        {
            int i = 0;
            List<bool> fullText = new List<bool>();
            while (i < table.Count)
            {
                if (table[i].Equals(stegoKey[0]))
                {
                    fullText.Add(false);
                    fullText.Add(false);
                }
                else if (table[i].Equals(stegoKey[1]))
                {
                    fullText.Add(false);
                    fullText.Add(true);
                }
                else if (table[i].Equals(stegoKey[2]))
                {
                    fullText.Add(true);
                    fullText.Add(false);

                }
                else if (table[i].Equals(stegoKey[3]))
                {
                    fullText.Add(true);
                    fullText.Add(true);
                }
                fullText.Add(binary[i * 2]);
                fullText.Add(binary[i * 2 + 1]);
                i++; ;
            }

            return fullText;
        }
        public static bool stegoEncryption(string txt, string fileName, string newFileName, out byte[] iv, out byte[] key, int[] stegoKey, string[] compressKey, out List<string> table)
        {
            Aes myAes = Aes.Create();
            iv = myAes.IV;
            key = myAes.Key;

            byte[] encrypted = AESEncypt.EncryptStringToBytes_Aes(txt, myAes.Key, myAes.IV);
            string encString = Convert.ToBase64String(encrypted);

            List<bool> compMsg = groupCompress(encString, compressKey, out table);
            bool ok = stegoChar(encString, fileName, newFileName, stegoKey, compMsg);
            return ok;


        }
        public static string stegoEncryptionRet(string fileName, byte[] iv, byte[] key, int[] stegoKey, string[] compressKey, List<string> table)
        {
            Aes myAes = Aes.Create();
            myAes.IV = iv;
            myAes.Key = key;
            string dec = stegoCharRet(fileName, stegoKey, compressKey, table);
            // dec = dec.Remove(dec.Length - 1, 1);

            byte[] bytes = Convert.FromBase64String(dec);

            string decrypted = AESEncypt.DecryptStringFromBytes_Aes(bytes, key, iv);

            return decrypted;
        }


    }
}

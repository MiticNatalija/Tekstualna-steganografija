using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Security.Cryptography;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Tekstualna_steganografija
{
    class Program
    {
        static void Main(string[] args)
        {
            string tajnaPoruka = "ps";
            string newFileName;
            string newFileName1;
            string fileName1 = "ravno.docx";
            string fileName = Path.GetFullPath(fileName1);
            string directory = Directory.GetCurrentDirectory();

            WordDocument.clearSpaces(fileName);
            while (true)
            {
                Console.WriteLine("\n");
                Console.WriteLine("Tajna poruka je: " + tajnaPoruka);
                Console.WriteLine("Unesite broj za zeljeni algoritam:");
                Console.WriteLine("1 - Metoda blanko znakova kod poravnatog teksta");
                Console.WriteLine("2 - Metoda blanko znakova kod neporavnatog teksta");
                Console.WriteLine("3 - ZWC metoda");
                Console.WriteLine("4 - Metoda koriscenja nevidljivih simbola");
                Console.WriteLine("5 - Metoda sa Hafmanovom kompresijom");
                Console.WriteLine("6 - Metoda sa enkripcijom");
                Console.WriteLine("7 - Promenite tajnu poruku");
                Console.WriteLine("q - kraj");
                string alg = Console.ReadLine();

                //1. algoritam - metoda blanko znakova kod poravnatog teksta
                if (alg == "1")
                {
                    newFileName1 = "stego1.docx";
                    newFileName = Path.Combine(directory, newFileName1);
                    if (OpenSpace.stegoJustification(tajnaPoruka, fileName, newFileName))
                    {
                        Console.WriteLine("Stego fajl je kreiran.");
                        Console.WriteLine("Poruka iz stego fajla je: " + OpenSpace.stegoJustificationRet(newFileName));
                    }
                    else
                        Console.WriteLine("Predugacka poruka!");

                }
                //2. algoritam - metoda blanko znakova kod neporavnatog teksta
                else if (alg == "2")
                {
                    newFileName1 = "stego2.docx";
                    newFileName = Path.Combine(directory, newFileName1);
                    if (OpenSpace.stego(tajnaPoruka, fileName, newFileName))
                    {
                        Console.WriteLine("Stego fajl je kreiran.");
                        Console.WriteLine("Poruka iz stego fajla je: " + OpenSpace.stegoRet(newFileName));
                    }
                    else
                        Console.WriteLine("Predugacka poruka!");
                }


                //3. algoritam - ZWC metoda
                else if (alg == "3")
                {
                    newFileName1 = "stego3.docx";
                    newFileName = Path.Combine(directory, newFileName1);
                    if (OpenSpace.stegoZWC(tajnaPoruka, fileName, newFileName))
                    {
                        Console.WriteLine("Stego fajl je kreiran.");
                        Console.WriteLine("Poruka iz stego fajla je: " + OpenSpace.stegoZWCret(newFileName));
                    }
                    else
                        Console.WriteLine("Predugacka poruka!");
                }
                //4. algoritam - Metoda koriscenja nevidljivih simbola
                else if (alg == "4")
                {
                    newFileName1 = "stego4.docx";
                    newFileName = Path.Combine(directory, newFileName1);
                    int[] stegoKey = { 8207, 8206, 8205, 8204 };
                    if (OpenSpace.stegoChar(tajnaPoruka, fileName, newFileName, stegoKey, null))
                    {
                        Console.WriteLine("Stego fajl je kreiran.");
                        Console.WriteLine("Poruka iz stego fajla je: " + OpenSpace.stegoCharRet(newFileName, stegoKey, null, null));
                    }
                    else
                        Console.WriteLine("Predugacka poruka!");
                }
                //5. algoritam - Metoda sa Hafmanovom kompresijom
                else if (alg == "5")
                {
                    newFileName1 = "stego5.docx";
                    newFileName = Path.Combine(directory, newFileName1);
                    Dictionary<char, int> freq = OpenSpace.stegoHuffman(tajnaPoruka, fileName, newFileName);
                    if (freq != null)
                    {
                        Console.WriteLine("Stego fajl je kreiran.");
                        Console.WriteLine("Poruka iz stego fajla je: " + OpenSpace.stegoHuffmanRet(newFileName, freq));
                    }
                    else
                        Console.WriteLine("Predugacka poruka!");
                }
                // 6. algoritam - Metoda sa enkripcijom
                else if (alg == "6")
                {
                    newFileName1 = "stego6.docx";
                    newFileName = Path.Combine(directory, newFileName1);
                    List<string> table;
                    string[] compressKey = { "G1", "G2", "G3", "G4" };


                    byte[] iv;
                    byte[] key;
                    int[] stegoKey = { 8207, 8206, 8205, 8204 };

                    bool ok = OpenSpace.stegoEncryption(tajnaPoruka, fileName, newFileName, out iv, out key, stegoKey, compressKey, out table);
                    if (ok)
                    {
                        Console.WriteLine("Stego fajl je kreiran.");
                        Console.WriteLine("Poruka iz stego fajla je: " + OpenSpace.stegoEncryptionRet(newFileName, iv, key, stegoKey, compressKey, table));
                    }
                    else
                        Console.WriteLine("Predugacka poruka!");
                }
                else if (alg == "7")
                {
                    Console.WriteLine("Unesite poruku: ");
                    tajnaPoruka = Console.ReadLine();
                }
                else if (alg == "q")
                {
                    break;
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Windows;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Drawing;
using System.Text.RegularExpressions;



namespace Tekstualna_steganografija
{
    public class WordDocument
    {

        public static void clearSpaces(string fileName)
        {
            _Application oWord;
            object oMissing = Type.Missing;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(fileName);

            Range rng = doc.Content;
            rng.Select();
            string cover = rng.Text;

            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex("[ ]{2,}", options);
            cover = regex.Replace(cover, " ");

            oWord.Selection.TypeText(cover);
            oWord.ActiveDocument.Save();
            oWord.Quit();
        }

        public static void createWord(string txt, string name)
        {
            Application oWord = new Application();
            oWord.Visible = false;
            object missing = System.Reflection.Missing.Value;
            Document document = oWord.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            document.Content.SetRange(0, 0);
            document.Content.Text = txt;
            document.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            object filename = name;
            document.SaveAs2(ref filename);
            document.Close(ref missing, ref missing, ref missing);
            document = null;
            oWord.Quit(ref missing, ref missing, ref missing);


        }
        public static int noWords(string name)
        {

            _Application oWord;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(name);


            //selektovanje celog dokumenta
            Range rng = doc.Content;
            rng.Select();
            int noWords = rng.ComputeStatistics(WdStatistic.wdStatisticWords);
            doc.Close();
            doc = null;
            oWord.Quit();
            return noWords;
        }
        public static long noChars(string name)
        {
            _Application oWord;
            object oMissing = Type.Missing;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(name);
            Range rng = doc.Content;
            rng.Select();
            long noWords = rng.ComputeStatistics(WdStatistic.wdStatisticCharactersWithSpaces);
            doc.Close();
            doc = null;
            oWord.Quit();
            return noWords;
        }
        public static string returnDocument(string name)
        {
            Application oWord;
            object oMissing = Type.Missing;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(name);

            Selection currSelection = oWord.Selection;



            Range rng = doc.Content;
            rng.Select();
            string cover = rng.Text;

            doc.Close();
            doc = null;
            oWord.Quit();

            return cover;
        }
        public static bool justificationMax(string fileName, int numBits, out List<int> lines)
        {

            lines = new List<int>();
            Application oWord;
            object oMissing = Type.Missing;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(fileName);

            string strLine;
            bool bolEOF = false;
            int noBits = 0;
            doc.Characters[1].Select();
            do
            {
                object unit = WdUnits.wdLine;
                object count = 1;
                oWord.Selection.MoveEnd(ref unit, ref count);

                strLine = oWord.Selection.Text;
                Range rgn = oWord.Selection.Range;
                float d = (rgn.ComputeStatistics(WdStatistic.wdStatisticWords) - 1) / 2;
                int maxSpaces = (int)Math.Floor(d);

                Microsoft.Office.Interop.Word.Font font = rgn.Font;
                string fontName = font.Name;
                float fontSize = font.Size;


                PageSetup p = oWord.Selection.PageSetup;

                float margR = p.RightMargin; //72
                float margL = p.LeftMargin; //72
                float pageWidth = p.PageWidth; //595.3 - 72*2 ukupno za tekst 451.3
                int u = doc.Application.Width;
                System.Drawing.Font f1 = new System.Drawing.Font(fontName, fontSize, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                Image fakeImage = new Bitmap(1, 1);
                Graphics graphics = Graphics.FromImage(fakeImage);
                graphics.PageUnit = GraphicsUnit.Point;
                SizeF size = graphics.MeasureString(strLine, f1, 0, StringFormat.GenericTypographic);
                SizeF sizeSpace = graphics.MeasureString(" ", f1);



                float maxSpacesPom = (pageWidth - margL - margR - size.Width)
                    / sizeSpace.Width;
                int maxSpaces1 = (int)Math.Floor(maxSpacesPom);

                maxSpaces = (maxSpaces < maxSpaces1) ? maxSpaces : maxSpaces1;  //konacan broj spaceova
                if (maxSpaces > 0)
                    noBits += maxSpaces;
                lines.Add(maxSpaces);

                object direction = WdCollapseDirection.wdCollapseEnd;
                oWord.Selection.Collapse(ref direction);

                if (oWord.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                    bolEOF = true;
            } while (!bolEOF);


            doc.Close();
            doc = null;
            oWord.Quit();
            if (noBits >= numBits)
                return true;
            return false;
        }
        public static List<string> returnLines(string fileName, List<bool> msg, List<int> numLines)
        {

            List<string> lines = new List<string>();
            Application oWord;
            object oMissing = Type.Missing;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(fileName);

            string strLine;
            bool bolEOF = false;
            int msgCounter = 0; //za tajnu poruku
            doc.Characters[1].Select();
            int i = 0;
            do
            {
                object unit = WdUnits.wdLine;
                object count = 1;
                oWord.Selection.MoveEnd(ref unit, ref count);

                strLine = oWord.Selection.Text;
                Range rgn = oWord.Selection.Range;

                int spaceCounter = 0;

                if (numLines[i] > 0 && msgCounter < msg.Count)
                {
                    // List<string> pom = strLine.Split(' ').ToList();

                    List<char> pom1 = strLine.ToList();
                    int lastBit = -1;
                    int j = 0;
                    while (j < pom1.Count && spaceCounter < numLines[i] && msgCounter < msg.Count)
                    {
                        if (pom1[j] == ' ')
                        {
                            if (lastBit == -1)
                            {
                                if (msg[msgCounter])
                                {
                                    pom1.Insert(j, ' ');
                                    lastBit = 0;
                                    j++;
                                }
                                else
                                {
                                    lastBit = 1;
                                }
                            }
                            else if (lastBit == 1)
                            {
                                pom1.Insert(j, ' ');
                                lastBit = -1;
                                spaceCounter++;
                                j++;
                                msgCounter++;
                            }
                            else
                            {
                                lastBit = -1;
                                spaceCounter++;
                                msgCounter++;
                            }
                        }
                        j++;
                    }


                    string stegoLine = new string(pom1.ToArray());
                    lines.Add(stegoLine);
                }
                else
                    lines.Add(strLine);

                object direction = WdCollapseDirection.wdCollapseEnd;
                oWord.Selection.Collapse(ref direction);

                if (oWord.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                    bolEOF = true;
                i++;
            } while (!bolEOF);


            doc.Close();
            doc = null;
            oWord.Quit();
            return lines;
        }
        public static string returnCodedLines(string fileName)
        {
            string toRet = "";
            int i;
            Application oWord;
            object oMissing = Type.Missing;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(fileName);

            string strLine;
            bool bolEOF = false;

            doc.Characters[1].Select();
            List<bool> codedMsg = new List<bool>();
            do
            {
                object unit = WdUnits.wdLine;
                object count = 1;
                oWord.Selection.MoveEnd(ref unit, ref count);

                strLine = oWord.Selection.Text;
                Range rgn = oWord.Selection.Range;
                object direction = WdCollapseDirection.wdCollapseEnd;
                oWord.Selection.Collapse(ref direction);
                List<char> pom = strLine.ToList();

                int last = -1;
                for (i = 0; i < pom.Count - 1; i++)
                {
                    if (pom[i] == ' ')
                    {
                        if (pom[i + 1] == ' ')
                        {
                            if (last == -1)
                            {
                                last = 1;
                            }
                            else
                            {
                                if (last == 0)
                                {
                                    codedMsg.Add(false);
                                    last = -1;
                                }
                                else break;
                            }
                            i++;
                        }

                        else
                        {

                            if (last == -1)
                            {
                                last = 0;
                            }
                            else
                            {
                                if (last == 1)
                                {
                                    codedMsg.Add(true);
                                    last = -1;
                                }
                                else break;
                            }
                        }
                    }
                }


                if (oWord.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                    bolEOF = true;
            } while (!bolEOF);
            i = 0;
            while (i + 8 <= codedMsg.Count)
            {

                List<bool> oneByte = codedMsg.GetRange(i, 8);
                int value = 0;

                for (int j = oneByte.Count - 1; j >= 0; j--)
                {
                    if (oneByte[j])
                        value += Convert.ToInt16(Math.Pow(2, oneByte.Count - j - 1));
                }

                toRet += (char)value;
                i += 8;
            }

            doc.Close();
            doc = null;
            oWord.Quit();
            return toRet;

        }
        public static List<char> returnChars(string fileName)
        {
            List<char> chars = new List<char>();
            Application oWord;
            object oMissing = Type.Missing;
            oWord = new Application();
            oWord.Visible = false;
            Document doc = oWord.Documents.Open(fileName);
            Selection currSelection = oWord.Selection;
            Range rng = doc.Content;
            rng.Select();

            chars = rng.Text.ToList();



            doc.Close();
            doc = null;
            oWord.Quit();
            return chars;
        }



    }
}

using System;
using System.Collections.Generic;
using System.IO;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace ConvertTxtToPpt
{
    public partial class Converter
    {
        public Converter()
        {

        }

        public static bool createPpt(string txtPath, string pptPath, string tickPath)
        {
            PPT.Application pptApp = new PPT.ApplicationClass();
            pptApp.Visible = MsoTriState.msoTrue;
            PPT.Presentations pptDoc = pptApp.Presentations;
            PPT.Presentation pres = pptDoc.Add(MsoTriState.msoFalse);
            PPT.Slides slides = pres.Slides;

            List<Puzzle> list = generateList(txtPath);
            int count = 1;
            PPT.Slide first = slides.Add(1, PPT.PpSlideLayout.ppLayoutBlank);
            int row = list.Count;
            Console.WriteLine(row);
            row = (row % 5 == 0) ? row / 5 : row / 5 + 1;
            Console.WriteLine(row);
            PPT.Table table = first.Shapes.AddTable(row, 5).Table;
            row = 1;
            int column = 1;
            foreach (Puzzle a in list)
            {
                PPT.Shape text = table.Cell(row, column).Shape;
                column++;
                if (column > table.Columns.Count)
                {
                    column = 1;
                    row++;
                }
                text.TextFrame.TextRange.Text = "" + count;
                count++;
                text.TextFrame.TextRange.ActionSettings[PPT.PpMouseActivation.ppMouseClick].Action = PPT.PpActionType.ppActionHyperlink;
                text.TextFrame.TextRange.ActionSettings[PPT.PpMouseActivation.ppMouseClick].Hyperlink.Address = "";
                text.TextFrame.TextRange.ActionSettings[PPT.PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = "" + count;
                PPT.Slide slide = slides.Add(count, PPT.PpSlideLayout.ppLayoutTitle);
                slide.Shapes[1].TextFrame.TextRange.Text = (count - 1) + ". " + a.question;
                slide.Shapes[2].TextFrame.TextRange.Text = a.answer;
                if (!a.isChoose)
                {
                    slide.Shapes[2].AnimationSettings.EntryEffect = PPT.PpEntryEffect.ppEffectRandom;
                }
                else
                {
                    Console.WriteLine(tickPath);
                    slide.Shapes.AddPicture(tickPath, MsoTriState.msoFalse, MsoTriState.msoTrue, 300, 300, 50, 50);
                    slide.Shapes[3].AnimationSettings.EntryEffect = PPT.PpEntryEffect.ppEffectRandom;
                }
            }
            pres.SaveAs(pptPath, PPT.PpSaveAsFileType.ppSaveAsPresentation);
            pres.Close();
            pptApp.Quit();
            return true;
        }

        private static List<Puzzle> generateList(String filePath)
        {
            StreamReader sr = new StreamReader(filePath);
            List<Puzzle> questions = new List<Puzzle>();
            while (!sr.EndOfStream)
            {
                Puzzle puzzle = new Puzzle();
                String[] line = sr.ReadLine().Split('#');
                puzzle.question = line[0];
                puzzle.answer = line[1];
                if (line.GetLength(0) == 3) puzzle.isChoose = true;
                questions.Add(puzzle);
            }
            return questions;
        }
    }
}

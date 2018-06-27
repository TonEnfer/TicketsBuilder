using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using TemplateEngine.Docx;
using System.IO;
using System.Diagnostics;

namespace ExamQuestionGenerator
{
    class Program
    {
        static string inputPath = @"E:\exam\input.txt";
        static string templatePath = @"E:\exam\ЭБ_ГЭ_2018.04.docx";
        static string endFilePath = @"E:\exam\tickets.docx";
        public struct Ticket
        {
            public int number;
            public string quest1;
            public string quest2;

            public Ticket(int n, string s1, string s2)
            {
                number = n;
                quest1 = s1;
                quest2 = s2;
            }
        }
        static void Main(string[] args)
        {
            List<Ticket> tickets = GetTickets(inputPath);
            if (tickets.Count % 2 != 0)
                tickets.Add(new Ticket(0, "УДАЛИ МЕНЯ", "УДАЛИ МЕНЯ"));
            for (int i = 0; i < tickets.Count; i += 2)
            {
                var valToFill = new Content(
                    new FieldContent("TicketNumber1", Convert.ToString(tickets[i].number)),
                    new FieldContent("Question11", tickets[i].quest1),
                    new FieldContent("Question12", tickets[i].quest2),
                    new FieldContent("TicketNumber2", Convert.ToString(tickets[i + 1].number)),
                    new FieldContent("Question21", tickets[i + 1].quest1),
                    new FieldContent("Question22", tickets[i + 1].quest2)
                    );
                var tmpFilePath = Path.GetDirectoryName(templatePath) + @"/tmp"+i+@".docx";
                if (File.Exists(tmpFilePath))
                    File.Delete(tmpFilePath);
                File.Copy(templatePath, tmpFilePath);
                using (var outputDocument = new TemplateProcessor(tmpFilePath)
                .SetRemoveContentControls(true))
                {
                    outputDocument.FillContent(valToFill);
                    outputDocument.SaveChanges();
                }
                //AddTicketToNewFile(tmpFilePath, endFilePath);
                //File.Delete(tmpFilePath);
            }
            Console.ReadKey();
        }

        static List<Ticket> GetTickets(string path)
        {
            List<Ticket> tickets = null;
            string[] text;

            try
            {
                text = File.ReadAllLines(path);
                if (text.Length != 0)
                    tickets = new List<Ticket>();
                else
                    return null;
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
                return null;
            }

            int i = 1;
            foreach (var s in text)
            {
                string[] tmp = s.Split('\t');
                tickets.Add(new Ticket(i, tmp[0], tmp[1]));
                i++;
            }

            return tickets;
        }
        static void AddTicketToNewFile(string pathToOldFile, string pathToNewFile)
        {
            WordprocessingDocument oldDoc = WordprocessingDocument.Open(pathToOldFile, false);
            Body oldFileContent = oldDoc.MainDocumentPart.Document.Body;

            WordprocessingDocument newDoc = WordprocessingDocument.Open(pathToNewFile, true);
            Body newFileContent = newDoc.MainDocumentPart.Document.Body;



            foreach (var q in oldFileContent.Elements())
            {
                var e = q.CloneNode(true);
                newFileContent.Append(e);
            }
            oldDoc.Close();
            newDoc.Close();
        }
    }
}

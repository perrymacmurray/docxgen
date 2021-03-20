using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace DocxGen
{
    class Program
    {
        static string DOCUMENT_NAME;
        static SimpleReplacement EMPTY = new SimpleReplacement("");

        static void Main(string[] args)
        {
            if (args.Length == 0)
                DOCUMENT_NAME = "TEMPLATE.docx";
            else
                DOCUMENT_NAME = args[0];

            Dictionary<string, Replacement> replacements = new Dictionary<string, Replacement>();
            Dictionary<string, Replacement> blockReplacements = new Dictionary<string, Replacement>();

            //Special replacements
            replacements.Add("{DATE}", new DateTimeReplacement());
            replacements.Add("{TIME}", new DateTimeReplacement("h:mm tt"));
            replacements.Add("{DATETIME}", new DateTimeReplacement("MM/dd/yyyy h:mm tt"));

            //User-specified replacements + editing
            Application word = null;
            Document doc = null;
            try
            {
                word = new Application();
                Console.WriteLine($"Opening document \"{DOCUMENT_NAME}\"...");
                if (!File.Exists(DOCUMENT_NAME))
                {
                    Console.WriteLine("ERROR: Template does not exist at the specified location.");
                    Console.ReadKey();
                    return;
                }

                doc = word.Documents.Open(Path.GetFullPath(DOCUMENT_NAME));

                word.Visible = false;
                doc.Activate();

                string text = doc.Range().Text;

                string blockText = "";
                for (int i = 0; i < text.Length; i++)
                {
                    if (blockText != "")
                        blockText += text[i];

                    if (text[i] != '{')
                        continue;

                    i++;
                    if (text[i] == '*' && text[i + 1] == '}')
                    {
                        i++;
                        if (blockText != "")
                        {
                            Console.WriteLine($"Malformed block: {blockText}");
                            Console.WriteLine("Application will exit.");
                            Console.ReadKey();
                            return;
                        }

                        blockText = "{*}";
                    }
                    else if (text[i] == '/')
                    {
                        i++;
                        blockText += "/}"; //First character already parsed

                        if (!Regex.IsMatch(blockText, "{\\*[^}]"))
                        {
                            Console.WriteLine($"Malformed block: {blockText}");
                            Console.WriteLine("Application will exit.");
                            Console.ReadKey();
                            return;
                        }

                        var fields = Regex.Matches(blockText, "{\\*[^}][^}]*}");
                        bool delete = true;
                        foreach (Match m in fields)
                        {
                            if (replacements[m.Value].ToString() != "")
                                delete = false;
                        }

                        if (delete)
                            blockReplacements.Add(blockText, EMPTY);
                        blockText = "";
                    }
                    else //Replacement 
                    {
                        string field = "";

                        while (text[i] != '}')
                        {
                            field += text[i];
                            i++;
                        }

                        string replace = field + '}';
                        field = '{' + field + '}';

                        if (!replacements.ContainsKey(field))
                        {
                            Console.Write($"Text for field {field}: ");
                            string answer = Console.ReadLine();

                            replacements.Add(field, new SimpleReplacement(answer));
                        }
                        if (blockText != "")
                            blockText += replace;
                    }
                }

                foreach (string key in blockReplacements.Keys)
                {
                    word.Selection.Find.Execute(FindText: key, ReplaceWith: blockReplacements[key].ToString(), Replace: WdReplace.wdReplaceAll);
                }
                foreach (string key in replacements.Keys)
                {
                    word.Selection.Find.Execute(FindText: key, ReplaceWith: replacements[key].ToString(), Replace: WdReplace.wdReplaceAll);
                }
                word.Selection.Find.Execute(FindText: "{*}", ReplaceWith: "", Replace: WdReplace.wdReplaceAll); //Replace unused block indicators
                word.Selection.Find.Execute(FindText: "{/}", ReplaceWith: "", Replace: WdReplace.wdReplaceAll);

                while (doc.Range().Text.Contains("  "))
                {
                    word.Selection.Find.Execute(FindText: "  ", ReplaceWith: " ", Replace: WdReplace.wdReplaceAll);
                }

                doc.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), $"Generated {DateTime.Now.ToString("yyMMdd.hh-mm-ss")}.docx"));
            } finally
            {
                doc?.Close();
                word?.Quit();
            }
        }
    }
}

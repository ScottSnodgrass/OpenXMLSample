// --------------------------------------------------------------------------------------
// <copyright file="program.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;

namespace OpenXmlDemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string input;

            do
            {
                DisplayMenu();
                input = Console.ReadLine();
                ProcessInput(input);
            } while (input != null && !input.Equals("7", StringComparison.InvariantCultureIgnoreCase));
        }


        private static void DisplayMenu()
        {
            Console.Clear();
            Console.WriteLine("Open XML SDK Demo");
            Console.WriteLine("=================");
            Console.WriteLine("");
            Console.WriteLine("Please select one of the following from:");
            Console.WriteLine("[1]: read author info from file");
            Console.WriteLine("[2]: Create New Document");
            Console.WriteLine("[3]: read Plain Text from Document");
            Console.WriteLine("[4]: generate Participant list template 1");
            Console.WriteLine("[5]: generate Participant list template 2");
            Console.WriteLine("[6]: creating a certificate");
            Console.WriteLine("[7]: exit program");
        }

        private static void ProcessInput(string input)
        {
            const string fileName = @".\Hallo OpenXML.docx";

            switch (input)
            {
                case "1":
                    DocumentInfoSample.DisplayAuthor(fileName);
                    break;
                case "2":
                    DocumentCreationSample.CreateDocument();
                    break;
                case "3":
                    Console.WriteLine(DocumentReaderSample.ReadTextFromDocument(fileName));
                    break;
                case "4":
                    AttendeeListSample1.CreateAttendeeList();
                    break;
                case "5":
                    AttendeeListSample2.CreateAttendeeList2();
                    break;
                case "6":
                    MailMergeSample.MailMerge();
                    break;
                case "7":
                    break;
                default:
                    Console.WriteLine("Bitte eine gültige Auswahl wählen");
                    break;
            }
            Console.WriteLine("Bitte Eingabetaste drücken");
            Console.ReadLine();
        }

   }
}
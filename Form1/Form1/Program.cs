//Gets user input and appends text and image saved to drive

using System;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

class WriteTextFile
{
    static void Main()
    {
        // This example assumes a "/Users/michaelhodge/Desktop/" path on your machine.
        // Modify the path if necessary

        // Gets user input
        string testString, testString2;
        string newLine = Environment.NewLine;
        Console.Write("Enter a title: ");
        testString = Console.ReadLine();
        Console.WriteLine("You entered '{0}'", testString);

        Console.Write("Enter a description: ");
        testString2 = Console.ReadLine();
        Console.WriteLine("You entered '{0}'", testString2);

        // Create Document
        Document document = new Document();
        Section s = document.AddSection();
        Paragraph p = s.AddParagraph();

        // Insert Image and Set Its Size
        p.AppendText(testString + newLine + testString2);
        DocPicture Pic = p.AppendPicture(Image.FromFile(@"/Users/michaelhodge/Desktop/photo.jpeg"));
        Pic.Width = 500;
        Pic.Height = 500;

        // Save and Launch  
        document.SaveToFile(@"/Users/michaelhodge/Desktop/MyPic.Docx", FileFormat.Docx);
        System.Diagnostics.Process.Start(@"/Users/michaelhodge/Desktop/MyPic.Docx");
    }
}

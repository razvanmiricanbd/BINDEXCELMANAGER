using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace OpenXmlNbdWrapper
{
    public class OpenXmlNbdReader
    {

        public static List<NbdHeadline> GetHeadlinesAsListFromDoc(String FilePath)
        {
            //Stream stream = File.Open(FilePath, FileMode.Open);
            List<NbdHeadline> headerParagrahList = new List<NbdHeadline>();
            // Open a WordProcessingDocument based on a stream.
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(FilePath, true))
            { 

                // Assign a reference to the existing document body.
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            Document wDocument = wordprocessingDocument.MainDocumentPart.Document;
            // Add new text.
            List<Paragraph> wParagraphList = wDocument.Descendants<Paragraph>().ToList();
            //String content = null;
            NbdHeadline currentHeadline = null;
            foreach (Paragraph p in wParagraphList)
            {
                if (
                    p.ParagraphProperties != null &&
                     p.ParagraphProperties.ParagraphStyleId != null &&
                    p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Heading"))
                {
                    currentHeadline = new NbdHeadline
                    {
                        StyleName = p.ParagraphProperties.ParagraphStyleId.Val.Value,
                        Text = p.InnerText
                    };
                    headerParagrahList.Add(currentHeadline);
                }
                else
                {
                        if (currentHeadline != null)
                        {
                            currentHeadline.Content += p.InnerText;
                        }

                }
            }
            }
            // Close the document handle.
            //wordprocessingDocument.Close();
            //stream.Close();
            return headerParagrahList;
           

        }

        public static NbdHeadline GetHeadlinesAsTreeFromDoc(String FilePath)
        {
            //Stream stream = File.Open(FilePath, FileMode.Open);
            NbdHeadline root = new NbdHeadline
            {
                Text = "",
                StyleName = "none"
            };
            // Open a WordProcessingDocument based on a stream.
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(FilePath, true))
            {

                // Assign a reference to the existing document body.
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                Document wDocument = wordprocessingDocument.MainDocumentPart.Document;
                List<ImagePart> imgPart = wordprocessingDocument.MainDocumentPart.ImageParts.ToList();
                // Add new text.
                List<Paragraph> wParagraphList = wDocument.Descendants<Paragraph>().ToList();
                NbdHeadline parentLeaf = null;
                NbdHeadline currentLeaf = null;
                foreach (Paragraph p in wParagraphList)
                {
                    if (
                        p.ParagraphProperties != null &&
                        p.ParagraphProperties.ParagraphStyleId != null &&
                        p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Heading"))

                    {
                        int level = getLevel(p.ParagraphProperties.ParagraphStyleId.Val.Value);

                        currentLeaf = new NbdHeadline
                        {
                            StyleName = p.ParagraphProperties.ParagraphStyleId.Val.Value,
                            Text = p.InnerText
                        };
                        if (parentLeaf == null)
                        {

                            parentLeaf = root;
                        }
                        if (parentLeaf.Level < level)
                        {
                            parentLeaf.AddChildren(currentLeaf);
                            parentLeaf = currentLeaf;
                        }
                        else
                        {
                            parentLeaf = GetParent(level, parentLeaf);
                            parentLeaf.AddChildren(currentLeaf);
                            parentLeaf = currentLeaf;
                        }
                    }
                    else
                    {
                        if (currentLeaf != null)
                            currentLeaf.Content += p.InnerText;
                        else
                            root.Content += p.InnerText;

                    }
                }
            }
            // Close the document handle.
            //wordprocessingDocument.Close();
            //stream.Close();
            return root;

        }
        private static int getLevel(String style) {
            return int.Parse(style.Substring(7, 1));
        }
        private static NbdHeadline GetParent(int level,NbdHeadline headline) {
            if (headline.Parent != null
               && headline.Parent.Level < level)
            {
                return headline.Parent;
            }
            else
            if (headline.Parent == null)
                return headline;
            else return GetParent(level, headline.Parent);

        }

       
    }
}

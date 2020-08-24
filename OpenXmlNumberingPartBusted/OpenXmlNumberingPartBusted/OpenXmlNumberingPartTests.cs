using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlNumberingPartBusted
{
    [TestClass]
    [TestCategory("Integration")]
    public class OpenXmlNumberingPartTests
    {
        private static IEnumerable<object[]> Factories => new[]
        {
            new object[] { new NumberingFactoryExpectedToWork() }, //will pass unit tests but produces documents without lists
            new object[] { new NumberingFactoryWorkaround() } //will produce documents with numbered lists but will not pass tests
        };

        [TestMethod]
        [DynamicData(nameof(Factories))]
        public void CreateADocument_WithNumberingParts_PartsComeOutNumbered(INumberingFactory numberingFactory)
        {
            var fileLocation = $"./TestResults/{Guid.NewGuid()}.docx"; // Using a real file instead of a memory stream so that the actual output can be inspected
            using var document = WordprocessingDocument.Create(fileLocation, WordprocessingDocumentType.Document);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document();
            mainPart.Document.AppendChild(new Body());

            var numbering1 = numberingFactory.CreateNewNumberingSequence(document);
            var paragraph = document.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.AppendChild(numbering1.CloneNode(true));
            paragraph.AppendChild(new Run(new Text("1")));
            paragraph = document.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.AppendChild(numbering1.CloneNode(true));
            paragraph.AppendChild(new Run(new Text("2")));

            var numbering2 = numberingFactory.CreateNewNumberingSequence(document);
            paragraph = document.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.AppendChild(numbering2.CloneNode(true));
            paragraph.AppendChild(new Run(new Text("one")));
            paragraph = document.MainDocumentPart.Document.Body.AppendChild(new Paragraph());
            properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.AppendChild(numbering2.CloneNode(true));
            paragraph.AppendChild(new Run(new Text("two")));

            document.Save();
            Console.WriteLine(fileLocation);
            document.Close();
            // Even when saved to disk and restored one method passes unit tests but does not produce numbers on the list and one produces numbers on the list but fails the test.
            using var saveDocument = WordprocessingDocument.Open(fileLocation, false);
            Assert.AreEqual(2, saveDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().Count());
            Assert.AreEqual(2, saveDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().Count());
        }

        public interface INumberingFactory
        {
            NumberingProperties CreateNewNumberingSequence(WordprocessingDocument document);
        }

        private class NumberingFactoryExpectedToWork : INumberingFactory
        {
            public NumberingProperties CreateNewNumberingSequence(WordprocessingDocument document)
            {
                var definitions = document.MainDocumentPart.NumberingDefinitionsPart ??
                                  document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                var numbering = definitions.Numbering ??= new Numbering(); // The property should be initialized since it is the root of the part.  Not that big a deal though.

                var baseNumbering = numbering.AppendChild
                (
                    new AbstractNum
                    (
                        new Level
                        {
                            NumberingFormat = new NumberingFormat
                            {
                                Val = NumberFormatValues.Decimal
                            },
                            LevelIndex = 0
                        }
                    )
                    {
                        AbstractNumberId = numbering
                            .Descendants<AbstractNum>()
                            .Select<AbstractNum, int>(an => an.AbstractNumberId)
                            .DefaultIfEmpty(0)
                            .Max() + 1,
                        MultiLevelType = new MultiLevelType {Val = MultiLevelValues.SingleLevel},
                    }
                );
                var numberingInstance = numbering.AppendChild
                (
                    new NumberingInstance
                    {
                        NumberID = numbering
                            .Descendants<NumberingInstance>()
                            .Select<NumberingInstance, int>(an => an.NumberID)
                            .DefaultIfEmpty(0)
                            .Max() + 1,
                        AbstractNumId = new AbstractNumId {Val = baseNumbering.AbstractNumberId}
                    }
                );

                // It doesn't make a difference if it gets saved or not.
                // If the instance of Numbering is associated with the Numbering property of the part the resulting document will not be numbered.
                //numbering.Save(definitions);

                return new NumberingProperties
                {
                    NumberingId = new NumberingId { Val = numberingInstance.NumberID },
                    NumberingLevelReference = new NumberingLevelReference { Val = 0 }
                };
            }
        }

        private class NumberingFactoryWorkaround : INumberingFactory
        {
            //save the numbering instance away from the part so that it can be repeatedly saved with more elements
            private readonly Numbering _numbering = new Numbering();

            public NumberingProperties CreateNewNumberingSequence(WordprocessingDocument document)
            {
                // The numbering part API does not work correctly.
                // Any Numbering instance that becomes associated with the part as its property does not produce numbering in the final document

                var definitions = document.MainDocumentPart.NumberingDefinitionsPart ??
                                  document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                definitions.Numbering ??= new Numbering(); // The property must be accessed but cannot be used.  It is null even on a new instance where this is the root part.
                // Using Numbering.Load produces the same result.

                var baseNumbering = new AbstractNum
                (
                    new Level
                    {
                        NumberingFormat = new NumberingFormat
                        {
                            Val = NumberFormatValues.Decimal
                        },
                        LevelIndex = 0
                    }
                )
                {
                    AbstractNumberId = _numbering
                        .Descendants<AbstractNum>()
                        .Select<AbstractNum, int>(an => an.AbstractNumberId)
                        .DefaultIfEmpty(0)
                        .Max() + 1,
                    MultiLevelType = new MultiLevelType {Val = MultiLevelValues.SingleLevel},
                };
                var numberingInstance = new NumberingInstance
                {
                    NumberID = _numbering
                        .Descendants<NumberingInstance>()
                        .Select<NumberingInstance, int>(an => an.NumberID)
                        .DefaultIfEmpty(0)
                        .Max() + 1,
                    AbstractNumId = new AbstractNumId {Val = baseNumbering.AbstractNumberId}
                };

                _numbering.AppendChild(baseNumbering);
                _numbering.AppendChild(numberingInstance);
                _numbering.Save(definitions);
                // Assigning a new Numbering to the property has the same result as using one accessed from the property, even if it was saved first.

                return new NumberingProperties
                {
                    NumberingId = new NumberingId {Val = numberingInstance.NumberID},
                    NumberingLevelReference = new NumberingLevelReference {Val = 0}
                };
            }
        }
    }
}

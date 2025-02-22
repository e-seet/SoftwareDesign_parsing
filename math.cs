using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Math; // ✅ Math Namespace
using DocumentFormat.OpenXml.Wordprocessing; // ✅ Wordprocessing Namespace

namespace Utilities
{
    public static class MathExtractor
    {
        // ✅ Extract math from paragraphs (detects Unicode math symbols & OfficeMath)
        public static List<Dictionary<string, object>> ExtractParagraphsWithMath(
            DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph
        )
        {
            var paragraphs = new List<Dictionary<string, object>>();

            // ✅ Extract normal text from the paragraph
            string text = string.Join(
                "",
                paragraph
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Select(t => t.Text)
            );

            // ✅ Extract math equations inside the paragraph
            var mathElements = paragraph
                .Descendants<DocumentFormat.OpenXml.Math.OfficeMath>()
                .ToList();
            var mathContentList = new List<string>();

            foreach (var mathElement in mathElements)
            {
                // ✅ Extract readable math content (not just MathML)
                string mathText = ExtractReadableMath(mathElement);
                // Console.WriteLine(mathElement.OuterXml); // Debugging

                mathContentList.Add(mathText);
            }

            //added this new line
            // ✅ Extract breaks (page breaks, line breaks)
            // var breaks = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Break>().ToList();
            // if (breaks.Any())
            // {
            // 	foreach (var br in breaks)
            // 	{
            // 		string breakType = br.Type?.Value switch
            // 		{
            // 			BreakValues.Page => "page_break",
            // 			BreakValues.TextWrapping => "line_break",
            // 			_ => "unknown_break"
            // 		};

            // 		paragraphs.Add(new Dictionary<string, object>
            // {
            // 	{ "type", breakType },
            // 	{ "content", breakType == "page_break" ? "[PAGE BREAK]" : "[LINE BREAK]" }
            // });
            // 	}
            // }

            // ✅ Store extracted text (if any)
            if (!string.IsNullOrWhiteSpace(text))
            {
                paragraphs.Add(
                    new Dictionary<string, object> { { "type", "paragraph" }, { "content", text } }
                );
            }

            // ✅ Store extracted math (if any)
            foreach (var mathText in mathContentList)
            {
                paragraphs.Add(
                    new Dictionary<string, object> { { "type", "math" }, { "content", mathText } }
                );
            }

            return paragraphs;
        }

        // ✅ Extract Readable Math Expression from OfficeMath
        public static string ExtractReadableMath(DocumentFormat.OpenXml.Math.OfficeMath mathElement)
        {
            var mathParts = new List<string>();

            // ✅ Extract Fractions (m:f)
            foreach (
                var fraction in mathElement.Descendants<DocumentFormat.OpenXml.Math.Fraction>()
            )
            {
                var numerator =
                    fraction
                        .Numerator?.Descendants<DocumentFormat.OpenXml.Math.Text>()
                        .Select(t => t.Text)
                        .FirstOrDefault() ?? "?";

                var denominator =
                    fraction
                        .Denominator?.Descendants<DocumentFormat.OpenXml.Math.Text>()
                        .Select(t => t.Text)
                        .FirstOrDefault() ?? "?";
                mathParts.Add($"({numerator}/{denominator})"); // ✅ Convert to (numerator/denominator)
            }

            // ✅ Extract Square Roots (m:rad)
            foreach (var radical in mathElement.Descendants<DocumentFormat.OpenXml.Math.Radical>())
            {
                var baseElement = radical
                    .Elements<DocumentFormat.OpenXml.Math.Base>()
                    .FirstOrDefault();
                var rootContent =
                    baseElement
                        ?.Descendants<DocumentFormat.OpenXml.Math.Text>()
                        .Select(t => t.Text)
                        .FirstOrDefault() ?? "?"; // ✅ Use Math.Text instead of Wordprocessing.Text
                mathParts.Add($"√({rootContent})"); // ✅ Convert to √(x)
            }

            // ✅ Extract Normal Math Text (e.g., Multiplication ×)
            foreach (var run in mathElement.Descendants<DocumentFormat.OpenXml.Math.Run>())
            {
                string mathText = string.Join(
                    "",
                    run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                        .Select(t => t.Text)
                );
                if (!string.IsNullOrWhiteSpace(mathText))
                {
                    mathParts.Add(mathText);
                }
            }

            return mathParts.Any() ? string.Join(" ", mathParts) : "Math content not detected.";
        }

        // ✅ Extract MathML and readable content from OfficeMath equations
        // public static Dictionary<string, object> ExtractMathEquation(DocumentFormat.OpenXml.Math.OfficeMath mathElement)
        // {

        // 	if (mathElement == null)
        // 	{
        // 		Console.WriteLine("No math element found.");
        // 		return new Dictionary<string, object>
        // 		{
        // 			{ "type", "math" },
        // 			{ "content", "No math content detected" },
        // 			{ "mathML", "" }
        // 		};
        // 	}

        // 	// ✅ Extract raw MathML XML representation
        // 	string mathML = mathElement.OuterXml;
        // 	Console.WriteLine("Extracted MathML: " + mathML); // Debugging

        // 	// ✅ Extract readable math content
        // 	string mathText = ExtractReadableMath(mathElement);
        // 	Console.WriteLine("Extracted Math Text: " + mathText); // Debugging

        // 	return new Dictionary<string, object>
        // 	{
        // 		{ "type", "math" },
        // 		{ "content", mathText },
        // 		{ "mathML", mathML }
        // 	};
        // }

        // ✅ Extract Math Expressions for JSON Output
        private static string ExtractMathContent(DocumentFormat.OpenXml.Math.OfficeMath mathElement)
        {
            var mathParts = mathElement
                .Descendants<DocumentFormat.OpenXml.Math.Run>()
                .Select(r =>
                    string.Join(
                        "",
                        r.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                            .Select(t => t.Text)
                    )
                )
                .ToList();

            if (!mathParts.Any())
            {
                Console.WriteLine("No Math text found.");
                return "Math content not detected.";
            }

            return string.Join(" ", mathParts);
        }
    }
}

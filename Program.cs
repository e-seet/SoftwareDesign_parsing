using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Utilities;

class Program
{
	static void Main()
	{
		string filePath = "Datarepository.docx"; // Change this to your actual file path
		string jsonOutputPath = "output.json"; // File where JSON will be saved

		using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
		{
			var documentData = new
			{
				// metadata = GetDocumentMetadata(wordDoc),
				// headers = ExtractHeaders(wordDoc),
				// !!footer still exists issues
				// footers = ExtractFooters(wordDoc),
				document = ExtractDocumentContents(wordDoc),
			};

			// Convert to JSON format with UTF-8 encoding fix (preserves emojis, math, and Chinese)
			string jsonOutput = JsonSerializer.Serialize(
				documentData,
				new JsonSerializerOptions
				{
					WriteIndented = true,
					Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
				}
			);

			// Write JSON to file
			File.WriteAllText(jsonOutputPath, jsonOutput);
			Console.WriteLine($"JSON output saved to {jsonOutputPath}");
		}
	}

	static List<string> ExtractHeaders(WordprocessingDocument doc)
	{
		var headers = new List<string>();

		// ✅ Check if MainDocumentPart is null
		if (doc.MainDocumentPart == null)
		{
			Console.WriteLine("Error: MainDocumentPart is null.");
			return headers;
		}

		// ✅ Check if HeaderParts exist
		if (!doc.MainDocumentPart.HeaderParts.Any())
		{
			Console.WriteLine("No headers found in the document.");
			return headers;
		}

		foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
		{
			var header = headerPart.Header;

			if (header != null)
			{
				foreach (
					var paragraph in header.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
				)
				{
					// ✅ Extract normal text from headers
					string text = string.Join(
						"",
						paragraph
							.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
							.Select(t => t.Text)
					);

					if (!string.IsNullOrWhiteSpace(text))
					{
						headers.Add(text);
					}
				}
			}
		}

		return headers;
	}

	static Dictionary<string, string> GetDocumentMetadata(WordprocessingDocument doc)
	{
		var metadata = new Dictionary<string, string>();

		if (doc.PackageProperties.Title != null)
			metadata["Title"] = doc.PackageProperties.Title;
		if (doc.PackageProperties.Creator != null)
			metadata["Author"] = doc.PackageProperties.Creator;

		return metadata;
	}

	// Extract thge document content
	static List<object> ExtractDocumentContents(WordprocessingDocument doc)
	{
		var elements = new List<object>();
		var body = doc.MainDocumentPart?.Document?.Body;

		if (body == null)
		{
			Console.WriteLine("Error: Document body is null.");
			return elements;
		}

		foreach (var element in body.Elements<OpenXmlElement>())
		{
			if (element is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
			{
				Console.WriteLine("Extract Paragraph");
				elements.Add(ExtractParagraph(paragraph));
			}
			else if (element is Table table)
			{
				Console.WriteLine("IDK case 3");
				elements.Add(ExtractTable(table)); // ✅ Keep table extraction as-is
			}
		}

		return elements;
	}

	static Dictionary<string, object> ExtractParagraph(
		DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph
	)
	{

		string text = string.Join(
			"",
			paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)
		);
		string style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
		bool isBold = paragraph.Descendants<Bold>().Any();
		bool isItalic = paragraph.Descendants<Italic>().Any();
		var alignment = paragraph.ParagraphProperties?.Justification?.Val?.ToString() ?? "left";
		var paragraphData = new Dictionary<string, object>();
		var havemath = false;
		// List<Dictionary<string, object>> mathContent = null;
		List<Dictionary<string, object>> mathContent = new List<Dictionary<string, object>>();

		// ✅ Check if paragraph is completely empty
		if (string.IsNullOrWhiteSpace(text) && !paragraph.Elements<Break>().Any())
		{
			paragraphData["type"] = "empty_paragraph";
			paragraphData["content"] = "";
			return paragraphData;
		}

		// ✅ Detect Page Breaks
		if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page))
		{
			paragraphData["type"] = "page_break";
			paragraphData["content"] = "[PAGE BREAK]";
			return paragraphData;
		}

		// ✅ Detect Line Breaks
		if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.TextWrapping))
		{
			paragraphData["type"] = "line_break";
			paragraphData["content"] = "[LINE BREAK]";
			return paragraphData;
		}

		if (paragraph.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any())
		{
			mathContent = MathExtractor.ExtractParagraphsWithMath(paragraph);

			havemath = true;
			// var mathContent = MathExtractor.ExtractParagraphsWithMath(paragraph);
			// elements.AddRange(MathExtractor.ExtractParagraphsWithMath(paragraph)); // ✅ Extract paragraphs & Unicode math
			// return mathContent;
		}


		// Check for page/line breaks at the paragraph level
		if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page))
		{
			Console.WriteLine("break value\n");
			return new Dictionary<string, object>
				{
					{ "type", "page_break" },
					{ "content", "[PAGE BREAK]" }
				};
		}

		if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.TextWrapping))
		{
			Console.WriteLine("line break\n");
			return new Dictionary<string, object>
				{
					{ "type", "line_break" },
					{ "content", "[LINE BREAK]" }
				};
		}

		// Collect each run's text and formatting
		var runsList = new List<Dictionary<string, object>>();

		foreach (var run in paragraph.Elements<Run>())
		{
			string runText = string.Join("", run.Descendants<Text>().Select(t => t.Text));
			if (string.IsNullOrWhiteSpace(runText))
			{
				Console.WriteLine("Continue\n");
				continue; // Skip empty runs
			}

			bool runBold = (run.RunProperties?.Bold != null);
			bool runItalic = (run.RunProperties?.Italic != null);

			runsList.Add(new Dictionary<string, object>
				{
					{ "text", runText },
					{ "bold", runBold },
					{ "italic", runItalic }
				});
		}

		if (!runsList.Any())
		{
			Console.WriteLine("No run or breaks. Maybe empty paragraph\n");
			return new Dictionary<string, object>
				{
					{ "type", "empty_paragraph" },
					{ "content", "" }
				};
		}
		else
		{
			Console.WriteLine("last test case of ExtractParagraph function\n");

			if (havemath == true)
			{
				var mathstring = "";
				Console.WriteLine("Getting back the result and we see what is inside the for loop\n");

				foreach (var mathEntry in mathContent)
				{
					Console.WriteLine(mathEntry["content"]);
					mathstring = mathEntry["content"] + mathstring;
				}

				return new Dictionary<string, object>
				{
					{ "type", GetParagraphType(style) },
					{ "content", mathstring },
					{ "bold", isBold },
					{ "italic", isItalic },
					{ "alignment", alignment },
				};
			}
			else
			{

				return new Dictionary<string, object>
				{
					{ "type", GetParagraphType(style) },
					{ "content", text },
					{ "bold", isBold },
					{ "italic", isItalic },
					{ "alignment", alignment },
				};
			}
		}

	}


	// /*
	static Dictionary<string, object> ExtractParagraph_v2(Paragraph paragraph)
	{
		// 1) Check for page/line breaks at the paragraph level
		if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page))
		{
			Console.WriteLine("break value\n");
			return new Dictionary<string, object>
				{
					{ "type", "page_break" },
					{ "content", "[PAGE BREAK]" }
				};
		}

		if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.TextWrapping))
		{
			Console.WriteLine("line break\n");
			return new Dictionary<string, object>
				{
					{ "type", "line_break" },
					{ "content", "[LINE BREAK]" }
				};
		}

		// 2) Collect each run's text and formatting
		var runsList = new List<Dictionary<string, object>>();

		foreach (var run in paragraph.Elements<Run>())
		{
			string runText = string.Join("", run.Descendants<Text>().Select(t => t.Text));
			if (string.IsNullOrWhiteSpace(runText))
			{
				Console.WriteLine("Continue\n");
				continue; // Skip empty runs
			}

			bool runBold = (run.RunProperties?.Bold != null);
			bool runItalic = (run.RunProperties?.Italic != null);

			runsList.Add(new Dictionary<string, object>
				{
					{ "text", runText },
					{ "bold", runBold },
					{ "italic", runItalic }
				});
		}

		// 3) If no runs and no breaks, it's likely an empty paragraph
		if (!runsList.Any())
		{
			Console.WriteLine("No run or breaks. Maybe empty paragraph\n");
			return new Dictionary<string, object>
				{
					{ "type", "empty_paragraph" },
					{ "content", "" }
				};
		}

		// 4) (Optional) If paragraph contains math, handle or store it
		if (paragraph.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any())
		{
			Console.WriteLine("Got office math\n");
			// Example: If you want to extract math separately, do something like:
			// var mathSegments = MathExtractor.ExtractParagraphsWithMath(paragraph);
			// Then merge them into your runsList or store them as a separate property.

			// original code taken
			var mathContent = MathExtractor.ExtractParagraphsWithMath(paragraph);

		}

		// 5) Determine the paragraph style (Heading1, Heading2, etc.)
		string style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
		string paragraphType = style switch
		{
			"Heading1" => "h1",
			"Heading2" => "h2",
			"Heading3" => "h3",
			_ => "paragraph",
		};

		// 6) Return a dictionary representing this paragraph
		return new Dictionary<string, object>
			{
				{ "type", paragraphType },
				{ "runs", runsList }
			};
	}
	// */
	static string GetParagraphType(string style)
	{
		return style switch
		{
			"Heading1" => "h1",
			"Heading2" => "h2",
			"Heading3" => "h3",
			_ => "paragraph",
		};
	}

	static Dictionary<string, object> ExtractTable(Table table)
	{
		var tableData = new List<List<string>>();

		foreach (var row in table.Elements<TableRow>())
		{
			var rowData = row.Elements<TableCell>()
				.Select(cell =>
					string.Join(
						"",
						cell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
							.Select(t => t.Text)
					)
				)
				.ToList(); // ✅ Fixed ambiguous reference
			tableData.Add(rowData);
		}

		return new Dictionary<string, object> { { "type", "table" }, { "content", tableData } };
	}

	/* Footer below. But need to fix the page number not being picked up*/
	static List<string> ExtractFooters(WordprocessingDocument doc)
	{
		// var footers = new List<string>();

		// foreach (var footerPart in doc.MainDocumentPart.FooterParts)
		// {
		// 	var footer = footerPart.Footer;

		// 	if (footer != null)
		// 	{
		// 		foreach (var paragraph in footer.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
		// 		{
		// 			string text = string.Join("", paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
		// 			footers.Add(text);

		// 		}
		// 	}
		// }


		var footers = new List<string>();

		// ✅ Check if MainDocumentPart is null
		if (doc.MainDocumentPart == null)
		{
			Console.WriteLine("Error: MainDocumentPart is null.");
			return footers;
		}

		// ✅ Check if FooterParts exist
		if (!doc.MainDocumentPart.FooterParts.Any())
		{
			Console.WriteLine("No footers found in the document.");
			return footers;
		}

		foreach (var footerPart in doc.MainDocumentPart.FooterParts)
		{
			var footer = footerPart.Footer;

			if (footer != null)
			{
				foreach (
					var paragraph in footer.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
				)
				{
					// ✅ Extract normal text
					string text = string.Join(
						"",
						paragraph
							.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
							.Select(t => t.Text)
					);

					// ✅ Extract FieldCode elements (e.g., { PAGE })
					var fieldCodes = paragraph
						.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldCode>()
						.Select(fc => fc.Text);

					// ✅ Extract SimpleField elements (for dynamic fields like page numbers)
					var simpleFields = paragraph
						.Descendants<DocumentFormat.OpenXml.Wordprocessing.SimpleField>()
						.SelectMany(sf =>
							sf.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
						)
						.Select(t => t.Text);

					// ✅ Combine extracted text
					string combinedText =
						$"{text} {string.Join(" ", fieldCodes)} {string.Join(" ", simpleFields)}".Trim();

					if (!string.IsNullOrWhiteSpace(combinedText))
					{
						footers.Add(combinedText);
					}
				}
			}
		}

		return footers;
	}

	// static List<string> ExtractFooters(WordprocessingDocument doc)
	// {
	// 	var footers = new List<string>();

	// 	foreach (var footerPart in doc.MainDocumentPart.FooterParts)
	// 	{
	// 		var footer = footerPart.Footer;

	// 		if (footer != null)
	// 		{
	// 			foreach (var paragraph in footer.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
	// 			{
	// 				// ✅ Extract normal text from the footer
	// 				string text = string.Join("", paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));

	// 				// ✅ Extract FieldCode elements (e.g., { PAGE } placeholders)
	// 				var fieldCodes = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldCode>()
	// 					.Select(fc => fc.Text);

	// 				// ✅ Extract SimpleField elements (for dynamic content like page numbers)
	// 				var simpleFields = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.SimpleField>()
	// 					.SelectMany(sf => sf.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
	// 					.Select(t => t.Text);

	// 				// ✅ Combine all extracted content
	// 				string combinedText = $"{text} {string.Join(" ", fieldCodes)} {string.Join(" ", simpleFields)}".Trim();

	// 				if (!string.IsNullOrWhiteSpace(combinedText))
	// 				{
	// 					footers.Add(combinedText);
	// 				}
	// 			}
	// 		}
	// 	}

	// 	return footers;
	// }
}

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

using System.Drawing;

class Program
{
	static void Main()
	{
		string filePath = "Datarepository_zx.docx"; // Change this to your actual file path
													// string filePath = "Datarepository.docx"; // Change this to your actual file path
													// string filePath = "Mathrepository.docx"; // Change this to your actual file path
		string jsonOutputPath = "output.json"; // File where JSON will be saved

		using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
		{
			var documentData = new
			{
				metadata = GetDocumentMetadata(wordDoc),

				// headers = DocumentHeadersFooters.ExtractHeaders(wordDoc),

				// !!footer still exists issues. Commented for now
				// footers = DocumentHeadersFooters.ExtractFooters(wordDoc),
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
	//
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
			// Check for a Drawing element inside the run
			var drawing = element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();
			if (drawing != null)
			{
				// Console.WriteLine("Extract Image");
				// Extract images from the drawing
				var imageObjects = ExtractContent.ExtractImagesFromDrawing(doc, drawing);
				elements.AddRange(imageObjects);
			}
			else if (element is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
			{
				// Console.WriteLine("Extract Paragraph");
				elements.Add(ExtractContent.ExtractParagraph(paragraph, doc));
			}
			else if (element is Table table)
			{
				Console.WriteLine("IDK case 3");
				elements.Add(ExtractContent.ExtractTable(table)); // ✅ Keep table extraction as-is
			}
		}
		return elements;
	}


}

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

namespace Utilities
{

	public static class ExtractContent
	{
		public static Dictionary<string, object> ExtractParagraph
		(
			DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
			WordprocessingDocument doc
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

			// ✅ Extract Font Type & Font Size from Paragraph Style
			string fontType = "Default Font";
			string? fontSizeRaw = null;

			string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
			var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;

			// ✅ Check if StyleDefinitionsPart exists
			if (stylesPart != null && stylesPart.Styles != null)
			{
				var paragraphStyle = stylesPart.Styles.Elements<Style>()
					.FirstOrDefault(s => s.StyleId == styleId);

				if (paragraphStyle != null)
				{
					fontType = paragraphStyle.StyleRunProperties?.RunFonts?.Ascii?.Value ?? "Default Font";
					fontSizeRaw = paragraphStyle.StyleRunProperties?.FontSize?.Val?.Value;
				}
			}

			// ✅ Convert font size from half-points
			int fontSize = fontSizeRaw != null ? int.Parse(fontSizeRaw) / 2 : 12; // Default 12pt
			var paragraphData = new Dictionary<string, object>();

			paragraphData["alignment"] = alignment;
			// paragraphData["fontType"] = fontType;
			// paragraphData["fontSize"] = fontSize;

			var havemath = false;
			// List<Dictionary<string, object>> mathContent = null;
			List<Dictionary<string, object>> mathContent = new List<Dictionary<string, object>>();

			// ✅ Extract Paragraph-Level Font & Size Correctly
			string paraFontType = FormatExtractor.GetParagraphFont(paragraph);
			int paraFontSize = FormatExtractor.GetParagraphFontSize(paragraph);

			var PropertiesList = new List<object>
			{
				new Dictionary<string, object>
				{
					{ "bold", isBold },
					{ "italic", isItalic } ,
					{ "alignment", alignment },
					{"fontsize", fontSize},
					{"fonttype", paraFontType},
				},
			};
			Console.WriteLine(paraFontSize);
			Console.WriteLine(paraFontType);

			// ✅ Check if paragraph is completely empty
			if (string.IsNullOrWhiteSpace(text) && !paragraph.Elements<Break>().Any())
			{
				paragraphData["type"] = "empty_paragraph1";
				paragraphData["content"] = "";
				paragraphData["styling"] = PropertiesList;

				// paragraphData["alignment"] = alignment;
				// paragraphData["fonttype"] = paraFontType;
				// paragraphData["fontsize"] = paraFontSize;
				return paragraphData;
			}

			// ✅ Detect Page Breaks
			if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page))
			{
				paragraphData["type"] = "page_break";
				paragraphData["content"] = "[PAGE BREAK]";
				// paragraphData["fonttype"] = paraFontType;
				// paragraphData["fontsize"] = paraFontSize;
				paragraphData["styling"] = PropertiesList;
				return paragraphData;
			}

			// ✅ Detect Line Breaks
			if (paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.TextWrapping))
			{
				paragraphData["type"] = "line_break";
				paragraphData["content"] = "[LINE BREAK]";
				// paragraphData["fonttype"] = paraFontType;
				// paragraphData["fontsize"] = paraFontSize;
				paragraphData["styling"] = PropertiesList;

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
					{ "content", "[PAGE BREAK]" },
					{ "fonttype", paraFontType },
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

				// 	// ✅ Extract Font Type
				string runfontType = run.RunProperties?.RunFonts?.Ascii?.Value ?? "Default Font";

				// 	// ✅ Extract Font Size (stored in half-points, so divide by 2)
				// string fontSizeRaw = run.RunProperties?.FontSize?.Val?.Value ?? null; // else null
				string? runFontSizeRaw = run.RunProperties?.FontSize?.Val?.Value;
				int runFontSize = runFontSizeRaw != null ? int.Parse(runFontSizeRaw) / 2 : 12; // Default to 12pt

				runsList.Add(new Dictionary<string, object>
			{
					{ "text", runText },
					{ "styling", PropertiesList}
			});
			}

			if (!runsList.Any())
			{
				Console.WriteLine("No run or breaks. Maybe empty paragraph. \n");
				Console.WriteLine("This may be creating issues. to check\n");
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
					{ "type", FormatExtractor.GetParagraphType(style) },
					{ "content", mathstring },
					{ "styling", PropertiesList}
				};
				}
				else
				{
					return new Dictionary<string, object>
				{
					{ "type", FormatExtractor.GetParagraphType(style) },
					{ "content", text },
					{ "styling", PropertiesList}

				};
				}
			}
		}



		public static Dictionary<string, object> ExtractTable
		(
			Table table
		)
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

		public static List<Dictionary<string, object>> ExtractImagesFromDrawing
		(
			WordprocessingDocument doc,
			DocumentFormat.OpenXml.Wordprocessing.Drawing drawing)
		{
			var imageList = new List<Dictionary<string, object>>();

			// 1. Ensure MainDocumentPart is not null
			var mainPart = doc.MainDocumentPart;
			if (mainPart == null)
			{
				Console.WriteLine("Error: MainDocumentPart is null.");
				return imageList;
			}

			// 2. Find the Blip element
			var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
			if (blip == null)
			{
				Console.WriteLine("No Blip found in Drawing.");
				return imageList;
			}

			// 3. Get the relationship ID (embed)
			string? embed = blip.Embed?.Value;
			if (string.IsNullOrEmpty(embed))
			{
				Console.WriteLine("Embed is null or empty.");
				return imageList;
			}

			// 4. Retrieve the ImagePart using the relationship ID
			var part = mainPart.GetPartById(embed);
			if (part == null)
			{
				Console.WriteLine($"No part found for embed ID: {embed}");
				return imageList;
			}

			// 5. Cast part to ImagePart
			if (part is not ImagePart imagePart)
			{
				Console.WriteLine("Part is not an ImagePart.");
				return imageList;
			}

			// 6. Save the image locally
			string fileName = $"Image_{embed}.png";
			using (var stream = imagePart.GetStream())
			using (var fileStream = new FileStream(fileName, FileMode.Create))
			{
				stream.CopyTo(fileStream);
			}

			// 7. Add image info to the result list
			imageList.Add(new Dictionary<string, object>
			{
				{ "type", "image" },
				{ "filename", fileName }
			});

			return imageList;
		}

	}
}

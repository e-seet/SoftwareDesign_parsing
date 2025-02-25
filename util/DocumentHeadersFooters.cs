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
	public static class DocumentHeadersFooters
	{

		public static List<string> ExtractHeaders(WordprocessingDocument doc)
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



		/* Footer below. But need to fix the page number not being picked up*/
		public static List<string> ExtractFooters(WordprocessingDocument doc)
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
					foreach
					(
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
	}
}
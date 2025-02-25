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
	public static class FormatExtractor
	{
		public static string GetParagraphFont(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
		{
			if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.ParagraphStyleId != null)
			{
				Console.WriteLine("paragraph font type" + paragraph.ParagraphProperties.ParagraphStyleId.Val?.Value + "\n");
				return paragraph.ParagraphProperties.ParagraphStyleId.Val?.Value ?? "Default Font";
			}
			Console.WriteLine(paragraph.ParagraphProperties);
			return "Default Font";
		}

		public static int GetParagraphFontSize(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
		{
			string? fontSizeRaw = paragraph.ParagraphProperties?
				.ParagraphMarkRunProperties?
				.Elements<FontSize>()
				.FirstOrDefault()?.Val?.Value;

			return fontSizeRaw != null ? int.Parse(fontSizeRaw) / 2 : 12; // Default 12pt
		}

		public static string GetParagraphType(string style)
		{
			return style switch
			{
				"Heading1" => "h1",
				"Heading2" => "h2",
				"Heading3" => "h3",
				_ => "paragraph",
			};
		}
	}
}

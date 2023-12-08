/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Ox = DocumentFormat.OpenXml;
using OxP = DocumentFormat.OpenXml.Packaging;
using OxW = DocumentFormat.OpenXml.Wordprocessing;
using OxD = DocumentFormat.OpenXml.Drawing;

namespace HtmlToOpenXml
{
	partial class HtmlConverter
	{
		//____________________________________________________________________
		//
		// Processing known tags

		#region ProcessAcronym

		private void ProcessAcronym(HtmlEnumerator en)
		{
			// Transform the inline acronym/abbreviation to a reference to a foot note.

			string title = en.Attributes["title"];
			if (title == null) return;

			AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);

			if (elements.Count > 0 && elements[0] is OxW.Run)
			{
				string runStyle;
				OxW.FootnoteEndnoteReferenceType reference;

				if (this.AcronymPosition == AcronymPosition.PageEnd)
				{
					reference = new OxW.FootnoteReference() { Id = AddFootnoteReference(title) };
					runStyle = htmlStyles.DefaultStyles.FootnoteReferenceStyle;
				}
				else
				{
					reference = new OxW.EndnoteReference() { Id = AddEndnoteReference(title) };
					runStyle = htmlStyles.DefaultStyles.EndnoteReferenceStyle;
				}

				OxW.Run run;
				elements.Add(
					run = new OxW.Run(
						new OxW.RunProperties {
							RunStyle = new OxW.RunStyle() { Val = htmlStyles.GetStyle(runStyle, OxW.StyleValues.Character) }
						},
						reference));
			}
		}

		#endregion

		#region ProcessBlockQuote

		private void ProcessBlockQuote(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			string tagName = en.CurrentTag;
			string cite = en.Attributes["cite"];

			htmlStyles.Paragraph.BeginTag(en.CurrentTag, new OxW.ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.IntenseQuoteStyle) });

			AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);

			if (cite != null)
			{
				string runStyle;
				OxW.FootnoteEndnoteReferenceType reference;

				if (this.AcronymPosition == AcronymPosition.PageEnd)
				{
					reference = new OxW.FootnoteReference() { Id = AddFootnoteReference(cite) };
					runStyle = htmlStyles.DefaultStyles.FootnoteReferenceStyle;
				}
				else
				{
					reference = new OxW.EndnoteReference() { Id = AddEndnoteReference(cite) };
					runStyle = htmlStyles.DefaultStyles.EndnoteReferenceStyle;
				}

				OxW.Run run;
				elements.Add(
					run = new OxW.Run(
						new OxW.RunProperties {
							RunStyle = new OxW.RunStyle() { Val = htmlStyles.GetStyle(runStyle, OxW.StyleValues.Character) }
						},
						reference));
			}

			CompleteCurrentParagraph(true);
			htmlStyles.Paragraph.EndTag(tagName);
		}

		#endregion

		#region ProcessBody

		private void ProcessBody(HtmlEnumerator en)
		{
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			// Unsupported W3C attribute but claimed by users. Specified at <body> level, the page
			// orientation is applied on the whole document
			string attr = en.StyleAttributes["page-orientation"];
			if (attr != null)
			{
				OxW.PageOrientationValues orientation = Converter.ToPageOrientation(attr);

				OxW.SectionProperties sectionProperties = mainPart.Document.Body.GetFirstChild<OxW.SectionProperties>();
				if (sectionProperties == null || sectionProperties.GetFirstChild<OxW.PageSize>() == null)
                {
                    mainPart.Document.Body.Append(HtmlConverter.ChangePageOrientation(orientation));
                }
                else
                {
					OxW.PageSize pageSize = sectionProperties.GetFirstChild<OxW.PageSize>();
                    if (!pageSize.Compare(orientation))
                    {
						OxW.SectionProperties validSectionProp = ChangePageOrientation(orientation);
                        if (pageSize != null) pageSize.Remove();
						sectionProperties.PrependChild(validSectionProp.GetFirstChild<OxW.PageSize>().CloneNode(true));
                    }
                }
            }
		}

		#endregion

		#region ProcessBr

		private void ProcessBr(HtmlEnumerator en)
		{
			elements.Add(new OxW.Run(new OxW.Break()));
		}

		#endregion

		#region ProcessCite

		private void ProcessCite(HtmlEnumerator en)
		{
			ProcessHtmlElement<OxW.RunStyle>(en, new OxW.RunStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.QuoteStyle, OxW.StyleValues.Character) });
		}

		#endregion

		#region ProcessDefinitionList

		private void ProcessDefinitionList(HtmlEnumerator en)
		{
			ProcessParagraph(en);
			currentParagraph.InsertInProperties(prop => prop.SpacingBetweenLines = new OxW.SpacingBetweenLines() { After = "0" });
		}

		#endregion

		#region ProcessDefinitionListItem

		private void ProcessDefinitionListItem(HtmlEnumerator en)
		{
			AlternateProcessHtmlChunks(en, "</dd>");

			currentParagraph = htmlStyles.Paragraph.NewParagraph();
			currentParagraph.Append(elements);
			currentParagraph.InsertInProperties(prop => {
				prop.Indentation = new OxW.Indentation() { FirstLine = "708" };
				prop.SpacingBetweenLines = new OxW.SpacingBetweenLines() { After = "0" };
			});

			// Restore the original elements list
			AddParagraph(currentParagraph);
			this.elements.Clear();
		}

		#endregion

		#region ProcessDiv

		private void ProcessDiv(HtmlEnumerator en)
		{
			// The way the browser consider <div> is like a simple Break. But in case of any attributes that targets
			// the paragraph, we don't want to apply the style on the old paragraph but on a new one.
			if (en.Attributes.Count == 0 || (en.StyleAttributes["text-align"] == null && en.Attributes["align"] == null && en.StyleAttributes.GetAsBorder("border").IsEmpty))
			{
				List<Ox.OpenXmlElement> runStyleAttributes = new List<Ox.OpenXmlElement>();
				bool newParagraph = ProcessContainerAttributes(en, runStyleAttributes);
				CompleteCurrentParagraph(newParagraph);

				if (runStyleAttributes.Count > 0)
					htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes);

				// Any changes that requires a new paragraph?
				if (newParagraph)
				{
					// Insert before the break, complete this paragraph and start a new one
					this.paragraphs.Insert(this.paragraphs.Count - 1, currentParagraph);
					AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
					CompleteCurrentParagraph();
				}
			}
			else
			{
				// treat div as a paragraph
				ProcessParagraph(en);
			}
		}

		#endregion

		#region ProcessFont

		private void ProcessFont(HtmlEnumerator en)
		{
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			string attrValue = en.Attributes["size"];
			if (attrValue != null)
			{
				Unit fontSize = Converter.ToFontSize(attrValue);
                if (fontSize.IsFixed)
					styleAttributes.Add(new OxW.FontSize { Val = (fontSize.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) });
			}

			attrValue = en.Attributes["face"];
			if (attrValue != null)
			{
				// Set HightAnsi. Bug fixed by xjpmauricio on github.com/onizet/html2openxml/discussions/285439
				// where characters with accents were always using fallback font
				styleAttributes.Add(new OxW.RunFonts { Ascii = attrValue, HighAnsi = attrValue });
			}

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
		}

		#endregion

		#region ProcessHeading

		private void ProcessHeading(HtmlEnumerator en)
		{
			char level = en.Current[2];

			// support also style attributes for heading (in case of css override)
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			AlternateProcessHtmlChunks(en, "</h" + level + ">");

			OxW.Paragraph p = new OxW.Paragraph(elements);
			p.InsertInProperties(prop =>
				prop.ParagraphStyleId = new OxW.ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.HeadingStyle + level, OxW.StyleValues.Paragraph) });

			// Check if the line starts with a number format (1., 1.1., 1.1.1.)
			// If it does, make sure we make the heading a numbered item
			Ox.OpenXmlElement firstElement = elements.FirstOrDefault();
			Match regexMatch = Regex.Match(firstElement.InnerText ?? string.Empty, @"(?m)^(\d+\.)*\s");

			// Make sure we only grab the heading if it starts with a number
			if (regexMatch.Groups.Count > 1 && regexMatch.Groups[1].Captures.Count > 0)
			{
				int indentLevel = regexMatch.Groups[1].Captures.Count;

				// Strip numbers from text
				firstElement.InnerXml = firstElement.InnerXml.Replace(firstElement.InnerText, firstElement.InnerText.Substring(indentLevel * 2 + 1)); // number, dot and whitespace

				htmlStyles.NumberingList.ApplyNumberingToHeadingParagraph(p, indentLevel);
			}

			htmlStyles.Paragraph.ApplyTags(p);
			htmlStyles.Paragraph.EndTag("<h" + level + ">");
			
			this.elements.Clear();
			AddParagraph(p);
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessHorizontalLine

		private void ProcessHorizontalLine(HtmlEnumerator en)
		{
			// Insert an horizontal line as it stands in many emails.
            CompleteCurrentParagraph(true);

			// If the previous paragraph contains a bottom border or is a Table, we add some spacing between the <hr>
			// and the previous element or Word will display only the last border.
			// (see Remarks: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.bottomborder%28office.14%29.aspx)
            if (paragraphs.Count >= 2)
            {
				Ox.OpenXmlCompositeElement previousElement = paragraphs[paragraphs.Count - 2];
                bool addSpacing = false;
				OxW.ParagraphProperties prop = previousElement.GetFirstChild<OxW.ParagraphProperties>();
                if (prop != null)
                {
                    if (prop.ParagraphBorders != null && prop.ParagraphBorders.BottomBorder != null
                        && prop.ParagraphBorders.BottomBorder.Size > 0U)
                            addSpacing = true;
                }
                else
                {
					if (previousElement is OxW.Table)
                        addSpacing = true;
                }

                if (addSpacing)
                {
					currentParagraph.InsertInProperties(p => p.SpacingBetweenLines = new OxW.SpacingBetweenLines() { Before = "240" });
                }
            }

			// if this paragraph has no children, it will be deleted in RemoveEmptyParagraphs()
			// in order to kept the <hr>, we force an empty run
			currentParagraph.Append(new OxW.Run());

			// Get style from border (only top) or use Default style 
			OxW.TopBorder hrBorderStyle = null;
						
			var border = en.StyleAttributes.GetAsBorder("border");
			if (!border.IsEmpty && border.Top.IsValid)
				hrBorderStyle = new OxW.TopBorder { Val = border.Top.Style, Color = Ox.StringValue.FromString(border.Top.Color.ToHexString()), Size = (uint)border.Top.Width.Value };			
			else
				hrBorderStyle = new OxW.TopBorder() { Val = OxW.BorderValues.Single, Size = 4U };

			currentParagraph.InsertInProperties(prop => 
			prop.ParagraphBorders = new OxW.ParagraphBorders {
				TopBorder = hrBorderStyle
			});
		}

		#endregion

		#region ProcessHtml

		private void ProcessHtml(HtmlEnumerator en)
		{
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
		}

		#endregion

		#region ProcessHtmlElement

		private void ProcessHtmlElement<T>(HtmlEnumerator en) where T : Ox.OpenXmlLeafElement, new()
		{
			ProcessHtmlElement<T>(en, new T());
		}

		/// <summary>
		/// Generic handler for processing style on any Html element.
		/// </summary>
		private void ProcessHtmlElement<T>(HtmlEnumerator en, Ox.OpenXmlLeafElement style) where T : Ox.OpenXmlLeafElement
		{
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>() { style };
			ProcessContainerAttributes(en, styleAttributes);
			htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
		}

		#endregion

		#region ProcessFigureCaption

		private void ProcessFigureCaption(HtmlEnumerator en)
		{
			this.CompleteCurrentParagraph(true);

			currentParagraph.Append(
					new OxW.ParagraphProperties {
						ParagraphStyleId = new OxW.ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.CaptionStyle, OxW.StyleValues.Paragraph) },
						KeepNext = new OxW.KeepNext()
					},
					new OxW.Run(
						new OxW.Text("Figure ") { Space = Ox.SpaceProcessingModeValues.Preserve }
					),
					new OxW.SimpleField(
						new OxW.Run(
							new OxW.Text(AddFigureCaption().ToString(CultureInfo.InvariantCulture)))
					) { Instruction = " SEQ Figure \\* ARABIC " }
				);

			ProcessHtmlChunks(en, "</figcaption>");

			if (elements.Count > 0) // any caption?
			{
				OxW.Text t = (elements[0] as OxW.Run).GetFirstChild<OxW.Text>();
				t.Text = " " + t.InnerText; // append a space after the numero of the picture
			}

			this.CompleteCurrentParagraph(true);
		}

		#endregion

		#region ProcessImage

		private void ProcessImage(HtmlEnumerator en)
		{
			OxW.Drawing drawing = null;
			OxW.Border border = new OxW.Border() { Val = OxW.BorderValues.None };
			string src = en.Attributes["src"];
			Uri uri = null;

			// Bug reported by Erik2014. Inline 64 bit images can be too big and Uri.TryCreate will fail silently with a SizeLimit error.
			// To circumvent this buffer size, we will work either on the Uri, either on the original src.
			if (src != null && (IO.DataUri.IsWellFormed(src) || Uri.TryCreate(src, UriKind.RelativeOrAbsolute, out uri)))
			{
				string alt = (en.Attributes["title"] ?? en.Attributes["alt"]) ?? String.Empty;

				Size preferredSize = Size.Empty;
				Unit wu = en.Attributes.GetAsUnit("width");
				if (!wu.IsValid) wu = en.StyleAttributes.GetAsUnit("width");
				Unit hu = en.Attributes.GetAsUnit("height");
				if (!hu.IsValid) hu = en.StyleAttributes.GetAsUnit("height");

				// % is not supported
				if (wu.IsFixed && wu.Value > 0)
				{
					preferredSize.Width = wu.ValueInPx;
				}
                if (hu.IsFixed && hu.Value > 0)
				{
					// Image perspective skewed. Bug fixed by ddeforge on github.com/onizet/html2openxml/discussions/350500
					preferredSize.Height = hu.ValueInPx;
				}

				SideBorder attrBorder = en.StyleAttributes.GetAsSideBorder("border");
				if (attrBorder.IsValid)
				{
					border.Val = attrBorder.Style;
					border.Color = attrBorder.Color.ToHexString();
					border.Size = (uint) attrBorder.Width.ValueInPx * 4;
				}
				else
				{
					var attrBorderWidth = en.Attributes.GetAsUnit("border");
					if (attrBorderWidth.IsValid)
					{
						border.Val = OxW.BorderValues.Single;
						border.Size = (uint) attrBorderWidth.ValueInPx * 4;
					}
				}

				drawing = AddImagePart(src, alt, preferredSize);
			}

			if (drawing != null)
			{
				OxW.Run run = new OxW.Run(drawing);
				if (border.Val != OxW.BorderValues.None) run.InsertInProperties(prop => prop.Border = border);
				elements.Add(run);
			}
		}

		#endregion

		#region ProcessLi

		private void ProcessLi(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(false);
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			int numberingId = htmlStyles.NumberingList.ProcessItem(en);
			int level = htmlStyles.NumberingList.LevelIndex;

			// Save the new paragraph reference to support nested numbering list.
			OxW.Paragraph p = currentParagraph;
			currentParagraph.InsertInProperties(prop => {
				prop.ParagraphStyleId = new OxW.ParagraphStyleId() { Val = GetStyleIdForListItem(en) };
				prop.Indentation = level < 2 ? null : new OxW.Indentation() { Left = (level * 780).ToString(CultureInfo.InvariantCulture) };
				prop.NumberingProperties = new OxW.NumberingProperties {
					NumberingLevelReference = new OxW.NumberingLevelReference() { Val = level - 1 },
					NumberingId = new OxW.NumberingId() { Val = numberingId }
				};
			});

			// Restore the original elements list
			AddParagraph(currentParagraph);

			// Continue to process the html until we found </li>
			HtmlStyles.Paragraph.ApplyTags(currentParagraph);
			AlternateProcessHtmlChunks(en, "</li>");
			p.Append(elements);
			this.elements.Clear();
		}

        private string GetStyleIdForListItem(HtmlEnumerator en) 
        { 
            return GetStyleIdFromClasses(en.Attributes.GetAsClass()) 
                   ?? GetStyleIdFromClasses(htmlStyles.NumberingList.GetCurrentListClasses) 
                   ?? htmlStyles.DefaultStyles.ListParagraphStyle; 
        }

        private string GetStyleIdFromClasses(string[] classes)  
        {  
            if (classes != null) 
            { 
                foreach (string className in classes) 
                {
					string styleId = htmlStyles.GetStyle(className, OxW.StyleValues.Paragraph, ignoreCase: true); 
                    if (styleId != null) 
                    { 
                        return styleId; 
                    } 
                } 
            } 
			 
            return null; 
        }

        #endregion

		#region ProcessLink

		private void ProcessLink(HtmlEnumerator en)
		{
			String att = en.Attributes["href"];
			OxW.Hyperlink h = null;
			Uri uri = null;


			if (!String.IsNullOrEmpty(att))
			{
				// handle link where the http:// is missing and that starts directly with www
				if(att.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
					att = "http://" + att;

				// is it an anchor?
				if (att[0] == '#' && att.Length > 1)
				{
					// Always accept _top anchor
					if (!this.ExcludeLinkAnchor || att == "#_top")
					{
						h = new OxW.Hyperlink(
							) { History = true, Anchor = att.Substring(1) };
					}
				}
				// ensure the links does not start with javascript:
				else if (Uri.TryCreate(att, UriKind.Absolute, out uri) && uri.Scheme != "javascript")
				{
					OxP.HyperlinkRelationship extLink = mainPart.AddHyperlinkRelationship(uri, true);

					h = new OxW.Hyperlink(
						) { History = true, Id = extLink.Id };
				}
			}

			if (h == null)
			{
				// link to a broken url, simply process the content of the tag
				ProcessHtmlChunks(en, "</a>");
				return;
			}

			att = en.Attributes["title"];
			if (!String.IsNullOrEmpty(att)) h.Tooltip = att;

			AlternateProcessHtmlChunks(en, "</a>");

			if (elements.Count == 0) return;

			// Let's see whether the link tag include an image inside its body.
			// If so, the Hyperlink OpenXmlElement is lost and we'll keep only the images
			// and applied a HyperlinkOnClick attribute.
			List<Ox.OpenXmlElement> imageInLink = elements.FindAll(e => { return e.HasChild<OxW.Drawing>(); });
			if (imageInLink.Count != 0)
			{
				for (int i = 0; i < imageInLink.Count; i++)
				{
					// Retrieves the "alt" attribute of the image and apply it as the link's tooltip
					OxW.Drawing d = imageInLink[i].GetFirstChild<OxW.Drawing>();
					var enDp = d.Descendants<OxD.Pictures.NonVisualDrawingProperties>().GetEnumerator();
					String alt;
					if (enDp.MoveNext()) alt = enDp.Current.Description;
					else alt = null;

					d.InsertInDocProperties(
							new OxD.HyperlinkOnClick() { Id = h.Id ?? h.Anchor, Tooltip = alt });
				}
			}

			// Append the processed elements and put them to the Run of the Hyperlink
			h.Append(elements);

			// can't use GetFirstChild<Run> or we may find the one containing the image
			foreach (var el in h.ChildElements)
			{
				OxW.Run run = el as OxW.Run;
				if (run != null && !run.HasChild<OxW.Drawing>())
				{
					run.InsertInProperties(prop =>
						prop.RunStyle = new OxW.RunStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.HyperlinkStyle, OxW.StyleValues.Character) });
					break;
				}
			}

			this.elements.Clear();

			// Append the hyperlink
			elements.Add(h);

			if (imageInLink.Count > 0) CompleteCurrentParagraph(true);
		}

		#endregion

		#region ProcessNumberingList

		private void ProcessNumberingList(HtmlEnumerator en)
		{
			htmlStyles.NumberingList.BeginList(en);
		}

		#endregion

		#region ProcessParagraph

		private void ProcessParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			// Respect this order: this is the way the browsers apply them
			String attrValue = en.StyleAttributes["text-align"];
			if (attrValue == null) attrValue = en.Attributes["align"];

			if (attrValue != null)
			{
				OxW.JustificationValues? align = Converter.ToParagraphAlign(attrValue);
				if (align.HasValue)
				{
					currentParagraph.InsertInProperties(prop => prop.Justification = new OxW.Justification { Val = align });
				}
			}

			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
				ProcessClosingParagraph(en);
			}
		}

		#endregion

		#region ProcessPre

		private void ProcessPre(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			// Oftenly, <pre> tag are used to renders some code examples. They look better inside a table
            if (this.RenderPreAsTable)
            {
				OxW.Table currentTable = new OxW.Table(
					new OxW.TableProperties(
						new OxW.TableStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.PreTableStyle, OxW.StyleValues.Table) },
						new OxW.TableWidth() { Type = OxW.TableWidthUnitValues.Pct, Width = "5000" } // 100% * 50
					),
					new OxW.TableGrid(
						new OxW.GridColumn() { Width = "5610" }),
					new OxW.TableRow(
						new OxW.TableCell(
                    // Ensure the border lines are visible (regardless of the style used)
							new OxW.TableCellProperties
                            {
								TableCellBorders = new OxW.TableCellBorders(
								   new OxW.TopBorder() { Val = OxW.BorderValues.Single },
								   new OxW.LeftBorder() { Val = OxW.BorderValues.Single },
								   new OxW.BottomBorder() { Val = OxW.BorderValues.Single },
								   new OxW.RightBorder() { Val = OxW.BorderValues.Single })
                            },
                            currentParagraph))
                );

                AddParagraph(currentTable);
                tables.NewContext(currentTable);
            }
            else
            {
                AddParagraph(currentParagraph);
            }

			// Process the entire <pre> tag and append it to the document
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());

			AlternateProcessHtmlChunks(en, "</pre>");

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.EndTag(en.CurrentTag);

			if (RenderPreAsTable)
				tables.CloseContext();

			CompleteCurrentParagraph();
		}

		#endregion

		#region ProcessQuote

		private void ProcessQuote(HtmlEnumerator en)
		{
			// The browsers render the quote tag between a kind of separators.
			// We add the Quote style to the nested runs to match more Word.

			OxW.Run run = new OxW.Run(
				new OxW.Text(" " + HtmlStyles.QuoteCharacters.Prefix) { Space = Ox.SpaceProcessingModeValues.Preserve }
			);

			htmlStyles.Runs.ApplyTags(run);
			elements.Add(run);

			ProcessHtmlElement<OxW.RunStyle>(en, new OxW.RunStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.QuoteStyle, OxW.StyleValues.Character) });
		}

		#endregion

		#region ProcessSpan

		private void ProcessSpan(HtmlEnumerator en)
		{
			// A span style attribute can contains many information: font color, background color, font size,
			// font family, ...
			// We'll check for each of these and add apply them to the next build runs.

			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);

			if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
				CompleteCurrentParagraph(true);
			}
		}

		#endregion

		#region ProcessSubscript

		private void ProcessSubscript(HtmlEnumerator en)
		{
			ProcessHtmlElement<OxW.VerticalTextAlignment>(en, new OxW.VerticalTextAlignment() { Val = OxW.VerticalPositionValues.Subscript });
		}

		#endregion

		#region ProcessSuperscript

		private void ProcessSuperscript(HtmlEnumerator en)
		{
			ProcessHtmlElement<OxW.VerticalTextAlignment>(en, new OxW.VerticalTextAlignment() { Val = OxW.VerticalPositionValues.Superscript });
		}

		#endregion

		#region ProcessUnderline

		private void ProcessUnderline(HtmlEnumerator en)
		{
			ProcessHtmlElement<OxW.Underline>(en, new OxW.Underline() { Val = OxW.UnderlineValues.Single });
		}

		#endregion

		#region ProcessTable

		private void ProcessTable(HtmlEnumerator en)
		{
			OxW.TableProperties properties = new OxW.TableProperties(
				new OxW.TableStyle() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.TableStyle, OxW.StyleValues.Table) }
			);
			OxW.Table currentTable = new OxW.Table(properties);

			string classValue = en.Attributes["class"];
			if (classValue != null)
			{
				classValue = htmlStyles.GetStyle(classValue, OxW.StyleValues.Table, ignoreCase: true);
				if (classValue != null)
					properties.TableStyle.Val = classValue;
			}

			int? border = en.Attributes.GetAsInt("border");
			if (border.HasValue && border.Value > 0)
			{
				bool handleBorders = true;
				if (classValue != null)
				{
					// check whether the style in use have borders
					String styleId = this.htmlStyles.GetStyle(classValue, OxW.StyleValues.Table, true);
					if (styleId != null)
                    {
						var s = mainPart.StyleDefinitionsPart.Styles.Elements<OxW.Style>().First(e => e.StyleId == styleId);
                        if (s.StyleTableProperties.TableBorders != null) handleBorders = false;
                    }
				}

				// If the border has been specified, we display the Table Grid style which display
				// its grid lines. Otherwise the default table style hides the grid lines.
				if (handleBorders && properties.TableStyle.Val != htmlStyles.DefaultStyles.TableStyle)
				{
					uint borderSize = border.Value > 1? (uint) new Unit(UnitMetric.Pixel, border.Value).ValueInDxa : 1;
					properties.TableBorders = new OxW.TableBorders() {
						TopBorder = new OxW.TopBorder { Val = OxW.BorderValues.None },
						LeftBorder = new OxW.LeftBorder { Val = OxW.BorderValues.None },
						RightBorder = new OxW.RightBorder { Val = OxW.BorderValues.None },
						BottomBorder = new OxW.BottomBorder { Val = OxW.BorderValues.None },
						InsideHorizontalBorder = new OxW.InsideHorizontalBorder { Val = OxW.BorderValues.Single, Size = borderSize },
						InsideVerticalBorder = new OxW.InsideVerticalBorder { Val = OxW.BorderValues.Single, Size = borderSize }
					};
				}
			}
			// is the border=0? If so, we remove the border regardless the style in use
			else if (border == 0)
			{
				properties.TableBorders = new OxW.TableBorders() {
					TopBorder = new OxW.TopBorder { Val = OxW.BorderValues.None },
					LeftBorder = new OxW.LeftBorder { Val = OxW.BorderValues.None },
					RightBorder = new OxW.RightBorder { Val = OxW.BorderValues.None },
					BottomBorder = new OxW.BottomBorder { Val = OxW.BorderValues.None },
					InsideHorizontalBorder = new OxW.InsideHorizontalBorder { Val = OxW.BorderValues.None },
					InsideVerticalBorder = new OxW.InsideVerticalBorder { Val = OxW.BorderValues.None }
				};
			}
			else
            {
				var styleBorder = en.StyleAttributes.GetAsBorder("border");
				if (!styleBorder.IsEmpty)
				{
					properties.TableBorders = new OxW.TableBorders();

					if (styleBorder.Left.IsValid)
						properties.TableBorders.LeftBorder = new OxW.LeftBorder { Val = styleBorder.Left.Style, Color = Ox.StringValue.FromString(styleBorder.Left.Color.ToHexString()), Size = (uint)styleBorder.Left.Width.ValueInDxa };
					if (styleBorder.Right.IsValid)
						properties.TableBorders.RightBorder = new OxW.RightBorder { Val = styleBorder.Right.Style, Color = Ox.StringValue.FromString(styleBorder.Right.Color.ToHexString()), Size = (uint)styleBorder.Right.Width.ValueInDxa };
					if (styleBorder.Top.IsValid)
						properties.TableBorders.TopBorder = new OxW.TopBorder { Val = styleBorder.Top.Style, Color = Ox.StringValue.FromString(styleBorder.Top.Color.ToHexString()), Size = (uint)styleBorder.Top.Width.ValueInDxa };
					if (styleBorder.Bottom.IsValid)
						properties.TableBorders.BottomBorder = new OxW.BottomBorder { Val = styleBorder.Bottom.Style, Color = Ox.StringValue.FromString(styleBorder.Bottom.Color.ToHexString()), Size = (uint)styleBorder.Bottom.Width.ValueInDxa };
				}
			}			

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case UnitMetric.Percent:
						properties.TableWidth = new OxW.TableWidth() { Type = OxW.TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) }; break;
					case UnitMetric.Point:
						properties.TableWidth = new OxW.TableWidth() { Type = OxW.TableWidthUnitValues.Dxa, Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) }; break;
					case UnitMetric.Pixel:
						properties.TableWidth = new OxW.TableWidth() { Type = OxW.TableWidthUnitValues.Dxa, Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) }; break;
				}
			}
			else
			{
				// Use Auto=0 instead of Pct=auto
				// bug reported by scarhand (https://html2openxml.codeplex.com/workitem/12494)
				properties.TableWidth = new OxW.TableWidth() { Type = OxW.TableWidthUnitValues.Auto, Width = "0" };
			}

			string align = en.Attributes["align"];
			if (align != null)
			{
				OxW.JustificationValues? halign = Converter.ToParagraphAlign(align);
				if (halign.HasValue)
					properties.TableJustification = new OxW.TableJustification() { Val = halign.Value.ToTableRowAlignment() };
			}

			// only if the table is left aligned, we can handle some left margin indentation
			// Right margin + Right align has no equivalent in OpenXml
			if (align == null || align == "left")
			{
				Margin margin = en.StyleAttributes.GetAsMargin("margin");

				// OpenXml doesn't support table margin in Percent, but Html does
				// the margin part has been implemented by Olek (patch #8457)

				OxW.TableCellMarginDefault cellMargin = new OxW.TableCellMarginDefault();
                if (margin.Left.IsFixed)
					cellMargin.TableCellLeftMargin = new OxW.TableCellLeftMargin() { Type = OxW.TableWidthValues.Dxa, Width = (short)margin.Left.ValueInDxa };
                if (margin.Right.IsFixed)
					cellMargin.TableCellRightMargin = new OxW.TableCellRightMargin() { Type = OxW.TableWidthValues.Dxa, Width = (short)margin.Right.ValueInDxa };
                if (margin.Top.IsFixed)
					cellMargin.TopMargin = new OxW.TopMargin() { Type = OxW.TableWidthUnitValues.Dxa, Width = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
                if (margin.Bottom.IsFixed)
					cellMargin.BottomMargin = new OxW.BottomMargin() { Type = OxW.TableWidthUnitValues.Dxa, Width = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture) };

                // Align table according to the margin 'auto' as it stands in Html
                if (margin.Left.Type == UnitMetric.Auto || margin.Right.Type == UnitMetric.Auto)
                {
					OxW.TableRowAlignmentValues justification;

                    if (margin.Left.Type == UnitMetric.Auto && margin.Right.Type == UnitMetric.Auto)
						justification = OxW.TableRowAlignmentValues.Center;
                    else if (margin.Left.Type == UnitMetric.Auto)
						justification = OxW.TableRowAlignmentValues.Right;
                    else
						justification = OxW.TableRowAlignmentValues.Left;

					properties.TableJustification = new OxW.TableJustification() { Val = justification };
                }

				if (cellMargin.HasChildren)
					properties.TableCellMarginDefault = cellMargin;
			}

			int? spacing = en.Attributes.GetAsInt("cellspacing");
			if (spacing.HasValue)
				properties.TableCellSpacing = new OxW.TableCellSpacing { Type = OxW.TableWidthUnitValues.Dxa, Width = new Unit(UnitMetric.Pixel, spacing.Value).ValueInDxa.ToString(CultureInfo.InvariantCulture) };

			int? padding = en.Attributes.GetAsInt("cellpadding");
            if (padding.HasValue)
            {
                int paddingDxa = (int) new Unit(UnitMetric.Pixel, padding.Value).ValueInDxa;

				OxW.TableCellMarginDefault cellMargin = new OxW.TableCellMarginDefault();
				cellMargin.TableCellLeftMargin = new OxW.TableCellLeftMargin() { Type = OxW.TableWidthValues.Dxa, Width = (short)paddingDxa };
				cellMargin.TableCellRightMargin = new OxW.TableCellRightMargin() { Type = OxW.TableWidthValues.Dxa, Width = (short)paddingDxa };
				cellMargin.TopMargin = new OxW.TopMargin() { Type = OxW.TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) };
				cellMargin.BottomMargin = new OxW.BottomMargin() { Type = OxW.TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) };
                properties.TableCellMarginDefault = cellMargin;
            }

			List<Ox.OpenXmlElement> runStyleAttributes = new List<Ox.OpenXmlElement>();
			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());


			// are we currently inside another table?
			if (tables.HasContext)
			{
				// Okay we will insert nested table but beware the paragraph inside OxW.TableCell should contains at least 1 run.

				OxW.TableCell currentCell = tables.CurrentTable.GetLastChild<OxW.TableRow>().GetLastChild<OxW.TableCell>();
				// don't add an empty paragraph if not required (bug #13608 by zanjo)
				if (elements.Count == 0) currentCell.Append(currentTable);
				else
				{
					currentCell.Append(new OxW.Paragraph(elements), currentTable);
					elements.Clear();
				}
			}
			else
			{
				CompleteCurrentParagraph();
				this.paragraphs.Add(currentTable);
			}

			tables.NewContext(currentTable);
		}

		#endregion

		#region ProcessTableCaption

		private void ProcessTableCaption(HtmlEnumerator en)
		{
			if (!tables.HasContext) return;

			string att = en.StyleAttributes["text-align"];
			if (att == null) att = en.Attributes["align"];

			ProcessHtmlChunks(en, "</caption>");

			var legend = new OxW.Paragraph(
					new OxW.ParagraphProperties {
						ParagraphStyleId = new OxW.ParagraphStyleId() { Val = htmlStyles.GetStyle(htmlStyles.DefaultStyles.CaptionStyle, OxW.StyleValues.Paragraph) }
					},
					new OxW.Run(
						new OxW.FieldChar() { FieldCharType = OxW.FieldCharValues.Begin }),
					new OxW.Run(
						new OxW.FieldCode(" SEQ TABLE \\* ARABIC ") { Space = Ox.SpaceProcessingModeValues.Preserve }),
					new OxW.Run(
						new OxW.FieldChar() { FieldCharType = OxW.FieldCharValues.End })
				);
			legend.Append(elements);
			elements.Clear();

			if (att != null)
			{
				OxW.JustificationValues? align = Converter.ToParagraphAlign(att);
				if (align.HasValue)
					legend.InsertInProperties(prop => prop.Justification = new OxW.Justification { Val = align });
			}
			else
			{
				// If no particular alignement has been specified for the legend, we will align the legend
				// relative to the owning table
				OxW.TableProperties props = tables.CurrentTable.GetFirstChild<OxW.TableProperties>();
				if (props != null)
				{
					OxW.TableJustification justif = props.GetFirstChild<OxW.TableJustification>();
					if (justif != null) legend.InsertInProperties(prop =>
						prop.Justification = new OxW.Justification { Val = justif.Val.Value.ToJustification() });
				}
			}

			if (this.TableCaptionPosition == OxW.CaptionPositionValues.Above)
				this.paragraphs.Insert(this.paragraphs.Count - 1, legend);
			else
				this.paragraphs.Add(legend);
		}

		#endregion

		#region ProcessTableRow

		private void ProcessTableRow(HtmlEnumerator en)
		{
			// in case the html is bad-formed and use <tr> outside a <table> tag, we will ensure
			// a table context exists.
			if (!tables.HasContext) return;

			OxW.TableRowProperties properties = new OxW.TableRowProperties();
			List<Ox.OpenXmlElement> runStyleAttributes = new List<Ox.OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);

			Unit unit = en.StyleAttributes.GetAsUnit("height");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("height");

			switch (unit.Type)
			{
				case UnitMetric.Point:
					properties.AddChild(new OxW.TableRowHeight() { HeightType = OxW.HeightRuleValues.AtLeast, Val = (uint)(unit.Value * 20) });
					break;
				case UnitMetric.Pixel:
					properties.AddChild(new OxW.TableRowHeight() { HeightType = OxW.HeightRuleValues.AtLeast, Val = (uint)unit.ValueInDxa });
					break;
			}

			// Do not explicitly set the tablecell spacing in order to inherit table style (issue 107)
			//properties.AddChild(new TableCellSpacing() { Type = TableWidthUnitValues.Dxa, Width = "0" });

			OxW.TableRow row = new OxW.TableRow();
			row.TableRowProperties = properties;

			htmlStyles.Runs.ProcessCommonAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

			tables.CurrentTable.Append(row);
			tables.CellPosition = new CellPosition(tables.CellPosition.Row + 1, 0);
		}

		#endregion

		#region ProcessTableColumn

		private void ProcessTableColumn(HtmlEnumerator en)
		{
			if (!tables.HasContext) return;

			OxW.TableCellProperties properties = new OxW.TableCellProperties();
            // in Html, table cell are vertically centered by default
            properties.TableCellVerticalAlignment = new OxW.TableCellVerticalAlignment() { Val = OxW.TableVerticalAlignmentValues.Center };

			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();
			List<Ox.OpenXmlElement> runStyleAttributes = new List<Ox.OpenXmlElement>();

			Unit unit = en.StyleAttributes.GetAsUnit("width");
			if (!unit.IsValid) unit = en.Attributes.GetAsUnit("width");

            // The heightUnit used to retrieve a height value.
            Unit heightUnit = en.StyleAttributes.GetAsUnit("height");
            if (!heightUnit.IsValid) heightUnit = en.Attributes.GetAsUnit("height");

            switch (unit.Type)
			{
				case UnitMetric.Percent:
					properties.TableCellWidth = new OxW.TableCellWidth() { Type = OxW.TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) };
					break;
				case UnitMetric.Point:
                    // unit.ValueInPoint used instead of ValueInDxa
					properties.TableCellWidth = new OxW.TableCellWidth() { Type = OxW.TableWidthUnitValues.Auto, Width = (unit.ValueInPoint * 20).ToString(CultureInfo.InvariantCulture) };
					break;
				case UnitMetric.Pixel:
					properties.TableCellWidth = new OxW.TableCellWidth() { Type = OxW.TableWidthUnitValues.Dxa, Width = (unit.ValueInDxa).ToString(CultureInfo.InvariantCulture) };
					break;
			}

			// fix an issue when specifying the RowSpan or ColSpan=1 (reported by imagremlin)
			int? colspan = en.Attributes.GetAsInt("colspan");
			if (colspan.HasValue && colspan.Value > 1)
			{
				properties.GridSpan = new OxW.GridSpan() { Val = colspan };
			}

			int? rowspan = en.Attributes.GetAsInt("rowspan");
			if (rowspan.HasValue && rowspan.Value > 1)
			{
				properties.VerticalMerge = new OxW.VerticalMerge() { Val = OxW.MergedCellValues.Restart };

				var p = tables.CellPosition;
                int shift = 0;
                // if there is already a running rowSpan on a left-sided column, we have to shift this position
                foreach (var rs in tables.RowSpan)
                    if (rs.CellOrigin.Row < p.Row && rs.CellOrigin.Column <= p.Column + shift) shift++;

                p.Offset(0, shift);
                tables.RowSpan.Add(new HtmlTableSpan(p) {
                    RowSpan = rowspan.Value - 1,
                    ColSpan = colspan.HasValue && rowspan.Value > 1 ? colspan.Value : 0
                });
			}

			// Manage vertical text (only for table cell)
			string direction = en.StyleAttributes["writing-mode"];
			if (direction != null)
			{
				switch (direction)
				{
					case "tb-lr":
						properties.TextDirection = new OxW.TextDirection() { Val = OxW.TextDirectionValues.BottomToTopLeftToRight };
						properties.TableCellVerticalAlignment = new OxW.TableCellVerticalAlignment() { Val = OxW.TableVerticalAlignmentValues.Center };
						htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new OxW.Justification() { Val = OxW.JustificationValues.Center });
						break;
					case "tb-rl":
						properties.TextDirection = new OxW.TextDirection() { Val = OxW.TextDirectionValues.TopToBottomRightToLeft };
						properties.TableCellVerticalAlignment = new OxW.TableCellVerticalAlignment() { Val = OxW.TableVerticalAlignmentValues.Center };
						htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new OxW.Justification() { Val = OxW.JustificationValues.Center });
						break;
				}
			}

			var padding = en.StyleAttributes.GetAsMargin("padding");
			if (!padding.IsEmpty)
			{
				OxW.TableCellMargin cellMargin = new OxW.TableCellMargin();
				var cellMarginSide = new List<KeyValuePair<Unit, OxW.TableWidthType>>();
				cellMarginSide.Add(new KeyValuePair<Unit, OxW.TableWidthType>(padding.Top, new OxW.TopMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, OxW.TableWidthType>(padding.Left, new OxW.LeftMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, OxW.TableWidthType>(padding.Bottom, new OxW.BottomMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, OxW.TableWidthType>(padding.Right, new OxW.RightMargin()));

				foreach (var pair in cellMarginSide)
				{
					if (!pair.Key.IsValid || pair.Key.Value == 0) continue;
					if (pair.Key.Type == UnitMetric.Percent)
					{
						pair.Value.Width = (pair.Key.Value * 50).ToString(CultureInfo.InvariantCulture);
						pair.Value.Type = OxW.TableWidthUnitValues.Pct;
					}
					else
					{
						pair.Value.Width = pair.Key.ValueInDxa.ToString(CultureInfo.InvariantCulture);
						pair.Value.Type = OxW.TableWidthUnitValues.Dxa;
					}

					cellMargin.AddChild(pair.Value);
				}

				properties.TableCellMargin = cellMargin;
			}

			var border = en.StyleAttributes.GetAsBorder("border");
			if (!border.IsEmpty)
			{
				properties.TableCellBorders = new OxW.TableCellBorders();

				if (border.Left.IsValid)
					properties.TableCellBorders.LeftBorder = new OxW.LeftBorder { Val = border.Left.Style, Color = Ox.StringValue.FromString(border.Left.Color.ToHexString()), Size = (uint)border.Left.Width.ValueInDxa };
				if (border.Right.IsValid)
					properties.TableCellBorders.RightBorder = new OxW.RightBorder { Val = border.Right.Style, Color = Ox.StringValue.FromString(border.Right.Color.ToHexString()), Size = (uint)border.Right.Width.ValueInDxa };
				if (border.Top.IsValid)
					properties.TableCellBorders.TopBorder = new OxW.TopBorder { Val = border.Top.Style, Color = Ox.StringValue.FromString(border.Top.Color.ToHexString()), Size = (uint)border.Top.Width.ValueInDxa };
				if (border.Bottom.IsValid)
					properties.TableCellBorders.BottomBorder = new OxW.BottomBorder { Val = border.Bottom.Style, Color = Ox.StringValue.FromString(border.Bottom.Color.ToHexString()), Size = (uint)border.Bottom.Width.ValueInDxa };
			}

			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
			if (styleAttributes.Count > 0)
				htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes);
			if (runStyleAttributes.Count > 0)
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());

			OxW.TableCell cell = new OxW.TableCell();
			if (properties.HasChildren) cell.TableCellProperties = properties;
                  
            // The heightUnit value used to append a height to the TableRowHeight.
            var row = tables.CurrentTable.GetLastChild<OxW.TableRow>();

            switch (heightUnit.Type)
            {
                case UnitMetric.Point:
					row.TableRowProperties.AddChild(new OxW.TableRowHeight() { HeightType = OxW.HeightRuleValues.AtLeast, Val = (uint)(heightUnit.Value * 20) });

                    break;
                case UnitMetric.Pixel:
					row.TableRowProperties.AddChild(new OxW.TableRowHeight() { HeightType = OxW.HeightRuleValues.AtLeast, Val = (uint)heightUnit.ValueInDxa });
                    break;
            }

            row.Append(cell);

            if (en.IsSelfClosedTag) // Force a call to ProcessClosingTableColumn
				ProcessClosingTableColumn(en);
			else
			{
				// we create a new currentParagraph to add new runs inside the OxW.TableCell
				cell.Append(currentParagraph = new OxW.Paragraph());
			}
		}

		#endregion

		#region ProcessTablePart

		private void ProcessTablePart(HtmlEnumerator en)
		{
			List<Ox.OpenXmlElement> styleAttributes = new List<Ox.OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
				htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes.ToArray());
		}

		#endregion

		#region ProcessXmlDataIsland

		private void ProcessXmlDataIsland(HtmlEnumerator en)
		{
			// Process inner Xml data island and do nothing.
			// The Xml has this format:
			/* <?xml:namespace prefix=o ns="urn:schemas-microsoft-com:office:office">
			   <globalGuideLine>
				   <employee>
					  <FirstName>Austin</FirstName>
					  <LastName>Hennery</LastName>
				   </employee>
			   </globalGuideLine>
			 */

			// Move to the first root element of the Xml then process until the end of the xml chunks.
			while (en.MoveNext() && !en.IsCurrentHtmlTag) ;

			if (en.Current != null)
			{
				string xmlRootElement = en.ClosingCurrentTag;
				while (en.MoveUntilMatch(xmlRootElement)) ;
			}
		}

		#endregion

		// Closing tags

		#region ProcessClosingBlockQuote

		private void ProcessClosingBlockQuote(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);
			htmlStyles.Paragraph.EndTag("<blockquote>");
		}

		#endregion

		#region ProcessClosingDiv

		private void ProcessClosingDiv(HtmlEnumerator en)
		{
			// Mimic the rendering of the browser:
			ProcessBr(en);
			ProcessClosingTag(en);
		}

		#endregion

		#region ProcessClosingTag

		private void ProcessClosingTag(HtmlEnumerator en)
		{
			string openingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(openingTag);
			htmlStyles.Paragraph.EndTag(openingTag);
		}

		#endregion

		#region ProcessClosingNumberingList

		private void ProcessClosingNumberingList(HtmlEnumerator en)
		{
			htmlStyles.NumberingList.EndList();

			// If we are no more inside a list, we move to another paragraph (as we created
			// one for containing all the <li>. This will ensure the next run will not be added to the <li>.
			if (htmlStyles.NumberingList.LevelIndex == 0)
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessClosingParagraph

		private void ProcessClosingParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			string tag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(tag);
			htmlStyles.Paragraph.EndTag(tag);
		}

		#endregion

		#region ProcessClosingQuote

		private void ProcessClosingQuote(HtmlEnumerator en)
		{
			OxW.Run run = new OxW.Run(
				new OxW.Text(HtmlStyles.QuoteCharacters.Suffix) { Space = Ox.SpaceProcessingModeValues.Preserve }
			);
			htmlStyles.Runs.ApplyTags(run);
			elements.Add(run);

			htmlStyles.Runs.EndTag("<q>");
		}

		#endregion

		#region ProcessClosingTable

		private void ProcessClosingTable(HtmlEnumerator en)
		{
			htmlStyles.Tables.EndTag("<table>");
			htmlStyles.Runs.EndTag("<table>");

			OxW.TableRow row = tables.CurrentTable.GetFirstChild<OxW.TableRow>();
			// Is this a misformed or empty table?
			if (row != null)
			{
				// Count the number of tableCell and add as much GridColumn as we need.
				OxW.TableGrid grid = new OxW.TableGrid();
				foreach (OxW.TableCell cell in row.Elements<OxW.TableCell>())
				{
					// If that column contains some span, we need to count them also
					int count = cell.TableCellProperties.GridSpan != null ? cell.TableCellProperties.GridSpan.Val.Value : 1;
					for (int i=0; i<count; i++) {
						grid.Append(new OxW.GridColumn());
					}
				}

				tables.CurrentTable.InsertAt<OxW.TableGrid>(grid, 1);
			}

			tables.CloseContext();

			if (!tables.HasContext)
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessClosingTablePart

		private void ProcessClosingTablePart(HtmlEnumerator en)
		{
			string closingTag = en.CurrentTag.Replace("/", "");

			htmlStyles.Tables.EndTag(closingTag);
		}

		#endregion

		#region ProcessClosingTableRow

		private void ProcessClosingTableRow(HtmlEnumerator en)
		{
			if (!tables.HasContext) return;
			OxW.TableRow row = tables.CurrentTable.GetLastChild<OxW.TableRow>();
			if (row == null) return;

			// Word will not open documents with empty rows (reported by scwebgroup)
			if (row.GetFirstChild<OxW.TableCell>() == null)
			{
				row.Remove();
				return;
			}

			// Add empty columns to fill rowspan
			if (tables.RowSpan.Count > 0)
			{
				int rowIndex = tables.CellPosition.Row;

				for (int i = 0; i < tables.RowSpan.Count; i++)
				{
					HtmlTableSpan tspan = tables.RowSpan[i];
					if (tspan.CellOrigin.Row == rowIndex) continue;

                    OxW.TableCell emptyCell = new OxW.TableCell(new OxW.TableCellProperties {
								            TableCellWidth = new OxW.TableCellWidth() { Width = "0" },
								            VerticalMerge = new OxW.VerticalMerge() },
										new OxW.Paragraph());

                    tspan.RowSpan--;
                    if (tspan.RowSpan == 0) { tables.RowSpan.RemoveAt(i); i--; }

                    // in case of both colSpan + rowSpan on the same cell, we have to reverberate the rowSpan on the next columns too
					if (tspan.ColSpan > 0) emptyCell.TableCellProperties.GridSpan = new OxW.GridSpan() { Val = tspan.ColSpan };

                    OxW.TableCell cell = row.GetFirstChild<OxW.TableCell>();
                    if (tspan.CellOrigin.Column == 0 || cell == null)
                    {
						row.InsertAfter(emptyCell, row.TableRowProperties);
                        continue;
                    }

                    // find the good column position, taking care of eventual colSpan
                    int columnIndex = 0;
                    while (columnIndex < tspan.CellOrigin.Column)
                    {
                        columnIndex += cell.TableCellProperties.GridSpan.Val ?? 1;
                    }
                    //while ((cell = cell.NextSibling<OxW.TableCell>()) != null);

                    if (cell == null) row.AppendChild(emptyCell);
                    else row.InsertAfter<OxW.TableCell>(emptyCell, cell);
                }
			}

			htmlStyles.Tables.EndTag("<tr>");
			htmlStyles.Runs.EndTag("<tr>");
		}

		#endregion

		#region ProcessClosingTableColumn

		private void ProcessClosingTableColumn(HtmlEnumerator en)
		{
			if (!tables.HasContext)
			{
				// When the Html is bad-formed and doesn't contain <table>, the browser renders the column separated by a space.
				// So we do the same here
				OxW.Run run = new OxW.Run(new OxW.Text(" ") { Space = Ox.SpaceProcessingModeValues.Preserve });
				htmlStyles.Runs.ApplyTags(run);
				elements.Add(run);
				return;
			}
			OxW.TableCell cell = tables.CurrentTable.GetLastChild<OxW.TableRow>().GetLastChild<OxW.TableCell>();

			// As we add automatically a paragraph to the cell once we create it, we'll remove it if finally, it was not used.
			// For all the other children, we will ensure there is no more empty paragraphs (similarly to what we do at the end
			// of the convert processing).
			// use a basic loop instead of foreach to allow removal (bug reported by antgraf)
			for (int i=0; i<cell.ChildElements.Count; )
			{
				OxW.Paragraph p = cell.ChildElements[i] as OxW.Paragraph;
				// care of hyperlinks as they are not inside Run (bug reported by mdeclercq github.com/onizet/html2openxml/workitem/11162)
				if (p != null && !p.HasChild<OxW.Run>() && !p.HasChild<OxW.Hyperlink>()) p.Remove();
				else i++;
			}

			// We add this paragraph regardless it has elements or not. A OxW.TableCell requires at least a Paragraph, as the last child of
			// of a table cell.
			// additional check for a proper cleaning (reported by antgraf github.com/onizet/html2openxml/discussions/272744)
			if (!(cell.LastChild is OxW.Paragraph) || elements.Count > 0) cell.Append(new OxW.Paragraph(elements));

			htmlStyles.Tables.ApplyTags(cell);

			// Reset all our variables and move to next cell
			this.elements.Clear();
			String openingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Tables.EndTag(openingTag);
			htmlStyles.Runs.EndTag(openingTag);

			var pos = tables.CellPosition;
			pos.Column++;
			tables.CellPosition = pos;
		}

		#endregion
	}
}

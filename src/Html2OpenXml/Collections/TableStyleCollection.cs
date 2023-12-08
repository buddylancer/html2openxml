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
using Ox = DocumentFormat.OpenXml;
using OxW = DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<Ox.OpenXmlElement>;


	sealed class TableStyleCollection : OpenXmlStyleCollectionBase
	{
		private readonly ParagraphStyleCollection paragraphStyle;
		private readonly HtmlDocumentStyle documentStyle;

        internal TableStyleCollection(HtmlDocumentStyle documentStyle)
		{
			this.documentStyle = documentStyle;
			paragraphStyle = new ParagraphStyleCollection(documentStyle);
		}

		internal override void Reset()
		{
			paragraphStyle.Reset();
			base.Reset();
		}

		//____________________________________________________________________
		//

		/// <summary>
		/// Apply all the current Html tag to the specified table cell.
		/// </summary>
		public override void ApplyTags(Ox.OpenXmlCompositeElement tableCell)
		{
			if (tags.Count > 0)
			{
				OxW.TableCellProperties properties = tableCell.GetFirstChild<OxW.TableCellProperties>();
				if (properties == null) tableCell.PrependChild<OxW.TableCellProperties>(properties = new OxW.TableCellProperties());

				var en = tags.GetEnumerator();
				while (en.MoveNext())
				{
					TagsAtSameLevel tagsOfSameLevel = en.Current.Value.Peek();
					foreach (Ox.OpenXmlElement tag in tagsOfSameLevel.Array)
						properties.AddChild(tag.CloneNode(true));
				}
			}

			// Apply some style attributes on the unique Paragraph tag contained inside a table cell.
			OxW.Paragraph p = tableCell.GetFirstChild<OxW.Paragraph>();
			paragraphStyle.ApplyTags(p);
		}

		public void BeginTagForParagraph(string name, params Ox.OpenXmlElement[] elements)
		{
			paragraphStyle.BeginTag(name, elements);
		}

		public override void EndTag(string name)
		{
			paragraphStyle.EndTag(name);
			base.EndTag(name);
		}

		#region ProcessCommonAttributes

		/// <summary>
		/// Move inside the current tag related to table (td, thead, tr, ...) and converts some common
		/// attributes to their OpenXml equivalence.
		/// </summary>
		/// <param name="en">The Html enumerator positionned on a <i>table (or related)</i> tag.</param>
		/// <param name="runStyleAttributes">The collection of attributes where to store new discovered attributes.</param>
		public void ProcessCommonAttributes(HtmlEnumerator en, IList<Ox.OpenXmlElement> runStyleAttributes)
		{
			List<Ox.OpenXmlElement> containerStyleAttributes = new List<Ox.OpenXmlElement>();

			var colorValue = en.StyleAttributes.GetAsColor("background-color");

            // "background-color" is also handled by RunStyleCollection which duplicate this attribute (bug #13212). 
			// Also apply on <th> (issue #20).
			// As on 05 Jan 2018, the duplication was due to the wrong argument passed during the td/th processing.
			// It was the runStyle and not the containerStyle that was provided. The code has been removed as no more useful
			if (colorValue.IsEmpty) colorValue = en.Attributes.GetAsColor("bgcolor");
            if (!colorValue.IsEmpty)
			{
				containerStyleAttributes.Add(
					new OxW.Shading() { Val = OxW.ShadingPatternValues.Clear, Color = "auto", Fill = colorValue.ToHexString() });
			}

			var htmlAlign = en.StyleAttributes["vertical-align"];
			if (htmlAlign == null) htmlAlign = en.Attributes["valign"];
			if (htmlAlign != null)
			{
				OxW.TableVerticalAlignmentValues? valign = Converter.ToVAlign(htmlAlign);
				if (valign.HasValue)
					containerStyleAttributes.Add(new OxW.TableCellVerticalAlignment() { Val = valign });
			}

			htmlAlign = en.StyleAttributes["text-align"];
			if (htmlAlign == null) htmlAlign = en.Attributes["align"];
			if (htmlAlign != null)
			{
				OxW.JustificationValues? halign = Converter.ToParagraphAlign(htmlAlign);
				if (halign.HasValue)
					this.BeginTagForParagraph(en.CurrentTag, new OxW.KeepNext(), new OxW.Justification { Val = halign });
			}

			// implemented by ddforge
			String[] classes = en.Attributes.GetAsClass();
			if (classes != null)
			{
				for (int i = 0; i < classes.Length; i++)
				{
					string className = documentStyle.GetStyle(classes[i], OxW.StyleValues.Table, ignoreCase: true);
					if (className != null) // only one Style can be applied in OpenXml and dealing with inheritance is out of scope
					{
						containerStyleAttributes.Add(new OxW.RunStyle() { Val = className });
						break;
					}
				}
			}

			this.BeginTag(en.CurrentTag, containerStyleAttributes);

			// Process general run styles
			documentStyle.Runs.ProcessCommonAttributes(en, runStyleAttributes);
		}

		#endregion
	}
}
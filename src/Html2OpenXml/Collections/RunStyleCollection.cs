﻿/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
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
using Ox = DocumentFormat.OpenXml;
using OxW = DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;


	sealed class RunStyleCollection : OpenXmlStyleCollectionBase
	{
		private readonly HtmlDocumentStyle documentStyle;

		internal RunStyleCollection(HtmlDocumentStyle documentStyle)
		{
			this.documentStyle = documentStyle;
		}

		/// <summary>
		/// Apply all the current Html tag (Run properties) to the specified run.
		/// </summary>
		public override void ApplyTags(Ox.OpenXmlCompositeElement run)
		{
			if (tags.Count == 0 && DefaultRunStyle == null) return;

			OxW.RunProperties properties = run.GetFirstChild<OxW.RunProperties>();
			if (properties == null) run.PrependChild<OxW.RunProperties>(properties = new OxW.RunProperties());

			var en = tags.GetEnumerator();
			while (en.MoveNext())
			{
				TagsAtSameLevel tagsOfSameLevel = en.Current.Value.Peek();
				foreach (Ox.OpenXmlElement tag in tagsOfSameLevel.Array)
					properties.AddChild(tag.CloneNode(true));
			}

			if (this.DefaultRunStyle != null)
				properties.RunStyle = new OxW.RunStyle() { Val = this.DefaultRunStyle };
		}

		#region ProcessCommonAttributes

		/// <summary>
		/// Converts some common styling attributes to their OpenXml equivalence.
		/// </summary>
        /// <param name="en">The Html parser.</param>
		/// <param name="styleAttributes">The collection of attributes where to store new discovered attributes.</param>
		public void ProcessCommonAttributes(HtmlEnumerator en, IList<Ox.OpenXmlElement> styleAttributes)
		{
			if (en.Attributes.Count == 0) return;

			var colorValue = en.StyleAttributes.GetAsColor("color");
			if (colorValue.IsEmpty) colorValue = en.Attributes.GetAsColor("color");
			if (!colorValue.IsEmpty)
				styleAttributes.Add(new OxW.Color { Val = colorValue.ToHexString() });

			colorValue = en.StyleAttributes.GetAsColor("background-color");
			if (!colorValue.IsEmpty)
			{
				// change the way the background-color renders. It now uses Shading instead of Highlight.
				// Changes brought by Wude on http://html2openxml.codeplex.com/discussions/277570
				styleAttributes.Add(new OxW.Shading { Val = OxW.ShadingPatternValues.Clear, Fill = colorValue.ToHexString() });
			}

			var decorations = Converter.ToTextDecoration(en.StyleAttributes["text-decoration"]);
			if ((decorations & TextDecoration.Underline) != 0)
			{
				styleAttributes.Add(new OxW.Underline { Val = OxW.UnderlineValues.Single });
			}
			if ((decorations & TextDecoration.LineThrough) != 0)
			{
				styleAttributes.Add(new OxW.Strike());
			}

			String[] classes = en.Attributes.GetAsClass();
			if (classes != null)
			{
				for (int i = 0; i < classes.Length; i++)
				{
					string className = documentStyle.GetStyle(classes[i], OxW.StyleValues.Character, ignoreCase: true);
					if (className != null) // only one Style can be applied in OpenXml and dealing with inheritance is out of scope
					{
						styleAttributes.Add(new OxW.RunStyle() { Val = className });
						break;
					}
				}
			}

			HtmlFont font = en.StyleAttributes.GetAsFont("font");
			if (!font.IsEmpty)
			{
				if (font.Style == FontStyle.Italic)
					styleAttributes.Add(new OxW.Italic());

				if (font.Weight == FontWeight.Bold || font.Weight == FontWeight.Bolder)
					styleAttributes.Add(new OxW.Bold());

				if (font.Variant == FontVariant.SmallCaps)
					styleAttributes.Add(new OxW.SmallCaps());

				if (font.Family != null)
					styleAttributes.Add(new OxW.RunFonts() { Ascii = font.Family, HighAnsi = font.Family });

				// size are half-point font size
                if (font.Size.IsFixed)
					styleAttributes.Add(new OxW.FontSize() { Val = (font.Size.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) });
			}
		}

        #endregion

        //____________________________________________________________________
        //
        // Properties

        /// <summary>
        /// Gets the default StyleId to apply on the any new runs.
        /// </summary>
        internal String DefaultRunStyle { get; set; }
	}
}
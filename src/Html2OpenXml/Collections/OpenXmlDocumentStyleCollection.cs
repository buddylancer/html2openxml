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
using OxW = DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Typed collection that holds the Style of a document and their name.
	/// OpenXml is case-sensitive but CSS is not. This collection handles both cases.
	/// </summary>
	sealed class OpenXmlDocumentStyleCollection : SortedList<String, OxW.Style>
	{
		public OpenXmlDocumentStyleCollection() : base(StringComparer.CurrentCulture)
		{
		}

		/// <summary>
		/// Gets the style associated with the specified name.
		/// </summary>
		/// <param name="name">The name whose style to get.</param>
		/// <param name="styleType">Specify the type of style seeked (Paragraph or Character).</param>
		/// <param name="style">When this method returns, the style associated with the specified name, if
		/// the key is found; otherwise, returns null. This parameter is passed uninitialized.</param>
		public bool TryGetValueIgnoreCase(String name, OxW.StyleValues styleType, out OxW.Style style)
		{
			// we'll use Binary Search algorithm because the collection is sorted (we inherits from SortedList)
			IList<String> keys = this.Keys;
			int low = 0, hi = keys.Count - 1, mid;

			while (low <= hi)
			{
				mid = low + (hi - low) / 2;
                // Do not use Ordinal for string comparison to avoid the '_' character not being considered (bug #13776 reported by giorand)
                int rc = String.Compare(name, keys[mid], StringComparison.CurrentCultureIgnoreCase);
				if (rc == 0)
				{
					style = this.Values[mid];
					OxW.Style firstFoundStyle = style;

					// we have found the named style but maybe the style doesn't match (Paragraph is not Character)
					for (int i = mid; i < keys.Count && !style.Type.Equals<OxW.StyleValues>(styleType); i++)
					{
						style = this.Values[i];
						if (!String.Equals(style.StyleName.Val, name, StringComparison.OrdinalIgnoreCase)) break;
					}

					if (!String.Equals(style.StyleName.Val, name, StringComparison.OrdinalIgnoreCase))
						style = firstFoundStyle;

					return style.Type.Equals<OxW.StyleValues>(styleType);
				}
				else if (rc < 0) hi = mid - 1;
				else low = mid + 1;
			}

			style = null;
			return false;
		}
	}
}
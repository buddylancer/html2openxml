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
using OxP = DocumentFormat.OpenXml.Packaging;
using OxW = DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	/// <summary>
	/// The event arguments used for a StyleMissing event.
	/// </summary>
	public class StyleEventArgs : EventArgs
	{
		internal StyleEventArgs(String styleId, OxP.MainDocumentPart mainPart, OxW.StyleValues type)
		{
			this.Name = styleId;
			this.StyleDefinitionsPart = mainPart.StyleDefinitionsPart;
			this.Type = type;
		}

		/// <summary>
		/// Gets the invariant name of the style.
		/// </summary>
		public String Name { get; private set; }

		/// <summary>
		/// Gets the styles definition part located inside MainDocumentPart.
		/// </summary>
		public OxP.StyleDefinitionsPart StyleDefinitionsPart { get; private set; }

		/// <summary>
		/// Gets the type of style seeked (character or paragraph).
		/// </summary>
		public OxW.StyleValues Type { get; private set; }
	}
}
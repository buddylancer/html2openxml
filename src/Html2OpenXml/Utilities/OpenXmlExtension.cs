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
using Ox = DocumentFormat.OpenXml;
using OxW = DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

	/// <summary>
	/// Helper class that provide some extension methods to OpenXml SDK.
	/// </summary>
    [System.Diagnostics.DebuggerStepThrough]
	static class OpenXmlExtension
    {
		public static bool HasChild<T>(this Ox.OpenXmlElement element) where T : Ox.OpenXmlElement
        {
            return element.GetFirstChild<T>() != null;
        }

		public static T GetLastChild<T>(this Ox.OpenXmlElement element) where T : Ox.OpenXmlElement
		{
			if (element == null) return null;

			for (int i = element.ChildElements.Count - 1; i >= 0; i--)
			{
				if (element.ChildElements[i] is T)
					return element.ChildElements[i] as T;
			}

			return null;
		}

		public static bool Equals<T>(this Ox.EnumValue<T> value, T comparand) where T : struct
        {
            return value != null && value.Value.Equals(comparand);
        }

		public static void InsertInProperties(this OxW.Paragraph p, Action<OxW.ParagraphProperties> @delegate)
		{
			OxW.ParagraphProperties prop = p.GetFirstChild<OxW.ParagraphProperties>();
			if (prop == null) p.PrependChild<OxW.ParagraphProperties>(prop = new OxW.ParagraphProperties());

			@delegate(prop);
		}

		public static void InsertInProperties(this OxW.Run r, Action<OxW.RunProperties> @delegate)
		{
			OxW.RunProperties prop = r.GetFirstChild<OxW.RunProperties>();
			if (prop == null) r.PrependChild<OxW.RunProperties>(prop = new OxW.RunProperties());

			@delegate(prop);
		}

		public static void InsertInDocProperties(this OxW.Drawing d, params Ox.OpenXmlElement[] newChildren)
		{
			wp.Inline inline = d.GetFirstChild<wp.Inline>();
			wp.DocProperties prop = inline.GetFirstChild<wp.DocProperties>();

			if (prop == null) inline.Append(prop = new wp.DocProperties());
			prop.Append(newChildren);
		}

		public static bool Compare(this OxW.PageSize pageSize, OxW.PageOrientationValues orientation)
        {
			OxW.PageOrientationValues pageOrientation;

            if (pageSize.Orient != null) pageOrientation = pageSize.Orient.Value;
			else if (pageSize.Width > pageSize.Height) pageOrientation = OxW.PageOrientationValues.Landscape;
			else pageOrientation = OxW.PageOrientationValues.Portrait;

            return pageOrientation == orientation;
        }

		// needed since December 2009 CTP refactoring, where casting is not anymore an option

		public static OxW.TableRowAlignmentValues ToTableRowAlignment(this OxW.JustificationValues val)
		{
			if (val == OxW.JustificationValues.Center) return OxW.TableRowAlignmentValues.Center;
			else if (val == OxW.JustificationValues.Right) return OxW.TableRowAlignmentValues.Right;
			else return OxW.TableRowAlignmentValues.Left;
		}
		public static OxW.JustificationValues ToJustification(this OxW.TableRowAlignmentValues val)
		{
			if (val == OxW.TableRowAlignmentValues.Left) return OxW.JustificationValues.Left;
			else if (val == OxW.TableRowAlignmentValues.Center) return OxW.JustificationValues.Center;
			else return OxW.JustificationValues.Right;
		}
    }
}
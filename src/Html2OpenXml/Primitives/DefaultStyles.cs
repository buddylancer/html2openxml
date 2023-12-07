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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Contains the default styles of Word elements
    /// </summary>
    public class DefaultStyles
    {
        #region Caption
        
        /// <summary>
        /// Default style for captions
        /// </summary>
        /// <value>Caption</value>
        public string CaptionStyle { get  { return mCaptionStyle; } set { mCaptionStyle = value; } }
		private string mCaptionStyle = "Caption";

        #endregion

        #region Endnotes

        /// <summary>
        /// Default style for new endnote texts
        /// </summary>
        /// <value>EndnoteText</value>
        public string EndnoteTextStyle { get { return mEndnoteTextStyle; } set { mEndnoteTextStyle = value; } }
        private string mEndnoteTextStyle  = "EndnoteText";

        /// <summary>
        /// Default style for new endnote references
        /// </summary>
        /// <value>EndnoteReference</value>
        public string EndnoteReferenceStyle { get { return mEndnoteReferenceStyle; } set { mEndnoteReferenceStyle = value; } }
        private string mEndnoteReferenceStyle = "EndnoteReference";

        #endregion

        #region Footnotes

        /// <summary>
        /// Default style for new footnote texts
        /// </summary>
        /// <value>FootnoteText</value>
        public string FootnoteTextStyle { get { return mFootnoteTextStyle; } set { mFootnoteTextStyle = value; } }
        private string mFootnoteTextStyle = "FootnoteText";

        /// <summary>
        /// Default style for new footnote references
        /// </summary>
        /// <value>FootnoteReference</value>
        public string FootnoteReferenceStyle { get { return mFootnoteReferenceStyle; } set { mFootnoteReferenceStyle = value; } }
        private string mFootnoteReferenceStyle = "FootnoteReference";

        #endregion

        #region Headings

        /// <summary>
        /// Default style for headings
        /// Appends the level at the end of the style name
        /// </summary>
        /// <value>Heading</value>
        public string HeadingStyle { get { return mHeadingStyle; } set { mHeadingStyle = value; } }
        private string mHeadingStyle = "Heading";

        #endregion

        #region Hyperlink

        /// <summary>
        /// Default style for hyperlinks
        /// </summary>
        /// <value>Hyperlink</value>
        public string HyperlinkStyle { get { return mHyperlinkStyle; } set { mHyperlinkStyle = value; } }
        private string mHyperlinkStyle = "Hyperlink";

        #endregion

        #region Lists

        /// <summary>
        /// Default style for list paragraphs
        /// </summary>
        /// <value>ListParagraph</value>
        public string ListParagraphStyle { get { return mListParagraphStyle; } set { mListParagraphStyle = value; } }
        private string mListParagraphStyle = "ListParagraph";

        #endregion

        #region Paragraph

        /// <summary>
        /// Default style for paragraphs
        /// </summary>
        /// <value>null</value>
        public string ParagraphStyle { get; set; }

        #endregion

        #region Pre

        /// <summary>
        /// Default style for the &lt;pre&gt; table
        /// </summary>
        /// <value>TableGrid</value>
        public string PreTableStyle { get { return mPreTableStyle; } set { mPreTableStyle = value; } }
        private string mPreTableStyle = "TableGrid";

        #endregion

        #region Quotes

        /// <summary>
        /// Default style for quotes
        /// </summary>
        /// <value>Quote</value>
        public string QuoteStyle { get { return mQuoteStyle; } set { mQuoteStyle = value; } }
        private string mQuoteStyle = "Quote";

        /// <summary>
        /// Default style for intense quotes
        /// </summary>
        /// <value>IntenseQuote</value>
		public string IntenseQuoteStyle { get { return mIntenseQuoteStyle; } set { mIntenseQuoteStyle = value; } }
        private string mIntenseQuoteStyle = "IntenseQuote";

        #endregion

        #region Table

        /// <summary>
        /// Default style for tables
        /// </summary>
        /// <value>TableGrid</value>
        public string TableStyle { get { return mTableStyle; } set { mTableStyle = value; } }
        private string mTableStyle = "TableGrid";

        #endregion
    }
}
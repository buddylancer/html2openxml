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
using System.IO;
using System.Threading;
using OxP = DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.IO
{
    /// <summary>
    /// Download and provison the metadata of a requested image.
    /// </summary>
    sealed class ImagePrefetcher
    {
        // Map extension to ImagePartType
		private static readonly Dictionary<string, OxP.ImagePartType> knownExtensions = new Dictionary<string, OxP.ImagePartType>(StringComparer.OrdinalIgnoreCase) {
            { ".gif", OxP.ImagePartType.Gif },
            { ".bmp", OxP.ImagePartType.Bmp },
            { ".emf", OxP.ImagePartType.Emf },
            { ".ico", OxP.ImagePartType.Icon },
            { ".jpeg", OxP.ImagePartType.Jpeg },
            { ".jpg", OxP.ImagePartType.Jpeg },
            { ".jpe", OxP.ImagePartType.Jpeg },
            { ".pcx", OxP.ImagePartType.Pcx },
            { ".png", OxP.ImagePartType.Png },
            { ".tiff", OxP.ImagePartType.Tiff },
            { ".wmf", OxP.ImagePartType.Wmf }
        };
		private readonly OxP.MainDocumentPart mainPart;
        private readonly IWebRequest resourceLoader;
        private HtmlImageInfoCollection prefetchedImages;


		public ImagePrefetcher(OxP.MainDocumentPart mainPart, IWebRequest resourceLoader)
        {
            this.mainPart = mainPart;
            this.resourceLoader = resourceLoader;
            this.prefetchedImages = new HtmlImageInfoCollection();
        }

        //____________________________________________________________________
        //
        // Public Functionality

        /// <summary>
        /// Download the remote or local image located at the specified url.
        /// </summary>
        public HtmlImageInfo Download(string imageUri)
        {
            if (prefetchedImages.Contains(imageUri))
                return prefetchedImages[imageUri];

            if (DataUri.IsWellFormed(imageUri)) // data inline, encoded in base64
            {
                return ReadDataUri(imageUri);
            }

            return DownloadRemoteImage(imageUri);
        }

        /// <summary>
        /// Download the image and try to find its format type.
        /// </summary>
        private HtmlImageInfo DownloadRemoteImage(string src)
        {
            Uri imageUri = new Uri(src, UriKind.RelativeOrAbsolute);
            Resource response;
            try
            {
                response = resourceLoader.FetchAsync(imageUri, CancellationToken.None).Result;
                if (response.Content == null)
                    return null;
            }
            catch (Exception exc)
            {
                if (Logging.On) Logging.PrintError(String.Format("Error fetching image from url: {0}", src), exc);
                return null;
            }

            HtmlImageInfo info = new HtmlImageInfo() { Source = src };
			OxP.ImagePartType type;
            using (response)
            {
                // For requested url with no filename, we need to read the media mime type if provided
                string mime; response.Headers.TryGetValue("Content-Type", out mime);
                if (!TryInspectMimeType(mime, out type)
                    && !TryGuessTypeFromUri(imageUri, out type)
                    && !TryGuessTypeFromStream(response.Content, out type))
                {
                    return null;
                }

                var ipart = mainPart.AddImagePart(type);
                using (var outputStream = ipart.GetStream(FileMode.Create))
                {
                    response.Content.CopyTo(outputStream);

                    outputStream.Seek(0L, SeekOrigin.Begin);
                    info.Size = GetImageSize(outputStream);
                }

                info.ImagePartId = mainPart.GetIdOfPart(ipart);
                return info;
            }
        }

        /// <summary>
        /// Parse the Data inline image.
        /// </summary>
        private HtmlImageInfo ReadDataUri(string src)
        {
			DataUri dataUri;
            if (DataUri.TryCreate(src, out dataUri))
            {
                Size size;
				OxP.ImagePartType type; knownContentType.TryGetValue(dataUri.Mime, out type);
                var ipart = mainPart.AddImagePart(type);
                using (var outputStream = ipart.GetStream(FileMode.Create))
                {
                    outputStream.Write(dataUri.Data, 0, dataUri.Data.Length);

                    outputStream.Seek(0L, SeekOrigin.Begin);
                    size = GetImageSize(outputStream);
                }

                return new HtmlImageInfo() {
                    Source = src,
                    ImagePartId = mainPart.GetIdOfPart(ipart),
                    Size = size
                };
            }

            return null;
        }

        //____________________________________________________________________
        //
        // Private Implementation

        // http://stackoverflow.com/questions/58510/using-net-how-can-you-find-the-mime-type-of-a-file-based-on-the-file-signature
		private static Dictionary<string, OxP.ImagePartType> knownContentType = new Dictionary<String, OxP.ImagePartType>(StringComparer.OrdinalIgnoreCase) {
            { "image/gif", OxP.ImagePartType.Gif },
            { "image/pjpeg", OxP.ImagePartType.Jpeg },
            { "image/jpg", OxP.ImagePartType.Jpeg },
            { "image/jpeg", OxP.ImagePartType.Jpeg },
            { "image/x-png", OxP.ImagePartType.Png },
            { "image/png", OxP.ImagePartType.Png },
            { "image/tiff", OxP.ImagePartType.Tiff },
            { "image/vnd.microsoft.icon", OxP.ImagePartType.Icon },
            // these icons mime type are wrong but we should nevertheless take care (http://en.wikipedia.org/wiki/ICO_%28file_format%29#MIME_type)
            { "image/x-icon", OxP.ImagePartType.Icon },
            { "image/icon", OxP.ImagePartType.Icon },
            { "image/ico", OxP.ImagePartType.Icon },
            { "text/ico", OxP.ImagePartType.Icon },
            { "text/application-ico", OxP.ImagePartType.Icon },
            { "image/bmp", OxP.ImagePartType.Bmp }
        };

        /// <summary>
        /// Inspect the response headers of a web request and decode the mime type if provided
        /// </summary>
        /// <returns>Returns the extension of the image if provideds.</returns>
		private static bool TryInspectMimeType(string contentType, out OxP.ImagePartType type)
        {
            // can be null when the protocol used doesn't allow response headers
            if (contentType != null &&
                knownContentType.TryGetValue(contentType, out type))
                return true;

			type = default(OxP.ImagePartType);
            return false;
        }

        /// <summary>
        /// Gets the OpenXml ImagePartType associated to an image.
        /// </summary>
		private static bool TryGuessTypeFromUri(Uri uri, out OxP.ImagePartType type)
        {
            string extension = Path.GetExtension(uri.IsAbsoluteUri ? uri.Segments[uri.Segments.Length - 1] : uri.OriginalString);
            if (knownExtensions.TryGetValue(extension, out type)) return true;

            // extension not recognized, try with checking the query string. Expecting to resolve something like:
            // ./image.axd?picture=img1.jpg
            extension = Path.GetExtension(uri.IsAbsoluteUri ? uri.AbsoluteUri : uri.ToString());
            if (knownExtensions.TryGetValue(extension, out type)) return true;

            return false;
        }

        /// <summary>
        /// Gets the OpenXml ImagePartType associated to an image.
        /// </summary>
		private static bool TryGuessTypeFromStream(Stream stream, out OxP.ImagePartType type)
        {
			ImageHeader.FileType guessType;
            if (ImageHeader.TryDetectFileType(stream, out guessType))
            {
                switch (guessType)
                {
					case ImageHeader.FileType.Bitmap: type = OxP.ImagePartType.Bmp; return true;
					case ImageHeader.FileType.Emf: type = OxP.ImagePartType.Emf; return true;
					case ImageHeader.FileType.Gif: type = OxP.ImagePartType.Gif; return true;
					case ImageHeader.FileType.Jpeg: type = OxP.ImagePartType.Jpeg; return true;
					case ImageHeader.FileType.Png: type = OxP.ImagePartType.Png; return true;
                }
            }
			type = OxP.ImagePartType.Bmp;
            return false;
        }

        /// <summary>
        /// Loads an image from a stream and grab its size.
        /// </summary>
        private static Size GetImageSize(Stream imageStream)
        {
            // Read only the size of the image
            try
            {
                return ImageHeader.GetDimensions(imageStream);
            }
            catch (ArgumentException)
            {
                return Size.Empty;
            }
        }
    }
}
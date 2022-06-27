using System;
using System.IO;
using Eto.Drawing;

namespace Elistia.DotNetRtfWriter
{
    /// <summary>
    /// 
    /// </summary>
    public static class RtfExtensions
    {
        public static RtfColor ToRtfColor(this Color color)
        {
            return new RtfColor((byte)color.Rb, (byte)color.Gb, (byte)color.Bb);
        }

        public static ColorDescriptor createColor(this RtfDocument document, Color color)
        {
            var rtfColor = color.ToRtfColor();
            return document.createColor(rtfColor);
        }

        // RtfImageType { Jpg = 1, Gif, Png }
        private static readonly ImageFormat[] EtoImageFormats = new ImageFormat[] { ImageFormat.Bitmap, ImageFormat.Jpeg, ImageFormat.Gif, ImageFormat.Png };

        public static RtfImage FromFile(string fileName, RtfImageType type)
        {
            var image = new Bitmap(fileName);
            float width = image.Width; //(image.Width / image.HorizontalResolution) * 72;
            float height = image.Height; //(image.Height / image.VerticalResolution) * 72;

            byte[] imgBytes;
            using (MemoryStream mStream = new MemoryStream()) {
                image.Save(mStream, EtoImageFormats[(int)type]);
                imgBytes = mStream.ToArray();
            }

            return new RtfImage(type, width, height, imgBytes, fileName);
        }

        /*public static RtfImage FromStream(MemoryStream stream)
        {
            byte[] imgBytes = stream.ToArray();

            var image = new Bitmap(stream);
            float width = image.Width; //(image.Width / image.HorizontalResolution) * 72;
            float height = image.Height; //(image.Height / image.VerticalResolution) * 72;

            RtfImageType imgType;
            if (image.RawFormat.Equals(ImageFormat.Png))
                imgType = RtfImageType.Png;
            else if (image.RawFormat.Equals(ImageFormat.Jpeg))
                imgType = RtfImageType.Jpg;
            else if (image.RawFormat.Equals(ImageFormat.Gif))
                imgType = RtfImageType.Gif;
            else
                throw new Exception("Image format is not supported: " + image.RawFormat.ToString());

            return new RtfImage(imgType, width, height, imgBytes, String.Empty);
        }*/


        /// <summary>
        /// Add an image to this container from a file with filetype provided.
        /// </summary>
        /// <param name="imgFname">Filename of the image.</param>
        /// <param name="imgType">File type of the image.</param>
        /// <returns>Image being added.</returns>
        public static RtfImage addImage(this RtfBlockList blockList, string imgFname, RtfImageType imgType)
        {
            if (!blockList._allowImage) {
                throw new Exception("Image is not allowed.");
            }
            RtfImage block = FromFile(imgFname, imgType);
            blockList.addBlock(block);
            return block;
        }

        /// <summary>
        /// Add an image to this container from a file. Will autodetect format from extension.
        /// </summary>
        /// <param name="imgFname">Filename of the image.</param>
        /// <returns>Image being added.</returns>
        public static RtfImage addImage(this RtfBlockList blockList, string imgFname)
        {
            int dot = imgFname.LastIndexOf(".");
            if (dot < 0) {
                throw new Exception("Cannot determine image type from the filename extension: " + imgFname);
            }

            string ext = imgFname.Substring(dot + 1).ToLower();
            switch (ext) {
                case "jpg":
                case "jpeg":
                    return addImage(blockList, imgFname, RtfImageType.Jpg);
                case "gif":
                    return addImage(blockList, imgFname, RtfImageType.Gif);
                case "png":
                    return addImage(blockList, imgFname, RtfImageType.Png);
                default:
                    throw new Exception("Cannot determine image type from the filename extension: " + imgFname);
            }
        }

        /// <summary>
        /// Add an image to this container from a stream.
        /// </summary>
        /// <param name="imageStream">MemoryStream containing image.</param>
        /// <returns>Image being added.</returns>
        /*public static RtfImage addImage(this RtfBlockList blockList, MemoryStream imageStream)
        {
            if (!blockList._allowImage) {
                throw new Exception("Image is not allowed.");
            }
            RtfImage block = FromStream(imageStream);
            blockList.addBlock(block);
            return block;
        }*/
    }
}

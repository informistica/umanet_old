/*  WBRESIZE 0.41 [ASP.net version]
	by imente 27/05/2008
	https://www.imente.org/short/wbresize
	
	**********************************************************************************
	
	WBRESIZE è rilasciato sotto licenza "Creative Commons Attribuzione 2.5 Italia"
	>>> https://creativecommons.org/licenses/by/2.5/it/
	>>> https://creativecommons.org/licenses/by/2.5/it/deed.en
	>>> E' possibile utilizzare, modificare e ridistribuire WBRESIZE
	>>> liberamente anche per fini commerciali a patto che venga riconosciuta la paternità
	 * IL CODICE IN OGGETTO VIENE FORNITO "COSÌ COM'È" AI SENSI DELLA
	 * PRESENTE LICENZA, SENZA ALCUNA GARANZIA, ESPRESSA O IMPLICITA, 
	 * COMPRESE, TRA LE ALTRE, LA GARANZIA DI ASSENZA DI DIFETTI,
	 * DI COMMERCIABILITÀ, DI IDONEITÀ PER UN FINE PARTICOLARE E DI 
	 * NON VIOLAZIONE DI DIRITTI ALTRUI. IL RISCHIO RELATIVO ALLA
	 * QUALITÀ E AL RENDIMENTO DEL CODICE IN OGGETTO È A CARICO 
	 * DELL'UTILIZZATORE.
	
	**********************************************************************************
	
	WBRESIZE è una copia esatta di "ic.image.resizer", classe sviluppata in seno al
	progetto "imenteCommon" (IC) e più precisa mente dal namespace "ic.image".
	
	**********************************************************************************
	
	QUESTA E' LA VERSIONE PER ASP.NET
	se cerchi la versione per ASP 3 vieni sul sito
	
	**********************************************************************************
*/


public class wbResize
    {

        public int thumbWidth = 0;
        public int thumbHeight = 0;
        private string FileType = "jpg"; //jpg,gif,bmp,tiff,png

        private System.Drawing.Image rImage = null;
        private int originalWidth = 0;
        private int originalHeight = 0;
        int JpegQuality = 100;  //0-100
        bool TiffCompress = true;
        bool HighQuality = true;
        InterpolationMode Filter = InterpolationMode.Low;
        byte GifPalette = 255;  //9-255
        byte GifDepth = 8;      //0,255,1,2,4,8
        long ByteSize;

        /****************************************************************************************/
        /****************************************************************************************/

        public void LoadImage(string path)
        {
            if (!File.Exists(path)) { throw new FileNotFoundException(); }
            try { LoadImage(System.Drawing.Image.FromFile(path)); }
            catch { throw new FileNotFoundException(); }
        }
        public void LoadImage(Stream stream)
        {
            try { LoadImage(System.Drawing.Image.FromStream(stream)); }
            catch { throw new FileNotFoundException(); }
        }
        public void LoadImage(System.Drawing.Image image)
        {
            rImage = new Bitmap(image);
            originalWidth = image.Width;
            originalHeight = image.Height;
			image.Dispose();
			image = null;
        }
        public void UnloadImage()
        {
            if (rImage!=null) rImage.Dispose();
            rImage = null;
        }

        /****************************************************************************************/
        /****************************************************************************************/

        private double scaleProportional(double originalA, double originalB, double requestedA)
        {
            //scala una delle due dimensioni eseguendo una proporzione
            return Math.Ceiling(requestedA * originalB / originalA);
        }

        public void SetThumbSize(double requestedWidth, double requestedHeight)
        { SetThumbSize(requestedWidth, requestedHeight, 0, false, false); }
        public void SetThumbSize(double requestedWidth, double requestedHeight, bool forceDimension)
        { SetThumbSize(requestedWidth, requestedHeight, 0, forceDimension, false); }
        public void SetThumbSize(int scaleFactor)
        { SetThumbSize(0, 0, scaleFactor, false, false); }
        public void SetThumbSize(double requestedWidth, double requestedHeight, int scaleFactor, bool forceDimension, bool allowSuperSize)
        {
            //calcola la dimensione dell'immagine ridotta
            //[restituisce 0,0 in caso di parametri insufficienti]

            int thumbwidth = 0, thumbheight = 0;

            bool executeResize = true;

            if (!allowSuperSize && scaleFactor > 100) scaleFactor = 100;
            if (scaleFactor < 0) scaleFactor = 0;

            //se una dimensione è più grande dell'immagine e l'altra no, 
            //la diminuisce fino a renderla pari all'immagine
            if (!(!forceDimension && requestedHeight != 0 && requestedWidth != 0) && !allowSuperSize)
            {
                if (originalWidth > requestedWidth && originalHeight < requestedHeight) requestedHeight = originalHeight;
                if (originalWidth < requestedWidth && originalHeight > requestedHeight) requestedWidth = originalWidth;
                executeResize = (originalWidth >= requestedWidth && originalHeight >= requestedHeight && scaleFactor >= 0 && scaleFactor < 100);
            }

            if (executeResize)
            {
                //l'immagine è da ridimensionare: calcola le dimensioni della miniatura
                if (scaleFactor == 0)
                {
                    if (requestedWidth == 0 && requestedHeight != 0)
                    {
                        // forzata requestedHeight
                        thumbwidth = Convert.ToInt32(scaleProportional(originalHeight, requestedHeight, originalWidth));
                        thumbheight = Convert.ToInt32(requestedHeight);
                    }
                    else if (requestedHeight == 0 && requestedWidth != 0)
                    {
                        // forzata requestedWidth
                        thumbwidth = Convert.ToInt32(requestedWidth);
                        thumbheight = Convert.ToInt32(scaleProportional(originalWidth, requestedWidth, originalHeight));
                    }
                    else if (requestedHeight != 0 && requestedWidth != 0)
                    {
                        if (forceDimension)
                        {
                            // forzata sia requestedWidth che requestedHeight
                            thumbwidth = Convert.ToInt32(requestedWidth);
                            thumbheight = Convert.ToInt32(requestedHeight);
                        }
                        else
                        {
                            double newWidth, newHeight;
                            // scala le dimensioni in proporzione per mantenerle
                            //nel riquadro requestedWidth/requestedHeight
                            newHeight = requestedHeight;
                            newWidth = scaleProportional(originalHeight, newHeight, originalWidth);
                            if (newWidth > requestedWidth)
                            {
                                newHeight = Convert.ToInt32(scaleProportional(newWidth, requestedWidth, newHeight));
                                newWidth = Convert.ToInt32(requestedWidth);
                            }
                            thumbheight = Convert.ToInt32(newHeight);
                            thumbwidth = Convert.ToInt32(newWidth);
                            if (thumbwidth > originalWidth && thumbheight > originalHeight && !allowSuperSize)
                            {
                                thumbheight = originalHeight;
                                thumbwidth = originalWidth;
                            }
                        }
                    }
                    else //requestedHeight == 0 && requestedWidth == 0
                    {
                        thumbwidth = 0;
                        thumbheight = 0;
                    }
                }
                else
                {
                    // scalata: percentuale delle dimensioni
                    thumbwidth = (originalWidth * scaleFactor / 100);
                    thumbheight = (originalHeight * scaleFactor / 100);
                }
            }
            else
            {
                // sovradimensionamento (si impostano le nuove dimensioni a 0 per evitare il ridimensionamento)
                thumbwidth = originalWidth;
                thumbheight = originalHeight;
            }
            thumbWidth = thumbwidth;
            thumbHeight = thumbheight;
        }

        /****************************************************************************************/
        /****************************************************************************************/

        public string Report() { return Report(""); }
        public string Report(string key)
        {
            string output = "";
            string depth = "";
            switch (key.ToLower())
            {
                case "width": output = thumbWidth.ToString(); break;
                case "height": output = thumbHeight.ToString(); break;
                case "filetype": output = FileType; break;
                case "jpegquality": output = JpegQuality.ToString(); break;
                case "gifpalette": output = GifPalette.ToString(); break;
                case "gifdepth": depth = GifDepth.ToString() + " Bit"; if (GifDepth == 255) depth = "Greyscale"; output = depth; break;
                case "tiffcompress": output = TiffCompress.ToString(); break;
                case "filter": output = Filter.ToString(); break;
                case "size": output = Convert.ToString(Math.Round((double)ByteSize / 1024,2)) + " Kb"; break;
                default:
                    output = output + FileType + " Image - " + thumbWidth.ToString() + "x" + thumbHeight.ToString();
                    switch (FileType)
                    {
                        case "jpg": output = output + " - Quality " + JpegQuality.ToString() + "%"; break;
                        case "gif":
                            if (GifDepth != 0) {
                                output = output + " - Palette " + GifPalette.ToString();
                                depth = GifDepth.ToString() + " Bit";
                                if (GifDepth == 255) depth = "Greyscale";
                                output = output + " - Color " + depth;
                            }
                            break;
                        case "tiff": output = output + " - Compress " + TiffCompress.ToString(); break;
                    }
                    output = output + " - Filter " + Filter.ToString();
                    output = output + " - Size " + Convert.ToString(Math.Round((double)ByteSize / 1024,2)) + " Kb";
                    break;
            }
            return output;
        }

        /****************************************************************************************/
        /****************************************************************************************/

        public void SetFileType(string filetype)
        {
            FileType = filetype.ToLower();
            //normalizza il filetype ai tipi supportati
            if (FileType != "gif" && FileType != "png" && FileType != "bmp" && FileType != "tiff") FileType = "jpg";
        }

        public void SetOptionJpegQuality(int jpegquality)
        {
            JpegQuality = jpegquality;
            if (JpegQuality > 100) JpegQuality = 100;
            if (JpegQuality < 0) JpegQuality = 0;
        }

        public void SetOptionTiffCompress(bool tiffcompress)
        {
            TiffCompress = tiffcompress;
        }

        public void SetOptionHighQuality(bool highquality)
        {
            HighQuality = highquality;
        }

        public void SetOptionFilter(InterpolationMode filter)
        {
            Filter = filter;
        }

        public void SetOptionGifPalette(byte gifpalette)
        {
            GifPalette = gifpalette;
            if (GifPalette < 2) GifPalette = 2;
        }        

        public void SetOptionGifDepth(byte gifdepth)
        {
            GifDepth = gifdepth;
            /* VERSIONE PER ASP
            if (GifDepth != 0 && GifDepth != 1 && GifDepth != 2 && GifDepth != 4 && GifDepth != 8) GifDepth = 4;*/
            /* VERSIONE GENERALE */
            if (GifDepth != 0 && GifDepth != 1 && GifDepth != 2 && GifDepth != 4 && GifDepth != 8 && GifDepth != 255) GifDepth = 4;
        }

        public void SetOptions(string filetype, int compression, bool highquality, InterpolationMode filter, byte gifpalette, byte gifdepth)
        {
            SetFileType(filetype);
            SetOptionJpegQuality(compression);
            SetOptionHighQuality(highquality);
            SetOptionFilter(filter);
            SetOptionGifPalette(gifpalette);
            SetOptionGifDepth(gifdepth);
        }

        /****************************************************************************************/
        /****************************************************************************************/

        public void ResizeAndSave(string fileoutput)
        {
            //salva su file un'immagine ridimensionandola
            //[in caso di errore restituisce false]

            try
            {
                System.Drawing.Image thumbnail = System.Drawing.Image.FromStream(Resize());
                thumbnail.Save(fileoutput);
                thumbnail.Dispose();
                thumbnail = null;
            }
            catch
            {
                throw new System.ApplicationException("Unable to save the image");
            }
        }

        public MemoryStream Resize()
        {
            //salva su memorystream un'immagine ridimensionandola

            if (rImage == null) throw new System.ApplicationException("Image not loaded!");

            MemoryStream memorystream = new MemoryStream();
            ByteSize = 0;
            try
            {
                System.Drawing.Image thumbnail = new Bitmap(thumbWidth, thumbHeight);
                System.Drawing.Graphics graphic = System.Drawing.Graphics.FromImage(thumbnail);
                if (HighQuality)
                {
                    graphic.InterpolationMode = Filter;
                    graphic.SmoothingMode = SmoothingMode.HighQuality;
                    graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphic.CompositingQuality = CompositingQuality.HighQuality;
                }
                else
                {
                    graphic.InterpolationMode = InterpolationMode.Low;
                    graphic.SmoothingMode = SmoothingMode.HighSpeed;
                    graphic.PixelOffsetMode = PixelOffsetMode.HighSpeed;
                    graphic.CompositingQuality = CompositingQuality.HighSpeed;
                }
                graphic.DrawImage(rImage, 0, 0, thumbWidth, thumbHeight);


                /* EVENTUALI OPERAZIONI SULL'IMMAGINE QUI... */

                EncoderParameters encoderParameters;
                encoderParameters = new EncoderParameters(1);
                ImageCodecInfo[] info = ImageCodecInfo.GetImageEncoders();

                switch (FileType)
                {
                    case "gif": //       info[2]
                        if (GifDepth!=0)
                        {
                            /* VERSIONE PER ASP
                            Quantizer quantizer = new OctreeQuantizer(GifPalette, GifDepth);
                            using (Bitmap quantized = quantizer.Quantize(thumbnail))
                            { quantized.Save(memorystream, ImageFormat.Gif); }*/
                            /* VERSIONE GENERALE */
                            Quantizer quantizer;
                            if (GifDepth == 255) quantizer = new GrayscaleQuantizer();
                            else quantizer = new OctreeQuantizer(GifPalette, GifDepth);
                            using (Bitmap quantized = quantizer.Quantize(thumbnail))
                                { quantized.Save(memorystream, ImageFormat.Gif); }
                        }
                        else thumbnail.Save(memorystream, System.Drawing.Imaging.ImageFormat.Gif);
                        break;
                    case "png": //       info[4]
                        thumbnail.Save(memorystream, System.Drawing.Imaging.ImageFormat.Png);
                        break;
                    case "bmp": //       info[0]
                        thumbnail.Save(memorystream, System.Drawing.Imaging.ImageFormat.Bmp);
                        break;
                    case "tiff": //      info[3]
                        encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, TiffCompress ? (long)EncoderValue.CompressionLZW : (long)EncoderValue.CompressionNone);
                        thumbnail.Save(memorystream, info[3], encoderParameters);
                        break;
                    default: //jpeg      info[1]
                        encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, JpegQuality);          
                        thumbnail.Save(memorystream, info[1], encoderParameters);
                        break;
                }
                ByteSize = memorystream.Length;
                graphic.Dispose();
                graphic = null;
                thumbnail.Dispose();
                thumbnail = null;
            }
            catch
            {
                throw new System.ApplicationException("Unable to resize the image");
            }
            return memorystream;
        }
    }
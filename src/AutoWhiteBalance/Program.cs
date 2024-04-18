using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices.ComTypes;


namespace AutoWhiteBalance
{
    internal class Program
    {
        static double[] WhitePoint = new double[2];  //白色点のxy座標

        static double[,] RGB2XYZ = {
            {0.4124, 0.3576, 0.1805},
            {0.2126, 0.7152, 0.0722},
            {0.0193, 0.1192, 0.9505}
            };
        static double[,] XYZ2RGB;

        static void Main(string[] args)
        {
            string whitePoint = "0.3127,0.3290"; // デフォルトの白色点座標
            double gamma = 2.2; // デフォルトのガンマ値
            string imagePath = null;
            string outputFilename = null;

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--whitepoint":
                    case "-w":
                        if (i + 1 < args.Length)
                        {
                            whitePoint = args[i + 1];
                            i++; // 次の引数はこのオプションの一部なのでスキップ
                        }
                        else
                        {
                            Console.WriteLine("Error: No coordinates provided for white point.");
                            return;
                        }
                        break;
                    case "--gamma":
                    case "-g":
                        if (i + 1 < args.Length && double.TryParse(args[i + 1], out gamma))
                        {
                            i++; // 次の引数はこのオプションの一部なのでスキップ
                        }
                        else
                        {
                            Console.WriteLine("Error: Invalid or no gamma value provided.");
                            return;
                        }
                        break;
                    case "--output":
                    case "-o":
                        if (i + 1 < args.Length)
                        {
                            outputFilename = args[i + 1];
                            i++;
                        }
                        else
                        {
                            Console.WriteLine("Error: No output filename provided.");
                            return;
                        }
                        break;
                    default:
                        if (imagePath == null)
                        {
                            imagePath = args[i];
                        }
                        else
                        {
                            Console.WriteLine("Error: Multiple image paths provided. Please specify only one image path.");
                            return;
                        }
                        break;
                }
            }

            if (string.IsNullOrEmpty(imagePath))
            {
                ShowUsage();
                return;
            }

            if (!File.Exists(imagePath))
            {
                Console.WriteLine("Error: The file specified does not exist.");
                return;
            }

            if (string.IsNullOrEmpty(outputFilename))
            {
                outputFilename = Path.GetFileNameWithoutExtension(imagePath) + "_awb" + Path.GetExtension(imagePath);
            }

            var coordinates = whitePoint.Split(',');
            if (
                !(coordinates.Length == 2 &&
                double.TryParse(coordinates[0], out WhitePoint[0]) &&
                double.TryParse(coordinates[1], out WhitePoint[1])))
            {
                Console.WriteLine("Error: Invalid white point coordinates. Please specify them as x,y.");
            }

            Console.WriteLine($"Processing image at {imagePath} with white point coordinates ({WhitePoint[0]}, {WhitePoint[1]}).");

            //ここからAWB処理の開始
            //RGB2XYZの逆行列を求めます
            XYZ2RGB = InvertMatrix(RGB2XYZ);

            //画像をBitmapに取り込む
            Bitmap bmp = null;
            try
            {
                using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        fs.CopyTo(ms);
                        using (Image image = Image.FromStream(ms))
                        {
                            bmp = new Bitmap(image);
                        }
                    }
                }
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine("Error: The file format is not supported.");
                bmp?.Dispose();
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                bmp?.Dispose();
                return;
            }

            //画像サイズを取得
            int width = bmp.Width;
            int height = bmp.Height;

            // Bitmapをロックし、BitmapDataオブジェクトを取得
            Rectangle rect = new Rectangle(0, 0, width, height);
            BitmapData bmpData = bmp.LockBits(rect, ImageLockMode.ReadWrite, bmp.PixelFormat);

            // BitmapDataからのピクセルデータのサイズを計算
            int byteCount = bmpData.Stride * height;
            byte[] byteArray = new byte[byteCount];

            // BitmapDataからbyte配列にピクセルデータをコピー
            IntPtr ptr = bmpData.Scan0;
            Marshal.Copy(ptr, byteArray, 0, byteCount);

            // Bitmapのロックを解除
            bmp.UnlockBits(bmpData);
            bmp?.Dispose();

            //XYZ刺激値配列
            double[][,] xyz = new double[3][,];
            xyz[0] = new double[height, width];
            xyz[1] = new double[height, width];
            xyz[2] = new double[height, width];
            byte r, g, b;
            double lr, lg, lb;
            double sx, sy, sz;
            int onePixByte = byteCount / (height * width);

            double totalx = 0.0, totaly = 0.0;
            long countx = 0, county = 0;
            //X,Y,Z刺激値にしたあと、L,x,y(輝度と色度)にして、x,y色度座標の平均を出します
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    //RGB取得
                    r = byteArray[onePixByte * (y * width + x) + 2];
                    g = byteArray[onePixByte * (y * width + x) + 1];
                    b = byteArray[onePixByte * (y * width + x)];

                    //線形RGB計算
                    lr = Math.Pow(r / 255.0, gamma);
                    lg = Math.Pow(g / 255.0, gamma);
                    lb = Math.Pow(b / 255.0, gamma);

                    //XYZ刺激値計算
                    sx = RGB2XYZ[0, 0] * lr + RGB2XYZ[0, 1] * lg + RGB2XYZ[0, 2] * lb;
                    sy = RGB2XYZ[1, 0] * lr + RGB2XYZ[1, 1] * lg + RGB2XYZ[1, 2] * lb;
                    sz = RGB2XYZ[2, 0] * lr + RGB2XYZ[2, 1] * lg + RGB2XYZ[2, 2] * lb;
                    xyz[0][y, x] = sx;
                    xyz[1][y, x] = sy;
                    xyz[2][y, x] = sz;

                    //合計値と平均値を出す個数の計算
                    if ((sx + sy + sz) != 0.0)
                    {
                        totalx += sx / (sx + sy + sz);
                        totaly += sy / (sx + sy + sz);
                        countx++;
                        county++;
                    }
                }
            }

            //色度x,yの平均値を求めます
            double avex = totalx / countx;
            double avey = totaly / county;

            //平均値と白点D65差分をオフセットとして求めます
            double offsetx = avex - WhitePoint[0];
            double offsety = avey - WhitePoint[1];

            double kx, ky;
            double[] lrgb = new double[3];
            double[][,] rgb = new double[3][,];
            rgb[0] = new double[height, width];
            rgb[1] = new double[height, width];
            rgb[2] = new double[height, width];

            for (int y = 0, index = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    //XYZ刺激値から色度(x,y)を求めます
                    kx = xyz[0][y, x] / (xyz[0][y, x] + xyz[1][y, x] + xyz[2][y, x]);
                    ky = xyz[1][y, x] / (xyz[0][y, x] + xyz[1][y, x] + xyz[2][y, x]);
                    //xyz[1][y, x] *= 3.0; //輝度を変えたいときは、コメントを有効にして、ここの係数を変える

                    //白色点座標からのオフセット分を引きます
                    kx -= offsetx;
                    ky -= offsety;

                    //オフセットした色度(x,y)を使って刺激値XとZを再計算します
                    xyz[0][y, x] = kx / ky * xyz[1][y, x];
                    xyz[2][y, x] = ((1.0 - kx - ky) / ky​) * xyz[1][y, x];

                    //刺激値XYZから線形RGBを求めます
                    lrgb[0] = XYZ2RGB[0, 0] * xyz[0][y, x] + XYZ2RGB[0, 1] * xyz[1][y, x] + XYZ2RGB[0, 2] * xyz[2][y, x];
                    lrgb[1] = XYZ2RGB[1, 0] * xyz[0][y, x] + XYZ2RGB[1, 1] * xyz[1][y, x] + XYZ2RGB[1, 2] * xyz[2][y, x];
                    lrgb[2] = XYZ2RGB[2, 0] * xyz[0][y, x] + XYZ2RGB[2, 1] * xyz[1][y, x] + XYZ2RGB[2, 2] * xyz[2][y, x];

                    lrgb[0] = Math.Min(1.0, lrgb[0]);
                    lrgb[1] = Math.Min(1.0, lrgb[1]);
                    lrgb[2] = Math.Min(1.0, lrgb[2]);

                    lrgb[0] = Math.Max(0.0, lrgb[0]);
                    lrgb[1] = Math.Max(0.0, lrgb[1]);
                    lrgb[2] = Math.Max(0.0, lrgb[2]);

                    //線形RGBをγを使って0～255のRGBに変えます
                    rgb[0][y, x] = (byte)(Math.Pow(lrgb[0], 1.0 / gamma) * 255);
                    rgb[1][y, x] = (byte)(Math.Pow(lrgb[1], 1.0 / gamma) * 255);
                    rgb[2][y, x] = (byte)(Math.Pow(lrgb[2], 1.0 / gamma) * 255);

                    //Bitmapに変換するためのbyte配列に入れます
                    byteArray[index++] = (byte)rgb[2][y, x];
                    byteArray[index++] = (byte)rgb[1][y, x];
                    byteArray[index++] = (byte)rgb[0][y, x];
                    byteArray[index++] = 0xFF;  //アルファ
                }
            }

            //自動ホワイトバランスを計算したrgbをBitmap画像に戻します
            // ビットマップを定義
            bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);

            // ビットマップデータを操作するために、マップされているアンマネージドメモリ領域をロックします。
            BitmapData bitmapData = bmp.LockBits(
                                 new Rectangle(0, 0, bmp.Width, bmp.Height),
                                 ImageLockMode.WriteOnly, bmp.PixelFormat);

            try
            {
                // 画像メモリ領域にデータをコピーします。
                Marshal.Copy(byteArray, 0, bitmapData.Scan0, byteArray.Length);
            }
            catch
            {
            }
            finally
            {
                // メモリ領域のロックを解放します。
                bmp.UnlockBits(bitmapData);
            }

            //フォルダとファイル名を読み込む
            string fileFolder = Path.GetDirectoryName(imagePath);
            string filename = Path.GetFileNameWithoutExtension(outputFilename);
            ImageFormat format = ImageFormat.Png;

            //保存します
            try
            {
                //元画像の形式を調べます
                using (Image image = Image.FromFile(imagePath))
                {
                    //コーデックを取得します
                    ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders().FirstOrDefault(c => c.FormatID == image.RawFormat.Guid);
                    //保存するファイル形式と拡張子をセットします
                    if (codec != null)
                    {
                        switch (codec.MimeType)
                        {
                            case "image/jpeg":
                                filename += ".jpg";
                                format = ImageFormat.Jpeg;
                                break;

                            case "image/png":
                                filename += ".png";
                                format = ImageFormat.Png;
                                break;

                            case "image/gif":
                                filename += ".gif";
                                format = ImageFormat.Gif;
                                break;

                            case "image/bmp":
                                filename += ".bmp";
                                format = ImageFormat.Bmp;
                                break;

                            case "image/tiff":
                                filename += ".tif";
                                format = ImageFormat.Tiff;
                                break;

                            default:
                                filename += ".png";
                                format = ImageFormat.Png;
                                break;
                        }
                    }
                    else
                    {
                        filename += ".png";
                        format = ImageFormat.Png;
                    }

                    //ファイルを保存します
                    bmp.Save(fileFolder + "\\" + filename, format);
                    Console.WriteLine("Image saved successfully in the same format as the input.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            bmp?.Dispose();
        }

        /// <summary>
        /// 説明
        /// </summary>
        static void ShowUsage()
        {
            Console.WriteLine("Usage: AutoWhiteBalance.exe --whitepoint <x,y> --gamma <value> --output <filename> <image_path>");
            Console.WriteLine("       AutoWhiteBalance.exe -w <x,y> -g <value> -o <filename> <image_path>");
            Console.WriteLine("       AutoWhiteBalance.exe <image_path> (uses default white point of 0.3127,0.3290, gamma of 2.2, and default output naming)");
            Console.WriteLine("Options:");
            Console.WriteLine("  -w, --whitepoint    Set the chromaticity coordinates of the white point as x,y.");
            Console.WriteLine("  -g, --gamma         Set the gamma value.");
            Console.WriteLine("  -o, --output        Specify the output filename.");
            Console.WriteLine("  <image_path>        Path to the image file.");
        }

        /// <summary>
        /// 逆行列の計算
        /// </summary>
        /// <param name="matrix"></param>
        /// <returns></returns>
        static double[,] InvertMatrix(double[,] matrix)
        {
            int n = matrix.GetLength(0);
            double[,] result = new double[n, n];
            double[,] temp = new double[n, 2 * n];

            // Initialize the matrix extension with identity matrix
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    temp[i, j] = matrix[i, j];
                    if (i == j)
                        temp[i, j + n] = 1;
                    else
                        temp[i, j + n] = 0;
                }
            }

            // Perform row operations to transform the matrix into the identity matrix
            for (int i = 0; i < n; i++)
            {
                // Make the diagonal element 1 and zero out the rest in the column
                double t = temp[i, i];
                for (int j = 0; j < 2 * n; j++)
                {
                    temp[i, j] /= t;
                }
                for (int k = 0; k < n; k++)
                {
                    if (k != i)
                    {
                        t = temp[k, i];
                        for (int j = 0; j < 2 * n; j++)
                        {
                            temp[k, j] -= t * temp[i, j];
                        }
                    }
                }
            }

            // Extract the inverse matrix
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    result[i, j] = temp[i, j + n];
                }
            }

            return result;
        }
    }
}
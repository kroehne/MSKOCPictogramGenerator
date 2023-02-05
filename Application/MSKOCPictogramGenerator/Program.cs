using System;
using System.IO;
using System.Collections.Generic;
using ExcelDataReader;
using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Drawing;
using System.Linq;

namespace PictogramGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("MSK-OC: PictogramGenerator ({0})\n", typeof(Program).Assembly.GetName().Version.ToString());
            Console.ResetColor();
 
            string _defFile = "SourceImages/ImageDefinition.xlsx";

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            if (args.Length < 1 && !File.Exists(_defFile))
            {
                Console.WriteLine("Please provide the (relative) path to an excel-file with texts to render into image(s) as 1st Argument. Expected column names are: ");
                Console.WriteLine(" - Filename: Name of the generated image (must be identical for all rows that belong to one image)");
                Console.WriteLine(" - Width: Pixel in the original metric (must be identical for all rows that belong to one image)");
                Console.WriteLine(" - Height: Pixel in the original metric (must be identical for all rows that belong to one image)");
                Console.WriteLine(" - Scaling: Resize of the image (should be identical for all images!)");
                Console.WriteLine(" - WrapTextWidth: Width of the text used for wrapping in the original metric");
                Console.WriteLine(" - ImageY: Y position of the image in this line in the original metric");
                Console.WriteLine(" - Image: File name of a 200x200 pixel image in the 'source folder'");
                Console.WriteLine(" - TextY: Y position of the text label for the image in this line in the original metric");
                Console.WriteLine(" - Text: Text as label for the image in this line");
                Console.WriteLine("Optional 2nd argument is the relative path to the output folder ('GeneratedImages' as default)");
                Console.WriteLine("Optional 3nd argument is the relative path to the output folder ('SourceImages' as default)");

            }
            else
            {                        
                string _outFolder = @"GeneratedImages\";
                string _sourceFolder = @"SourceImages\";
                if (args.Length > 0)
                {
                    _defFile = args[0];
                }
                if (args.Length > 1)
                {
                    _outFolder = args[1];
                }
                if (args.Length > 2)
                {
                    _sourceFolder = args[2];
                }

                string _configBackgroundColor = "255;255;255";
                var _configBlueColor = "68;114;196";
                var _configTextColor = "0;0;0";
                int _configBorderStartX = 240;
                int _configBorderStartY = 180;
                int _configBorderBorderRight = 10;
                int _configBorderBorderBottom = 10;

                Dictionary<string, LegendDefinition> data = new Dictionary<string, LegendDefinition>();
                 
                try
                {
                    if (File.Exists(_defFile))
                    {
                        Dictionary<string, int> _columnOrder = new Dictionary<string, int>();
                        using (var stream = File.Open(_defFile, FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                bool isHeader = true;
                                do
                                {
                                    while (reader.Read())
                                    { 
                                        if (isHeader)
                                        {
                                            for (int i = 0; i < reader.FieldCount; i++)
                                            {
                                                string _key = reader.GetString(i);
                                                if (_columnOrder.ContainsKey(_key))
                                                    _columnOrder[_key] = i;
                                                else
                                                {
                                                    _columnOrder.Add(_key, i);
                                                }
                                            } 
                                            isHeader = false;
                                        } 
                                        else 
                                        {
                                            try
                                            {

                                                // General

                                                string _currentImageFile = "example.png";
                                                if (_columnOrder.ContainsKey("Filename"))
                                                    _currentImageFile = reader.GetValue(_columnOrder["Filename"]).ToString();

                                                if (!data.ContainsKey(_currentImageFile))
                                                    data.Add(_currentImageFile, new LegendDefinition() { Filename = _currentImageFile});
                                                 
                                                int _currentImageWidth = 1200;                                                
                                                if (_columnOrder.ContainsKey("Width"))
                                                    int.TryParse(reader.GetValue(_columnOrder["Width"]).ToString(), out _currentImageWidth);
                                                data[_currentImageFile].Width = _currentImageWidth;

                                                int _currentImageHeight = 500;
                                                if (_columnOrder.ContainsKey("Height"))
                                                    int.TryParse(reader.GetValue(_columnOrder["Height"]).ToString(), out _currentImageHeight);
                                                data[_currentImageFile].Height = _currentImageHeight;

                                                float _currentImageScalingFactor = 0.25f;                                                 
                                                if (_columnOrder.ContainsKey("Scaling"))
                                                    float.TryParse(reader.GetValue(_columnOrder["Scaling"]).ToString(), out _currentImageScalingFactor);
                                                data[_currentImageFile].Scaling = _currentImageScalingFactor;

                                                // Line Specific

                                                int _currentLineWrapTextWidht = 800;
                                                if (_columnOrder.ContainsKey("WrapTextWidth"))
                                                    int.TryParse(reader.GetValue(_columnOrder["WrapTextWidth"]).ToString(), out _currentLineWrapTextWidht);

                                                int _currentLineImageY = -1;
                                                if (_columnOrder.ContainsKey("ImageY"))
                                                    int.TryParse(reader.GetValue(_columnOrder["ImageY"]).ToString(), out _currentLineImageY);

                                                int _currentLineTextY = -1;
                                                if (_columnOrder.ContainsKey("TextY"))
                                                    int.TryParse(reader.GetValue(_columnOrder["TextY"]).ToString(), out _currentLineTextY);

                                                string _currentLineImage = "";
                                                if (_columnOrder.ContainsKey("Image"))
                                                    _currentLineImage = reader.GetValue(_columnOrder["Image"]).ToString();
                                                 
                                                string _currentLineText = "";
                                                if (_columnOrder.ContainsKey("Text"))
                                                    _currentLineText = reader.GetValue(_columnOrder["Text"]).ToString();

                                                data[_currentImageFile].Lines.Add(new LegendLine() { Image = _currentLineImage, Text = _currentLineText, ImageY = _currentLineImageY, TextY = _currentLineTextY, WrapTextWidth = _currentLineWrapTextWidht });
                                                 
                                            }
                                            catch (Exception _ex)
                                            {
                                                Console.WriteLine(_ex);
                                            }
                                        }
                                        
                                    }
                                } while (reader.NextResult());
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(_defFile + " not found");
                    }


                    foreach (var v in data.Values)
                    {
                        GenerateImage(_outFolder, _sourceFolder,  _configBackgroundColor, _configBlueColor, _configTextColor, _configBorderStartX, _configBorderStartY, _configBorderBorderRight, _configBorderBorderBottom, v);
                    }

                }
                catch (Exception _ex)
                {
                    Console.WriteLine(_ex);
                }
            }
        }

        private static void GenerateImage(string _outFolder, string _sourceFolder, string _configBackgroundColor, string _configBlueColor, string _configTextColor, int _configBorderStartX, int _configBorderStartY, int _configBorderBorderRight, int _configBorderBorderBottom, LegendDefinition v)
        {
            int _width = _configBorderStartX + v.Width+ _configBorderBorderRight;
            int _height = _configBorderStartY + v.Height + _configBorderBorderBottom;

            using (var image = new Image<Rgba32>(_width, _height))
            {
                var _blueBorderPen = new Pen(GetColor(_configBlueColor), 10);
                var _createdFontTitle = SystemFonts.CreateFont("Calibri", 140, FontStyle.Bold);
                var _createdFontTextEntry = SystemFonts.CreateFont("Calibri", 140, FontStyle.Bold);
                var _textOptionsTitle = new TextOptions(_createdFontTitle)
                {
                    Origin = new PointF(370, 30),
                    KerningMode = KerningMode.Normal, 
                    VerticalAlignment = VerticalAlignment.Top,
                    HorizontalAlignment = SixLabors.Fonts.HorizontalAlignment.Left,
                };


                var _myRectRounded = ApplyRoundCorners(new RectangularPolygon(_configBorderStartX, _configBorderStartY, v.Width, v.Height), 25);

                image.Mutate(ctx => ctx
                         .Fill(GetBrush(_configBackgroundColor, _height, _width))
                         .DrawImage(Image.Load(System.IO.Path.Combine(_sourceFolder, "Zahnraeder_340x400.png")), new Point(10, 10), 1)
                         .Draw(_blueBorderPen, _myRectRounded)
                        .DrawText(_textOptionsTitle, "So geht's:",  Brushes.Solid(GetColor(_configBlueColor)))
                        );

                foreach (var l in v.Lines)
                {
                    var _textOptionsTextEntry = new TextOptions(_createdFontTextEntry)
                    {
                        Origin = new PointF(520, l.TextY+20),
                        KerningMode = KerningMode.Normal, 
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = SixLabors.Fonts.HorizontalAlignment.Left,
                        WrappingLength = l.WrapTextWidth,
                        LineSpacing = (float)0.8,
                       
                    };

                    var _v = Image.Load(System.IO.Path.Combine(_sourceFolder, l.Image));
    
                    image.Mutate(ctx => ctx

                     .DrawImage(_v,new Point(280, l.ImageY-_v.Height/2), 1)
                     .DrawText(_textOptionsTextEntry,  l.Text, Brushes.Solid(GetColor(_configTextColor)))
                    );
                }
                 
                image.Mutate(ctx => ctx
                        .Resize(new Size((int)(_width * v.Scaling), (int)(_height * v.Scaling)))
                );

                Console.WriteLine(v.Filename);

                image.SaveAsPng(System.IO.Path.Combine(_outFolder, v.Filename));

            }
        }

        private static IPath ApplyRoundCorners(RectangularPolygon rectangularPolygon, float radius)
        {
            var squareSize = new SizeF(radius, radius);
            var ellipseSize = new SizeF(radius * 2, radius * 2);
            var offsets = new[]
            {
                (0, 0),
                (1, 0),
                (0, 1),
                (1, 1),
            };

            var holes = offsets.Select(
                offset =>
                {
                    var squarePos = new PointF(
                        offset.Item1 == 0 ? rectangularPolygon.Left : rectangularPolygon.Right - radius,
                        offset.Item2 == 0 ? rectangularPolygon.Top : rectangularPolygon.Bottom - radius
                    );
                    var circlePos = new PointF(
                        offset.Item1 == 0 ? rectangularPolygon.Left + radius : rectangularPolygon.Right - radius,
                        offset.Item2 == 0 ? rectangularPolygon.Top + radius : rectangularPolygon.Bottom - radius
                    );
                    return new RectangularPolygon(squarePos, squareSize)
                        .Clip(new EllipsePolygon(circlePos, ellipseSize));
                }
            );
            return rectangularPolygon.Clip(holes);
        }
 
        private static Rgba32 GetColor(string colorRGB)
        {
            int _R = 0;
            int _G = 0;
            int _B = 0;
            int _A = 0;
            string[] _colorComp = colorRGB.Split(";");
            if (_colorComp.Length != 3 && _colorComp.Length != 4)
            {
                return new Rgba32(_R, _G, _B);
            } 
            else
            {
                int.TryParse(_colorComp[0], out _R);
                int.TryParse(_colorComp[1], out _G);
                int.TryParse(_colorComp[2], out _B);

                if (_colorComp.Length == 3)
                { 
                    return new Rgba32((float)_R/255, (float) _G / 255, (float)_B / 255);
                } 
                else
                {
                    int.TryParse(_colorComp[3], out _A);
                    return new Rgba32((float)_R / 255, (float)_G / 255, (float)_B / 255, (float)_A / 255);
                }
            }
        }

        private static LinearGradientBrush GetBrush(string colorRGB, int height, int width)
        {
            
            var linearGradientBrush = new LinearGradientBrush(new Point(0, 0), new Point(0, height), GradientRepetitionMode.Repeat,
                                            new ColorStop(0, Color.White), new ColorStop(1, Color.White));

            string[] _multipleColors = colorRGB.Split("|");
             
            string[] _colorComp1 = _multipleColors[0].Split(";");
            if (_colorComp1.Length != 3 && _colorComp1.Length != 4)
            {
                return linearGradientBrush;
            }
            else
            {
                ColorStop c1 = new ColorStop();
                ColorStop c2 = new ColorStop();

                byte _R1 = 0;
                byte _G1 = 0;
                byte _B1 = 0;
                byte _A1 = 0;

                byte.TryParse(_colorComp1[0], out _R1);
                byte.TryParse(_colorComp1[1], out _G1);
                byte.TryParse(_colorComp1[2], out _B1);

                if (_colorComp1.Length == 3)
                {
                    c1 = new ColorStop(0, Color.FromRgb(_R1, _G1, _B1));
                    c2 = new ColorStop(1, Color.FromRgb(_R1, _G1, _B1));
                }
                else
                {
                    byte.TryParse(_colorComp1[3], out _A1);
                    c1 = new ColorStop(0, Color.FromRgba(_R1, _G1, _B1, _A1));
                    c2 = new ColorStop(1, Color.FromRgba(_R1, _G1, _B1, _A1));
                }

                if (_multipleColors.Length == 2)
                {
                    string[] _colorComp2 = _multipleColors[1].Split(";");

                    byte _R2 = 0;
                    byte _G2 = 0;
                    byte _B2 = 0;
                    byte _A2 = 0;

                    byte.TryParse(_colorComp2[0], out _R2);
                    byte.TryParse(_colorComp2[1], out _G2);
                    byte.TryParse(_colorComp2[2], out _B2);

                    if (_colorComp2.Length == 3)
                    {
                        c2 = new ColorStop(1, Color.FromRgb(_R2, _G2, _B2));
                    }
                    else
                    {
                        byte.TryParse(_colorComp2[3], out _A2);
                        c2 = new ColorStop(1, Color.FromRgba(_R2, _G2, _B2, _A2));
                    }

                }
                  
                return new LinearGradientBrush(new Point(0, 0), new Point(0, height), GradientRepetitionMode.Repeat, c1, c2);

            }
        }

    } 


    class LegendDefinition
    {
        public string Filename { get; set; }
        public int Width { get; set; }
        public int Height{ get; set; }        
        public float Scaling{ get; set; }

        public List<LegendLine> Lines { get; set; } = new List<LegendLine>();
    }
    class LegendLine
    {
        public string Text { get; set; }
        public string Image { get; set; }
        public int TextY { get; set; }
        public int ImageY { get; set; }
        public int WrapTextWidth { get; set; }

    }
}

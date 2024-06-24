using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using System.IO;
using System.Windows;

namespace TTN
{
    internal class Filter
    {
        bool ochist = true;
        public Filter(bool ochist_) 
        { 
            ochist = ochist_;
        }
        public void FilterWhite(string path1, string outpath)
        {
            using (Image<Rgba32> image = Image.Load<Rgba32>(path1))
            {
                if (ochist == true)
                {
                    image.Mutate(ctx =>
                    {
                        ctx.Contrast(2.0f);
                        ctx.Brightness(1.5f);
                        ctx.Brightness(0.8f);
                        ctx.Saturate(0f);
                        ctx.Brightness(1.2f);
                        ctx.GaussianSharpen(1.5f);
                        ctx.GaussianBlur(0.5f);
                        ctx.Contrast(1.3f);
                    });
                }
                image.Save(Path.Combine(outpath, $"doc1.png"));
                MessageBox.Show("???");
                MessageBox.Show(Path.Combine(outpath, $"doc1.png"));
            }
        }
    }
}

using System.Drawing;

namespace DER.Utility.Helpers
{
    public static class ColorHelpers
    {
        public static string ApplyTint(this Color color, double tint)
        {
            if (tint < -1.0) tint = -1.0;
            if (tint > 1.0) tint = 1.0;

            Color colorRgb = color;
            double fHue = colorRgb.GetHue();
            double fSat = colorRgb.GetSaturation();
            double fLum = colorRgb.GetBrightness();
            if (tint < 0)
            {
                fLum = fLum * (1.0 + tint);
            }
            else
            {
                fLum = fLum * (1.0 - tint) + (1.0 - 1.0 * (1.0 - tint));
            }
            return $"{color.A:X2}{ToHexColor(fHue, fSat, fLum)}";
        }

        private static string ToHexColor(double hue, double saturation, double luminance)
        {
            double chroma = (1.0 - Math.Abs(2.0 * luminance - 1.0)) * saturation;
            double fHue = hue / 60.0;
            double fHueMod2 = fHue;
            while (fHueMod2 >= 2.0) fHueMod2 -= 2.0;
            double fTemp = chroma * (1.0 - Math.Abs(fHueMod2 - 1.0));

            double fRed, fGreen, fBlue;
            if (fHue < 1.0)
            {
                fRed = chroma;
                fGreen = fTemp;
                fBlue = 0;
            }
            else if (fHue < 2.0)
            {
                fRed = fTemp;
                fGreen = chroma;
                fBlue = 0;
            }
            else if (fHue < 3.0)
            {
                fRed = 0;
                fGreen = chroma;
                fBlue = fTemp;
            }
            else if (fHue < 4.0)
            {
                fRed = 0;
                fGreen = fTemp;
                fBlue = chroma;
            }
            else if (fHue < 5.0)
            {
                fRed = fTemp;
                fGreen = 0;
                fBlue = chroma;
            }
            else if (fHue < 6.0)
            {
                fRed = chroma;
                fGreen = 0;
                fBlue = fTemp;
            }
            else
            {
                fRed = 0;
                fGreen = 0;
                fBlue = 0;
            }

            double fMin = luminance - 0.5 * chroma;
            fRed += fMin;
            fGreen += fMin;
            fBlue += fMin;

            fRed *= 255.0;
            fGreen *= 255.0;
            fBlue *= 255.0;

            var red = Convert.ToInt32(Math.Truncate(fRed));
            var green = Convert.ToInt32(Math.Truncate(fGreen));
            var blue = Convert.ToInt32(Math.Truncate(fBlue));

            red = Math.Min(255, Math.Max(red, 0));
            green = Math.Min(255, Math.Max(green, 0));
            blue = Math.Min(255, Math.Max(blue, 0));

            return $"{red:X2}{green:X2}{blue:X2}";
        }
    }
}

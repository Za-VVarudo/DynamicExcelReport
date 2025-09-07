using DocumentFormat.OpenXml.Drawing;

namespace DER.Utility.Extensions
{
    public static class OpenXmlExtensions
    {
        public static string ToRGBA(this RgbColorModelHex hex, int alpha = 255)
        {
            if (hex == null || !hex.Val.TryGetBytes(out var rgb)) return null;
            return $"{Math.Clamp(alpha, 0, 255):X2}{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
        }
    }
}

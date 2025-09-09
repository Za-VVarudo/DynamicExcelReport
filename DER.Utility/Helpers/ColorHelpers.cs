namespace DER.Utility.Helpers
{
    public static class ColorHelpers
    {
        public static string ToRGBA(int a, int r, int g, int b, double tint = 0)
        {
            return $"{a:X2}{ApplyTint(r, tint):X2}{ApplyTint(g, tint):X2}{ApplyTint(b, tint):X2}";
        }

        private static int ApplyTint(int channel, double tint)
        {
            return (int)(tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint);
        }
    }
}

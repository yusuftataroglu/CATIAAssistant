namespace CATIAAssistant.Helpers
{
    public class ParseQuantityHelper
    {
        public (int drawn, int mirror) ParseDrawnMirror(string cellValue)
        {
            if (string.IsNullOrWhiteSpace(cellValue))
                return (0, 0);
            if (!cellValue.Contains("x", StringComparison.OrdinalIgnoreCase))
                return (0, 0);
            var parts = cellValue.Split('/');
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].Trim();
            }
            int drawn = 0, mirror = 0;

            // "2x" → drawn = 2
            if (parts.Length > 0)
                drawn = ParseQuantity(parts[0]);

            // "3x" → mirror = 3
            if (parts.Length > 1)
                mirror = ParseQuantity(parts[1]);

            return (drawn, mirror);
        }

        public int ParseQuantity(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return 0;

            if (input.EndsWith("x", StringComparison.OrdinalIgnoreCase))
                input = input.Substring(0, input.Length - 1);

            if (int.TryParse(input, out int result))
                return result;

            return 0;
        }
    }
}

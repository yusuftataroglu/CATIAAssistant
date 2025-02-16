namespace CATIAAssistant.Helpers
{
    public class ParseQuantityHelper
    {
        public int ParseQuantity(string input)
        {
            input = input.Trim();
            if (string.IsNullOrWhiteSpace(input))
                return 0;

            // Metin "2x" gibi bitiyorsa 'x' karakterini kaldır.
            if (input.EndsWith("x", StringComparison.OrdinalIgnoreCase))
            {
                // Substring ile son karakteri atıyoruz
                input = input.Substring(0, input.Length - 1);
            }

            // Kalan metni sayıya dönüştürmeye çalış
            if (int.TryParse(input, out int result))
                return result;

            return 0; // parse edilemezse 0 döndürüyoruz
        }
    }
}

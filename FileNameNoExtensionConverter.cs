using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace CocktailViewer
{
    public sealed class FileNameNoExtensionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not string s || string.IsNullOrWhiteSpace(s))
                return string.Empty;

            try
            {
                return Path.GetFileNameWithoutExtension(s);
            }
            catch
            {
                return s;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
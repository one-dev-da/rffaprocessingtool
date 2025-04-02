using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace RffaDataComparisonTool
{
    /// <summary>
    /// Converter to transform a Dictionary into a collection of KeyValuePair items for display
    /// </summary>
    public class DictionaryItemsConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length > 0 && values[0] is Dictionary<string, List<string>> dictionary)
            {
                return dictionary.Select(kvp => new KeyValuePair<string, List<string>>(kvp.Key, kvp.Value));
            }

            return null;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
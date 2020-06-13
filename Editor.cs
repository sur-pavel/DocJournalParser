using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace DocJournalParser
{
    public class Editor : Autor
    {
        public string Function { get; set; } = string.Empty;

        private PropertyInfo[] _PropertyInfos = null;

        public override string ToString()
        {
            if (_PropertyInfos == null)
                _PropertyInfos = GetType().GetProperties();

            var builder = new StringBuilder();

            foreach (var info in _PropertyInfos)
            {
                var value = info.GetValue(this, null) ?? "(null)";
                builder.AppendLine(info.Name + ": " + value.ToString());
            }

            return builder.ToString();
        }

        public override bool Equals(object obj)
        {
            return obj is Editor editor &&
                   base.Equals(obj) &&
                   LastName == editor.LastName &&
                   Initials == editor.Initials &&
                   Rank == editor.Rank &&
                   Invertion == editor.Invertion &&
                   Function == editor.Function;
        }

        public override int GetHashCode()
        {
            int hashCode = 2083592609;
            hashCode = hashCode * -1521134295 + base.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(LastName);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Initials);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Rank);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Invertion);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Function);
            return hashCode;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocJournalParser
{
    public class Autor
    {
        public string LastName { get; set; } = string.Empty;
        public string Initials { get; set; } = string.Empty;
        public string Rank { get; set; } = string.Empty;
        public string Invertion { get; set; } = string.Empty;
        public string LNameVariation { get; set; }= string.Empty;

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
            return obj is Autor autor &&
                   LastName == autor.LastName &&
                   Initials == autor.Initials &&
                   Rank == autor.Rank &&
                   Invertion == autor.Invertion;
        }

        public override int GetHashCode()
        {
            int hashCode = -1373312750;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(LastName);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Initials);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Rank);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Invertion);
            return hashCode;
        }
    }
}
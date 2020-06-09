using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocJournalParser
{
    public class JDiscription
    {
        public string LastName { get; set; } = string.Empty;
        public string Initials { get; set; } = string.Empty;
        public string Rank { get; set; } = string.Empty;
        public string Invertion { get; set; } = string.Empty;
        public string FirstEdLastName { get; set; } = string.Empty;
        public string FirstEdInitials { get; set; } = string.Empty;
        public string FirstEdRank { get; set; } = string.Empty;
        public string FirstEdInvertion { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string TitleInfo { get; set; } = string.Empty;
        public string Editors { get; set; } = string.Empty;
        public string Year { get; set; } = string.Empty;
        public string JVolume { get; set; } = string.Empty;
        public string JNumber { get; set; } = string.Empty;
        public string Pages { get; set; } = string.Empty;
        public string Notes { get; set; } = string.Empty;
        public string FullPubYear { get; set; } = string.Empty;
        public string FullPubVolume { get; set; } = string.Empty;
        public string FullPubNumber { get; set; } = string.Empty;
        public string FullPubPageRange { get; set; } = string.Empty;
        public dynamic DiscriptionNumber { get; set; } = string.Empty;

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
            return obj is JDiscription discription &&
                   LastName == discription.LastName &&
                   Initials == discription.Initials &&
                   Rank == discription.Rank &&
                   Invertion == discription.Invertion &&
                   Title == discription.Title &&
                   TitleInfo == discription.TitleInfo &&
                   Year == discription.Year &&
                   JVolume == discription.JVolume &&
                   JNumber == discription.JNumber &&
                   Pages == discription.Pages &&
                   Notes == discription.Notes &&
                   FullPubYear == discription.FullPubYear &&
                   FullPubVolume == discription.FullPubVolume &&
                   FullPubNumber == discription.FullPubNumber &&
                   FullPubPageRange == discription.FullPubPageRange;
        }

        public override int GetHashCode()
        {
            int hashCode = -803217302;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(LastName);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Initials);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Rank);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Invertion);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Title);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(TitleInfo);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Year);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(JVolume);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(JNumber);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Pages);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Notes);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FullPubYear);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FullPubVolume);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FullPubNumber);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FullPubPageRange);
            return hashCode;
        }
    }
}
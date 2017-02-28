using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Castle.ActiveRecord;

namespace XiZhi.WordRepeat.Model
{
    [ActiveRecord("wordRES")]
    class WordRES : ActiveRecordBase<WordRES>
    {
        [PrimaryKey(PrimaryKeyType.Identity, "num")]
        public int num { get; set; }

        [Property("word")]
        public String word { get; set; }


        [Property("subject")]
        public int subject { get; set; }

        [Property("phoneticSymbol")]
        public String phoneticSymbol { get; set; }

        [Property("wordMeaning")]
        public String wordMeaning { get; set; }

        [Property("picture")]
        public String picture { get; set; }

        [Property("sound")]
        public String sound { get; set; }

        [Property("unit")]
        public int unit { get; set; }

        [Property("book")]
        public int book { get; set; }

        [Property("fontType")]
        public int fontType { get; set; }

        public static WordRES Find(int id)
        {
            return (WordRES)FindByPrimaryKey(typeof(WordRES), id);
        }
    }
}

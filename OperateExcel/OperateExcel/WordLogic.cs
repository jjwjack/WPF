using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Castle.ActiveRecord;
using System.Collections;

namespace OperateExcel
{
    [ActiveRecord("wordLogic")]
    public class WordLogic : ActiveRecordBase<WordLogic>
    {
        [PrimaryKey(PrimaryKeyType.Identity, "num")]
        public int num { get; set; }

        [Property("word")]
        public String word { get; set; }

        [Property("subject")]
        public int subject { get; set; }

        [Property("remCount")]
        public int remCount { get; set; }

        [Property("lastRemTime")]
        public DateTime lastRemTime { get; set; }

        [Property("nextRemTime")]
        public DateTime nextRemTime { get; set; }

        [Property("unit")]
        public int unit { get; set; }

        [Property("book")]
        public int book { get; set; }

        [HasMany(typeof(WordRES), Table = "wordRES", ColumnKey = "word")]
        public IList Data { get; set; }
    }
}

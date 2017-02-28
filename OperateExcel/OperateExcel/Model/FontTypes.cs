using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Castle.ActiveRecord;

namespace XiZhi.WordRepeat.Model
{
    [ActiveRecord("fontTypes")]
    class FontTypes : WordReader
    {
        [PrimaryKey(PrimaryKeyType.Identity, "fontName")]
        public String fontName { get; set; }

        [Property("fontNum")]
        public int fontNum { get; set; }
    }
}

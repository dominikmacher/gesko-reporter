using System;
using System.Collections.Generic;
using System.Text;

namespace GeskoReporter
{
    public class Column
    {
        public int index;
        public string name;

        public Column(int index, string name)
        {
            this.index = index;
            this.name = name;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
namespace BussinessLogicLayer
{
    class Logger 
    {
        public enum Operations
        { 
            ArchiveDocument,
            CreateBinder
        };
        public enum Statuses
        {
            OK,
            Error,
            MovedToManual
        };
        Logger(){}
        public static void Log(int doc_type_id, Operations op, Statuses st, int doc_id, string file_name, string txt, string title, string text)
        {
        }

    }
}

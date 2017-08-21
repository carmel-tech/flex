using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BussinessLogicLayer
{
    class DOXFields
    {
        public static DOXAPI.Field newField(string name, DOXAPI.DocTypeAttribute att, object value)
        {
            DOXAPI.Field f = new DOXAPI.Field();
            f.Attr = att;
            f.Attr.Name = name;
            f.Value = value;
            return f;
        }

        public static object GetField(DOXAPI.TreeItemWithDocType entity, string fname)
        {
            foreach (DOXAPI.Field f in entity.Fields) if (f.Attr.Name == fname) return f.Value;
            return null;
        }

        public static bool SetField(DOXAPI.TreeItemWithDocType entity, string fname, object value)
        {
            if (entity.Fields == null) return false;
            foreach (DOXAPI.Field f in entity.Fields) if (f.Attr.Name == fname)
                {
                    f.Value = value;
                    return true;
                }
            return false;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class Component
    {
        struct s_COMPONENT
        {
            public String _ref;
            public String _value;
            public String _refDigikey;
        };

        struct s_listing
        {
            public List<String> _refPart;
            public String _refDigikey;
            public String _val;
        };

        List<s_listing> lis = new List<s_listing>();

        List<s_COMPONENT> compList = new List<s_COMPONENT>();

        s_COMPONENT compTemp;
        Boolean isEmpty = true;
        static Excel excel = new Excel();

        public void create()
        {
            compTemp = new s_COMPONENT();
            isEmpty = false;
        }

        public void setRef(String val)
        {
            String tmp;
            if (isEmpty == false)
            {
                tmp = val.Replace("(comp (ref ", "").Trim();
                while (tmp.EndsWith(")"))
                    tmp = tmp.Substring(0, tmp.Length - 1);

                compTemp._ref = tmp;
            }
        }

        public void setValue(String val)
        {
            String tmp;
            if (isEmpty == false)
            {
                tmp = val.Replace("\"", "");
                tmp = tmp.Replace("(value ", "").Trim();
                while (tmp.EndsWith(")"))
                    tmp = tmp.Substring(0, tmp.Length - 1);

                compTemp._value = tmp;
            }
        }

        public void setRefDigikey(String val)
        {
            String tmp;
            if (isEmpty == false)
            {
                tmp = val.Replace("(field (name \"Ref Digikey\")", "").Trim();
                while (tmp.EndsWith(")"))
                    tmp = tmp.Substring(0, tmp.Length - 1);

                compTemp._refDigikey = tmp;
            }
        }

        public void save()
        {
            if ( isEmpty == false )
                compList.Add(compTemp);

            this.create();
        }

        public void createList()
        {
            while (true)
            {
                foreach (s_COMPONENT comp in compList)
                {
                    string aRechercher = comp._refDigikey;
                    List<s_COMPONENT> lcd = compList.FindAll(delegate(s_COMPONENT t)
                    {
                        return t._refDigikey == aRechercher;
                    });
                    aRechercher = comp._ref.Substring(0, 1);
                    List<s_COMPONENT> lce = lcd.FindAll(delegate(s_COMPONENT t)
                    {
                        return t._ref.Contains(aRechercher);
                    });

                    aRechercher = comp._value;
                    List<s_COMPONENT> lc = lce.FindAll(delegate(s_COMPONENT t)
                    {
                        return t._value.Equals(aRechercher);
                    });


                    s_listing lf = new s_listing();
                    lf._refDigikey = comp._refDigikey;
                    lf._val = comp._value;
                    lf._refPart = new List<String>();
                    foreach (s_COMPONENT c in lc)
                    {
                        lf._refPart.Add(c._ref);
                        compList.Remove(c);
                    }

                    lis.Add(lf);

                    break;
                }

                if (compList.Count == 0)
                {
                    break;
                }
            }
        }

        public void createBOM()
        {
            excel.create();

            foreach (s_listing lf in lis)
            {
                excel.newLine();

                foreach( String s in lf._refPart )
                {
                    excel.setRefClient(s);
                }

                excel.setValue(lf._val);

                excel.setRefDigikey(lf._refDigikey);

                excel.setQuantity(lf._refPart.Count.ToString());
            }

            excel.close();
        }

    }
}

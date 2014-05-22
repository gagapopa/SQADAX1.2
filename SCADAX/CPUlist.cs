using System;
using System.Collections.Generic;
using System.Linq;

namespace SCADAX
{
    [Serializable]
        public class Signal
        {
            public string KKS { get; set; }
            public string RealKKS { get; set; }
            public string Marka { get; set; }
            public string KMS { get; set; }
            public string Cabel { get; set; }
            public string Connect { get; set; }
            private string _key;
            public string Key
            {
                get { return _key; }
                set
                {
                    value = _key;
                    int i = 0;
                }
            }

            public List<string> ZlkforkeyList; //для составления ключа
           
            public List<string> Zhilka { get; set; }
      
            public Signal(string _kks = "", string _marka = "", string _kms = "", string _con = null, string _c = "", string _realkks = "", List<string> _zlk = null, List<string> _zlkforkey = null)
            {
                KKS = _kks.Trim('_');
                RealKKS = _realkks;
                Marka = _marka;
                KMS = _kms;
                Cabel = _c;
                Connect = _con;
                Zhilka = new List<string>();
                Zhilka = _zlk;
                ZlkforkeyList = new List<string>();
                ZlkforkeyList = _zlkforkey;
                if (KKS.Contains('_'))
                {
                    if (ZlkforkeyList != null)
                        _key = Connect + "#" + ZlkforkeyList.Aggregate("", (a, b) => a + b + ",").Trim(',') + "#" + KKS.Substring(KKS.LastIndexOf('_') + 1).Trim();
                }
                else
                {
                    if (ZlkforkeyList != null)
                        _key = Connect + "#" + ZlkforkeyList.Aggregate("", (a, b) => a + b + ",").Trim(',') + "##";
                }
                int i = 0;
            }
        }
    [Serializable]
        public class Module
        {
            public List<Signal> Signals { get; set; }
            public string Type { get; set; }

            public Module(string _type = "")
            {
                Type = _type;
                Signals = new List<Signal>(50);
                for (int i = 0; i < Signals.Capacity; i++)
                {
                    Signals.Add(new Signal());
                };
            }

        }
    [Serializable]
        public class CPU
        {
            public List<Module> modules { get; set; }
            public string Name { get; set; }

            public CPU(string _name = "")
            {
                Name = _name;
                modules = new List<Module>(70);
                for (int i = 0; i < modules.Capacity; i++)
                {
                    modules.Add(new Module(""));
                };
            }
        }

}

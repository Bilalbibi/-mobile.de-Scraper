using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mobile.de_Scraper.Models
{
    public class Make
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public List<Model> Models { get; set; }
        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object obj)
        {
            if (obj.GetType() == typeof(Make))
                return Id.Equals(((Make)obj).Id);
            return base.Equals(obj);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAllocatev2
{
    [Serializable]
    public class Requirement
    {
        public string Name { get; set; }
        public int Priority { get; set; }
        public int Id { get; set; }
        public int DependsOnId { get; set; }
        public List<Dependency> DependencyList = new List<Dependency>();
        public Dictionary<string,int> EffortByType = new Dictionary<string,int>();
        public Dictionary<string, int> LongPoleByType = new Dictionary<string, int>();
        public Dictionary<string, int> MaximimResourcesByType = new Dictionary<string, int>();

        public override string ToString()
        {
            string value = string.Format("Name = [{0}],Priority=[{1}] EffortByType =[{2}], LongPoleByType = [{3}]," +
            "MaximimResourcesByType=[{4}],Dependency=[{5}]", Name, Priority, DictionaryToString(EffortByType), DictionaryToString(LongPoleByType), DictionaryToString(MaximimResourcesByType), String.Join(",",DependencyList));
            return value;

        }


        private static string DictionaryToString(Dictionary<string, int>  dictionary)
        {
            string textValue = "";
            foreach (KeyValuePair<string, int> kvp in dictionary)
            {                
                textValue += string.Format("Key = {0}, Value = {1}", kvp.Key, kvp.Value);
            }
            return textValue;
        }
    }

   [Serializable]
   public class Dependency
   {
       public string Depender { get; set;}
       public string Dependent { get; set; }
       public int Days { get; set; }
       public override string ToString()
       {
           return string.Format("Depender = {0},Dependent = {1},Days={2}", Depender, Dependent, Days);
       }
   }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppSrez2.models
{

        public class Client
        {
            public string lastName { get; set; }
            public string firstName { get; set; }
            public string patronymic { get; set; }

            public string FIO
            {
                get
                {
                    char a = firstName.FirstOrDefault();
                    char b = patronymic.FirstOrDefault();
                    return lastName + " " + a + ". " + b + ".";
                }
            }
        }
    
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace convertExcel.Models
{
    public class HashData
    {
        public object data { get; set; }
         public int result {  get; set; }
        public string message { get; set; }
    }
    public class hashdata1
    {
        public List<Account1> data { get; set; }
        public hashdata1()
        {
            data = new List<Account1>();
            
        }
        
    }
   
    public class hash1
    {
        public List<product> data { get; set; }
        public hash1()
        {
            
            data = new List<product>();
        }

    }
    public class Account1
    {
        public string iAccountType { get; set; }
        
        public string scode { get; set; }
        public string sname { get; set; }

    }

    public class product

    {
        public string scode { get; set; }
        public string sname { get; set; }
    }
     

   
    public class loginData
    {
        public string userName { get; set; }
        public string password { get; set; }
        public string CompanyId { get; set; }
        public string fSessionId { get; set; }
    }
    public class Data
    {
        public List<loginData> data { get; set; }

        public Data()
        {
            data = new List<loginData>();
        }

    }

}
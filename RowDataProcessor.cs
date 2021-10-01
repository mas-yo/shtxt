using System;
using System.Collections.Generic;

namespace MasterDataConverter
{
    public class RowDataProcessor : IRowDataProcessor
    {
        public void Process(IList<string> rowdata)
        {
            foreach (var data in rowdata)
            {
                Console.Write(data);
                Console.Write(", ");
            }

            Console.WriteLine("");
            // throw new System.NotImplementedException();
        }
    }
}
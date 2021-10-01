using System;
using System.Collections.Generic;

namespace MasterDataConverter
{
    public interface IRowDataProcessor
    {
        void Process(IList<String> rowdata);
    }
}